# Copyright (c) 2025 AmyLin <zhi_lin@qq.com>
# Licensed under the MIT License. See LICENSE file for details.

"""Flask web service for doc2md converter.

Provides a web UI for uploading .docx files, previewing the
converted Markdown, and downloading the result — a single .md file when
no extra resources exist, or a .zip when images or multiple outputs are
involved.
"""

from __future__ import annotations

import atexit
import io
import logging
import os
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template

from converter.word2md import convert_word_to_markdown

logger = logging.getLogger(__name__)

app = Flask(
    __name__,
    template_folder=os.path.join(os.path.dirname(__file__), "..", "templates"),
    static_folder=os.path.join(os.path.dirname(__file__), "..", "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB upload limit

ALLOWED_EXTENSIONS = {".docx"}

# ── Temporary directory configuration ──
# Default: store in project directory for easy testing
PROJECT_ROOT = Path(__file__).parent.parent
UPLOADS_DIR = PROJECT_ROOT / "uploads"      # User uploaded files
CONVERTED_DIR = PROJECT_ROOT / "converted"  # Generated markdown and images

# Create directories
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
CONVERTED_DIR.mkdir(parents=True, exist_ok=True)

logger.info(f"Upload directory: {UPLOADS_DIR}")
logger.info(f"Converted directory: {CONVERTED_DIR}")

# Allow override via environment variable
_CUSTOM_UPLOAD_DIR = os.environ.get("DOC2MD_UPLOAD_DIR")
_CUSTOM_CONVERTED_DIR = os.environ.get("DOC2MD_CONVERTED_DIR")
if _CUSTOM_UPLOAD_DIR:
    UPLOADS_DIR = Path(_CUSTOM_UPLOAD_DIR)
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    logger.info(f"Using custom upload directory: {UPLOADS_DIR}")
if _CUSTOM_CONVERTED_DIR:
    CONVERTED_DIR = Path(_CUSTOM_CONVERTED_DIR)
    CONVERTED_DIR.mkdir(parents=True, exist_ok=True)
    logger.info(f"Using custom converted directory: {CONVERTED_DIR}")

# ── Temporary result storage ──
# { result_id: { "work_dir": Path, "files": [...], "needs_zip": bool, "created": float } }
_results: dict[str, dict] = {}
_results_lock = threading.Lock()
_RESULT_TTL = 600  # 10 minutes before auto-cleanup
_cleanup_thread = None
_cleanup_stop_event = threading.Event()


def _cleanup_expired():
    """Remove results older than _RESULT_TTL seconds."""
    now = time.time()
    cleaned_count = 0
    
    with _results_lock:
        expired = [k for k, v in _results.items() if now - v["created"] > _RESULT_TTL]
        for k in expired:
            try:
                upload_dir = _results[k].get("upload_dir")
                converted_dir = _results[k].get("converted_dir")
                
                if upload_dir and Path(upload_dir).exists():
                    shutil.rmtree(upload_dir, ignore_errors=True)
                    logger.info(f"Cleaned up upload dir: {upload_dir}")
                
                if converted_dir and Path(converted_dir).exists():
                    shutil.rmtree(converted_dir, ignore_errors=True)
                    logger.info(f"Cleaned up converted dir: {converted_dir}")
                
                del _results[k]
                cleaned_count += 1
            except Exception as e:
                logger.warning(f"Failed to clean up result {k}: {e}")
    
    return cleaned_count


def _background_cleanup():
    """Background thread to clean up expired results every 60 seconds."""
    logger.info("Background cleanup thread started")
    while not _cleanup_stop_event.is_set():
        try:
            count = _cleanup_expired()
            if count > 0:
                logger.info(f"Background cleanup removed {count} expired result(s)")
        except Exception as e:
            logger.error(f"Background cleanup error: {e}")
        # Wait 60 seconds or until stop event
        _cleanup_stop_event.wait(60)
    logger.info("Background cleanup thread stopped")


def _start_cleanup_thread():
    """Start the background cleanup thread."""
    global _cleanup_thread
    if _cleanup_thread is None or not _cleanup_thread.is_alive():
        _cleanup_stop_event.clear()
        _cleanup_thread = threading.Thread(target=_background_cleanup, daemon=True)
        _cleanup_thread.start()


def _stop_cleanup_thread():
    """Stop the background cleanup thread."""
    _cleanup_stop_event.set()
    if _cleanup_thread:
        _cleanup_thread.join(timeout=5)


# Start cleanup thread on module load
_start_cleanup_thread()

# Ensure cleanup on exit
atexit.register(_stop_cleanup_thread)


def _allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    """Render the upload page."""
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    """Handle file upload and conversion.

    Returns JSON with preview content and a download ID:
      {
        id: "...",
        files: [ { name: "xxx.md", content: "..." } ],
        needs_zip: true/false,
        errors: [ ... ]
      }
    """
    _cleanup_expired()

    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files selected"}), 400

    # Parse options from form
    extract_images = request.form.get("extract_images", "true").lower() == "true"
    skip_cover = request.form.get("skip_cover", "false").lower() == "true"
    toc_mode = request.form.get("toc_mode", "none")  # none / toc_only / before_toc / before_toc_keep_abstract
    if toc_mode not in ("none", "toc_only", "before_toc", "before_toc_keep_abstract"):
        toc_mode = "none"
    table_cell_break = request.form.get("table_cell_break", "space")
    if table_cell_break not in ("space", "br"):
        table_cell_break = "space"

    # Create temp working directories with session ID
    session_id = uuid.uuid4().hex[:12]
    upload_dir = UPLOADS_DIR / session_id
    converted_dir = CONVERTED_DIR / session_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    converted_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"[{session_id}] Created directories - upload: {upload_dir}, converted: {converted_dir}")

    converted_files: list[dict] = []
    errors: list[dict] = []

    try:
        for file in files:
            if not file.filename or file.filename == "":
                continue

            original_name = file.filename
            ext = Path(original_name).suffix.lower()

            if not _allowed_file(original_name):
                errors.append({
                    "file": original_name,
                    "error": f"Unsupported format: {ext} (only .docx)",
                })
                continue

            # Save uploaded file to upload_dir with UUID prefix to avoid conflicts
            safe_name = f"{uuid.uuid4().hex[:8]}_{Path(original_name).name}"
            input_path = upload_dir / safe_name
            
            try:
                file.save(str(input_path))
                logger.info(f"[{session_id}] Saved upload: {original_name} → {input_path}")
            except Exception as e:
                errors.append({"file": original_name, "error": f"Upload failed: {e}"})
                continue

            # Prepare output paths in converted_dir
            md_name = Path(original_name).stem + ".md"
            output_subdir = converted_dir / Path(original_name).stem
            output_subdir.mkdir(parents=True, exist_ok=True)
            output_md = output_subdir / md_name

            try:
                if ext == ".docx":
                    convert_word_to_markdown(
                        input_path, output_md,
                        extract_images=extract_images,
                        skip_cover=skip_cover,
                        toc_mode=toc_mode,
                        table_cell_break=table_cell_break,
                    )

                logger.info(f"[{session_id}] Converted: {original_name} → {output_md}")

                # Read back converted markdown for preview
                md_content = output_md.read_text(encoding="utf-8") if output_md.exists() else ""

                # Check if there are extra resource files (images, etc.)
                all_output_files = list(output_subdir.rglob("*"))
                resource_files = [f for f in all_output_files if f.is_file() and f != output_md]

                converted_files.append({
                    "file": original_name,
                    "md_name": md_name,
                    "content": md_content,
                    "output_dir": str(output_subdir),
                    "has_resources": len(resource_files) > 0,
                })
            except Exception as e:
                import traceback
                traceback.print_exc()
                errors.append({"file": original_name, "error": str(e)})
                logger.error(f"[{session_id}] Conversion failed for {original_name}: {e}")

        if not converted_files:
            # No successful conversions — clean up and return error
            shutil.rmtree(upload_dir, ignore_errors=True)
            shutil.rmtree(converted_dir, ignore_errors=True)
            logger.warning(f"[{session_id}] No successful conversions, cleaned up directories")
            return jsonify({
                "error": "No files were converted successfully",
                "details": errors,
            }), 400

        # Decide download format
        has_resources = any(f["has_resources"] for f in converted_files)
        needs_zip = len(converted_files) > 1 or has_resources

        # Store result for later download
        result_id = uuid.uuid4().hex[:12]
        with _results_lock:
            _results[result_id] = {
                "upload_dir": str(upload_dir),
                "converted_dir": str(converted_dir),
                "files": converted_files,
                "needs_zip": needs_zip,
                "created": time.time(),
            }

        logger.info(f"[{session_id}] Result stored: {result_id} ({len(converted_files)} file(s), needs_zip={needs_zip})")

        return jsonify({
            "id": result_id,
            "files": [
                {"name": f["md_name"], "content": f["content"]}
                for f in converted_files
            ],
            "needs_zip": needs_zip,
            "errors": errors,
        })

    except Exception as e:
        # Catch-all for unexpected errors
        import traceback
        traceback.print_exc()
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(converted_dir, ignore_errors=True)
        logger.error(f"[{session_id}] Conversion request failed: {e}")
        return jsonify({"error": f"Server error: {e}"}), 500


@app.route("/download/<result_id>")
def download(result_id: str):
    """Download converted result.

    - Single .md with no extra resources → direct .md file download
    - Multiple files or has images       → .zip package
    """
    with _results_lock:
        result = _results.get(result_id)

    if not result:
        return jsonify({"error": "Result not found or expired"}), 404

    converted_files = result["files"]
    needs_zip = result["needs_zip"]

    try:
        if not needs_zip and len(converted_files) == 1:
            # Single .md file — send directly
            item = converted_files[0]
            md_content = item["content"]
            buf = io.BytesIO(md_content.encode("utf-8"))
            logger.info(f"Serving single .md file for result {result_id}")
            return send_file(
                buf,
                mimetype="text/markdown; charset=utf-8",
                as_attachment=True,
                download_name=item["md_name"],
            )

        # Build ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in converted_files:
                output_dir = Path(item["output_dir"])
                prefix = output_dir.name

                for filepath in output_dir.rglob("*"):
                    if filepath.is_file():
                        arcname = f"{prefix}/{filepath.relative_to(output_dir)}"
                        zf.write(str(filepath), arcname)

        zip_buffer.seek(0)

        if len(converted_files) == 1:
            zip_name = f"{Path(converted_files[0]['file']).stem}.zip"
        else:
            zip_name = "converted_markdown.zip"

        logger.info(f"Serving .zip for result {result_id} ({len(converted_files)} file(s))")
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_name,
        )
    except Exception as e:
        logger.error(f"Download failed for result {result_id}: {e}")
        return jsonify({"error": f"Download failed: {e}"}), 500


@app.route("/files/<result_id>/<path:filepath>")
def serve_file(result_id: str, filepath: str):
    """Serve resource files (images, etc.) from converted results."""
    with _results_lock:
        result = _results.get(result_id)

    if not result:
        return jsonify({"error": "Result not found or expired"}), 404

    converted_dir = Path(result["converted_dir"])
    target = (converted_dir / filepath).resolve()

    # Security: ensure the resolved path is within the converted_dir
    if not str(target).startswith(str(converted_dir.resolve())):
        return jsonify({"error": "Access denied"}), 403

    if not target.is_file():
        return jsonify({"error": "File not found"}), 404

    return send_file(str(target))


@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "0.1.0"})


@app.route("/config")
def config():
    """Show current configuration (for debugging)."""
    with _results_lock:
        active_results = len(_results)
    
    return jsonify({
        "uploads_dir": str(UPLOADS_DIR),
        "converted_dir": str(CONVERTED_DIR),
        "result_ttl_seconds": _RESULT_TTL,
        "active_results": active_results,
    })


@app.route("/cleanup", methods=["POST"])
def cleanup():
    """Manually trigger cleanup of expired results (admin endpoint)."""
    count = _cleanup_expired()
    with _results_lock:
        active_count = len(_results)
    return jsonify({
        "status": "ok",
        "cleaned": count,
        "active_results": active_count,
    })


def main(host: str = "0.0.0.0", port: int = 5000, debug: bool = False):
    """Start the web server."""
    print(f"Starting doc2md web service on http://{host}:{port}")
    app.run(host=host, port=port, debug=debug)


if __name__ == "__main__":
    main(debug=True)
