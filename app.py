import io
from flask import Flask, request, send_file, abort
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR
import re
import logging

# Configure basic logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

app = Flask(__name__)


# --- Your existing DOCX formatting logic (paste the function here) ---
def apply_formatting_to_docx_stream(input_docx_stream: io.BytesIO) -> io.BytesIO:
    try:
        document = Document(input_docx_stream)
    except Exception as e:
        logging.error(f"Error opening document: {e}")
        raise ValueError("Invalid DOCX file format.")

    for paragraph in document.paragraphs:
        full_text_in_paragraph = "".join(run.text for run in paragraph.runs)

        if (
            "{{HIGHLIGHT_START}}" in full_text_in_paragraph
            or "{{BOLD_START}}" in full_text_in_paragraph
        ):
            for run_element in paragraph._element.xpath("./w:r"):
                paragraph._element.remove(run_element)

            marker_pattern = re.compile(
                r"(\{\{HIGHLIGHT_START\}\}|\{\{HIGHLIGHT_END\}\}|\{\{BOLD_START\}\}|\{\{BOLD_END\}\})"
            )
            parts = marker_pattern.split(full_text_in_paragraph)

            is_highlighted = False
            is_bold = False

            for part in parts:
                if not part:
                    continue

                print(f"Processing part: {part}")  # Debugging output

                if part == "{{HIGHLIGHT_START}}":
                    is_highlighted = True
                elif part == "{{HIGHLIGHT_END}}":
                    is_highlighted = False
                elif part == "{{BOLD_START}}":
                    is_bold = True
                elif part == "{{BOLD_END}}":
                    is_bold = False
                else:
                    new_run = paragraph.add_run(part)
                    if is_bold:
                        new_run.bold = True
                    if is_highlighted:
                        new_run.font.highlight_color = WD_COLOR.YELLOW
        else:
            if not paragraph.runs and full_text_in_paragraph:
                paragraph.add_run(full_text_in_paragraph)

    output_stream = io.BytesIO()
    document.save(output_stream)
    output_stream.seek(0)
    return output_stream


# --- Flask Endpoint Definition (MODIFIED) ---
@app.route("/format", methods=["POST"])
def format_document():
    logging.info("Received request to /format endpoint.")

    # 1. Check if the request body is empty
    if not request.data:
        logging.warning("Empty request body received.")
        return abort(400, "Request body is empty. Please upload a .docx file.")

    # 2. Check the Content-Type header (optional but good practice for binary uploads)
    # Power Automate sends 'application/octet-stream' for raw binary
    content_type = request.headers.get("Content-Type")
    if content_type != "application/octet-stream" and not content_type.startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ):
        logging.warning(
            f"Unsupported Content-Type: {content_type}. Expected application/octet-stream or docx MIME type."
        )
        return abort(
            415,
            f"Unsupported Media Type: {content_type}. Please send as binary data with Content-Type: application/octet-stream.",
        )

    try:
        # Read the raw binary content directly from request.data
        input_docx_stream = io.BytesIO(request.data)
        logging.info("Binary data read into stream. Calling formatting function...")

        # Apply formatting
        formatted_docx_stream = apply_formatting_to_docx_stream(input_docx_stream)
        logging.info("Formatting function executed. Preparing response...")

        # Return the formatted DOCX file
        return send_file(
            formatted_docx_stream,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,  # Suggests browser to download
            download_name="formatted_document.docx",  # Suggested filename for download
        )

    except (
        ValueError
    ) as ve:  # Specific error from apply_formatting_to_docx_stream if docx is invalid
        logging.error(f"Formatting error: {ve}")
        return abort(400, str(ve))
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True)
        return abort(
            500, "An internal server error occurred during document processing."
        )


# --- Run the Flask App ---
if __name__ == "__main__":
    logging.info("Starting Flask server...")
    # Use Waitress for a more production-like local environment on Windows
    # On Linux/macOS, you'd typically use gunicorn
    # from waitress import serve
    # serve(app, host='0.0.0.0', port=5000)

    # For quick testing, you can use Flask's dev server
    app.run(debug=True, host="0.0.0.0", port=5000)
