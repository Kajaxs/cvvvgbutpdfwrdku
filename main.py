from flask import Flask, request, render_template, send_file
from io import BytesIO
from pdf2docx import Converter
from docx import Document
from fpdf import FPDF

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    conversion_type = request.form["conversion_type"]
    uploaded_file = request.files["file"]

    if uploaded_file:
        if conversion_type == "pdf_to_word":
            # Konversi PDF ke Word
            pdf_stream = BytesIO(uploaded_file.read())
            word_stream = BytesIO()
            cv = Converter(pdf_stream)
            cv.convert(word_stream, start=0, end=None)
            cv.close()
            word_stream.seek(0)
            return send_file(word_stream, as_attachment=True, download_name="converted.docx")

        elif conversion_type == "word_to_pdf":
            # Konversi Word ke PDF
            word_stream = BytesIO(uploaded_file.read())
            pdf_stream = BytesIO()
            doc = Document(word_stream)
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            for paragraph in doc.paragraphs:
                pdf.multi_cell(0, 10, paragraph.text)
            pdf.output(pdf_stream)
            pdf_stream.seek(0)
            return send_file(pdf_stream, as_attachment=True, download_name="converted.pdf")

    return "File tidak valid!", 400

if __name__ == "__main__":
    app.run(debug=True)
