from flask import Flask, render_template, request, send_file
from docx import Document
import os

app = Flask(__name__)

# Path to the template file (Ensure 'template.docx' exists in the same directory)
TEMPLATE_PATH = "template.docx"

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/generate-doc', methods=['POST'])
def generate_doc():
    try:
        # Get user input
        emp_name = request.form.get('emp_name', 'John Doe')
        todays_date = request.form.get('todays_date', 'YYYY-MM-DD')
        super_name_one = request.form.get('super_name_one', 'Supervisors name')
        super_name_two = request.form.get('super_name_two', 'Supervisors name')
        badge = request.form.get('badge', 'badge')

        # Load the template
        doc = Document(TEMPLATE_PATH)

        # Replace placeholders in the template
        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace("{emp_name}", emp_name)
            paragraph.text = paragraph.text.replace("{todays_date}", todays_date)
            paragraph.text = paragraph.text.replace("{super_name_one}", super_name_one)
            paragraph.text = paragraph.text.replace("{super_name_two}", super_name_two)
            paragraph.text = paragraph.text.replace("{badge}", badge)

        # Save the modified document
        output_filename = "{name} - Time Off Request.docx".format(name = emp_name)
        doc.save(output_filename)

        # Serve the file for download
        return send_file(output_filename, as_attachment=True, download_name=output_filename)

    except Exception as e:
        return f"Error generating document: {e}"

if __name__ == '__main__':
    app.run(debug=True)
