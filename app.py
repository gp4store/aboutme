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
        
# Added on 02 06 2025 
        start_date = request.form.get('start_date', 'start_date')
        return_date = request.form.get('return_date', 'return_date')
        vac_hours = request.form.get('vac_hours', 'vac_hours')
        cva_hours = request.form.get('cva_hours', 'cva_hours')
        vac_cva_actual = request.form.get('vac_cva_actual', 'vac_cva_actual')       
        hfl_hours = request.form.get('hfl_hours', 'hfl_hours')
        efh_hours = request.form.get('efh_hours', 'efh_hours')
        hfl_efh_actual = request.form.get('hfl_efh_actual', 'vac_cva_actual')        
        cto_hours = request.form.get('cto_hours', 'cto_hours')
        cto_actual = request.form.get('cto_actual', 'cto_actual')
        sick_start_date = request.form.get('sick_start_date', 'sick_start_date')
        sick_hours = request.form.get('sick_hours', 'sick_hours')
        
        
        # Load the template
        doc = Document(TEMPLATE_PATH)

        # Replace placeholders in the template
        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace("{emp_name}", emp_name)
            paragraph.text = paragraph.text.replace("{todays_date}", todays_date)
            paragraph.text = paragraph.text.replace("{super_name_one}", super_name_one)
            paragraph.text = paragraph.text.replace("{super_name_two}", super_name_two)
            paragraph.text = paragraph.text.replace("{badge}", badge)

# Added on 02 06 2025
            paragraph.text = paragraph.text.replace("{start_date}", start_date)
            paragraph.text = paragraph.text.replace("{return_date}", return_date)
            paragraph.text = paragraph.text.replace("{vac_hours}", vac_hours)
            paragraph.text = paragraph.text.replace("{cva_hours}", cva_hours)
            paragraph.text = paragraph.text.replace("{vac_cva_actual}", vac_cva_actual)
            paragraph.text = paragraph.text.replace("{hfl_hours}", hfl_hours)
            paragraph.text = paragraph.text.replace("{efh_hours}", efh_hours)
            paragraph.text = paragraph.text.replace("{hfl_efh_actual}", hfl_efh_actual)
            paragraph.text = paragraph.text.replace("{cto_hours}", cto_hours)
            paragraph.text = paragraph.text.replace("{cto_actual}", cto_actual)            
            paragraph.text = paragraph.text.replace("{sick_start_date}", sick_start_date)
            paragraph.text = paragraph.text.replace("{sick_hours}", sick_hours)             
        
        # Save the modified document
        output_filename = "{name} - Time Off Request.docx".format(name = emp_name)
        doc.save(output_filename)

        # Serve the file for download
        return send_file(output_filename, as_attachment=True, download_name=output_filename)

    except Exception as e:
        return f"Error generating document: {e}"

if __name__ == '__main__':
    app.run(debug=True)
