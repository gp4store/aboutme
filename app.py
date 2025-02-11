from flask import Flask, render_template, request, send_file, redirect, url_for
from docx import Document
import os
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
db = SQLAlchemy(app)

TEMPLATE_PATH = "template.docx"
class FormData(db.Model):
    
    id = db.Column(db.Integer, primary_key=True)
    PTO = db.Column(db.String(100), nullable=False)
    HOURS = db.Column(db.String(100), nullable=False)
    DATE = db.Column(db.String(100), nullable=False)
    
    def __init__(self, PTO, HOURS, DATE):
        self.PTO = PTO
        self.HOURS = HOURS
        self.DATE = DATE

with app.app_context():
    db.create_all()

@app.route('/records', methods=['GET', 'POST'])
def index():
    
    if request.method == 'POST':
        name_pto = request.form['pto']
        hours_pto = request.form['hours']
        date_pto = request.form['date']
        Type_hours = FormData(name_pto, hours_pto, date_pto)
        db.session.add(Type_hours)
        db.session.commit()

        return redirect(url_for('success'))
    return render_template('past_requests.html')

@app.route('/log')
def log():
    users = FormData.query.all()
    return render_template("log.html", users=users)

@app.route('/succes')
def success():
    return render_template("data.html")

@app.route('/')
def home():
    now = datetime.now()
    formatted_date_time = now.strftime("%Y-%m-%d %H:%M:%S")
    return render_template("home.html", current_date_time=formatted_date_time)

@app.route('/past_request')
def past_request():
    return render_template("past_request.html")

@app.route('/form')
def form():
    return render_template("form.html")

@app.route('/generate-doc', methods=['POST'])
def generate_doc():
   
    try:

        emp_name = request.form.get('emp_name', 'John Doe')
        todays_date = request.form.get('todays_date', 'YYYY-MM-DD')
        super_name_one = request.form.get('super_name_one', 'Supervisors name')
        super_name_two = request.form.get('super_name_two', 'Supervisors name')
        badge = request.form.get('badge', 'badge')
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
        
        doc = Document(TEMPLATE_PATH)

        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace("{emp_name}", emp_name)
            paragraph.text = paragraph.text.replace("{todays_date}", todays_date)
            paragraph.text = paragraph.text.replace("{super_name_one}", super_name_one)
            paragraph.text = paragraph.text.replace("{super_name_two}", super_name_two)
            paragraph.text = paragraph.text.replace("{badge}", badge)
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
        
        output_filename = "{name} - Time Off Request.docx".format(name = emp_name)
        doc.save(output_filename)
        return send_file(output_filename, as_attachment=True, download_name=output_filename)
    except Exception as e:
        return f"Error generating document: {e}"

if __name__ == '__main__':
    app.run(debug=True)
