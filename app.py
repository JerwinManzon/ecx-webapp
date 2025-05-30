import os
import re
import random
import smtplib
import json
from datetime import datetime
from flask import Flask, render_template, request, redirect, flash, send_from_directory
from email.message import EmailMessage
from werkzeug.utils import secure_filename
from google.oauth2.service_account import Credentials
import gspread
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from dotenv import load_dotenv
load_dotenv()


UPLOAD_FOLDER = 'static/uploads'
TEMPLATE_PATH = 'Template.docx'
CREDENTIALS_PATH = 'ecx-progression-app-460516-54ee54fb1563.json'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

# Email
SENDER_EMAIL = os.getenv("ECX_SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("ECX_EMAIL_PASSWORD")


# Google Sheets setup
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = json.loads(os.getenv("GOOGLE_CREDS_JSON"))
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)

client = gspread.authorize(creds)
sheet = client.open("ECX PROGRESSION").sheet1

app = Flask(__name__)
app.secret_key = "ecx-secret"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

FIELDS = [
    "Employee Email", "Immediate Supervisor Email", "Reported By", "Reported By Title/Role", "Date of Report",
    "Employee Name", "Employee Title/Role", "Date of Incident", "Time of Incident",
    "Immediate Supervisor", "Department Head", "Alleged Violation", "Location",
    "Specific Area of Location", "Additional Person(s) Involved", "Witnesses",
    "Incident Description", "Employee Explanation", "Action Taken by Immediate Supervisor",
    "Immediate Supervisor Recommendation"
]

RECOMMENDATION_OPTIONS = [
    "Select recommendation...", "NA", "Coaching", "Verbal Warning",
    "1st Written Warning", "2nd Written Warning", "3rd Written Warning",
    "Final Warning", "For HR Review", "For HR Intervention", "For Management Review"
]


MULTILINE_FIELDS = {
    "Additional Person(s) Involved", "Witnesses",
    "Incident Description", "Employee Explanation",
    "Action Taken by Immediate Supervisor"
}

DEPARTMENT_HEAD_OPTIONS = [
    "Select Department Head...","Amelyn Talastas", "Monique Curia", "Samuel Pastrana", "Jaycee Quijano"
]

DEPARTMENT_HEAD_EMAILS = {
    "Amelyn Talastas": "amelyntalastas.ecx@gmail.com",
    "Monique Curia": "moniquecuria.ecx@gmail.com",
    "Samuel Pastrana": "samuelpastrana.ecx@gmail.com",
    "Jaycee Quijano": "Jaycee@ecxperience.com"
}

ROLE_OPTIONS = [
    "Select Title/Role...", "Agent", "Team Lead","Trainer", "Data Scientist", "Assistant Operations Manager",
    "Operations Manager","Operations Success Coordinator", "HR", "Finance", "Accounting", "IT", "Security", "Utility"
]

def is_valid_email(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email)

def generate_incident_no():
    existing_ids = sheet.col_values(2)[1:]
    while True:
        ir = f"IR-{random.randint(1000000000, 9999999999)}"
        if ir not in existing_ids:
            return ir

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        values = []
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        incident_no = generate_incident_no()
        values.extend([timestamp, incident_no])

        for field in FIELDS:
            val = request.form.get(field, "").strip()
            if not val:
                flash(f"Please complete the '{field}' field.")
                return redirect("/")
            values.append(val)

        # Email validation
        if not is_valid_email(values[2]) or not is_valid_email(values[3]):
            flash("Invalid email address.", "danger")
            return redirect("/")

        # Save to Google Sheet
        sheet.append_row(values)

        # Process uploaded images
        images = request.files.getlist("images")
        image_paths = []
        for image in images:
            if image and allowed_file(image.filename):
                filename = secure_filename(image.filename)
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                image.save(save_path)
                image_paths.append(save_path)

        # Generate docx
        docx_path = os.path.join("generated", f"{incident_no}.docx")
        os.makedirs("generated", exist_ok=True)
        fill_docx(values, image_paths, docx_path, incident_no)

        # Email
        # Determine recipients
        recipients = [values[2], values[3]] # employee, supervisor HR
        dep_head = values[12]  # index for "Department Head"

        if dep_head in DEPARTMENT_HEAD_EMAILS:
            recipients.append(DEPARTMENT_HEAD_EMAILS[dep_head])

        send_email_with_attachment(
            recipients,
            f"Automated Notification: Incident Report Logged – {incident_no}",
            f"""This is an automated notification to inform you that an incident report has been logged in to the system:

{incident_no}

The matter is currently under review. Updates will be provided as they become available.

For questions or follow-ups, please contact our HR team at hr@ecxperience.com.

This is a system-generated message. No action is required unless otherwise indicated.""",
            docx_path
        )

        flash("Incident report submitted successfully.", "success")
        return redirect(f"/?generated={incident_no}")

    # GET method (default view)
    today_str = datetime.now().strftime("%Y-%m-%d")
    generated_file = request.args.get("generated")
    return render_template(
        "form.html",
        fields=FIELDS,
        multiline_fields=MULTILINE_FIELDS,
        generated_file=generated_file,
        today=today_str,
        recommendation_options=RECOMMENDATION_OPTIONS,
        department_head_options=DEPARTMENT_HEAD_OPTIONS,
        role_options=ROLE_OPTIONS
    )


def fill_docx(data, images, save_path, incident_no):
    from docx.oxml import OxmlElement
    from docx.shared import Inches
    from docx import Document
    from datetime import datetime

    def insert_paragraph_after(paragraph):
        new_p = OxmlElement("w:p")
        paragraph._p.addnext(new_p)
        return paragraph._parent.add_paragraph()

    doc = Document(TEMPLATE_PATH)

    # Format Time of Incident
    try:
        raw_time = data[10]
        time_12hr = datetime.strptime(raw_time, "%H:%M").strftime("%I:%M %p")
    except Exception:
        time_12hr = data[10]

    placeholders = {
        '[Incident No.]': incident_no,
        '[Reported By]': data[4],
        '[Title / Role]': data[5],
        '[Date of Report]': data[6],
        '[Employee Name]': data[7],
        '[Employee Title / Role]': data[8],
        '[Date of Incident]': data[9],
        '[Time of Incident]': time_12hr,
        '[Immediate Supervisor]': data[11],
        '[Department Head]': data[12],
        '[Alleged Violation]': data[13],
        '[Location]': data[14],
        '[Specific Area of Location]': data[15],
        '[Additional Person(s) Involved]': data[16],
        '[Witnesses]': data[17],
        '[Incident Description]': data[18],
        '[Employee Explanation]': data[19],
        '[Action Taken]': data[20],
        '[Recommendation]': data[21],
    }

    def replace_in_paragraph(paragraph):
        full_text = ''.join(run.text for run in paragraph.runs)
        new_text = full_text
        for key, val in placeholders.items():
            new_text = new_text.replace(key, val)
        if new_text != full_text:
            for run in paragraph.runs:
                run.text = ''
            paragraph.runs[0].text = new_text

    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        replace_in_paragraph(para)

    # Replace placeholders in table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para)

    # Find the paragraph with "INDEX:"
    index_paragraph = None
    for para in doc.paragraphs:
        if para.text.strip().upper() == "INDEX:":
            index_paragraph = para
            break

    if index_paragraph:
        for img_path in images:
            run = index_paragraph.add_run()
            run.add_break()  # spacing after "INDEX:"
            run.add_picture(img_path, width=Inches(4))
            run.add_break()  # spacing between pictures

    else:
        for img_path in images:
            doc.add_picture(img_path, width=Inches(4))
            doc.add_paragraph("")

    doc.save(save_path)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory("generated", filename, as_attachment=True)

def send_email_with_attachment(to_emails, subject, body, attachment_path):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(attachment_path))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.starttls()
            smtp.login(SENDER_EMAIL, EMAIL_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        print(f"Failed to send email: {e}")
        flash("Form submitted, but failed to send email notification.", "danger")


if __name__ == "__main__":
    app.run(debug=True)
