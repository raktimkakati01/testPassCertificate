import random
import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, send_file
import pandas as pd
from fpdf import FPDF
import matplotlib
matplotlib.use('Agg')  # Use the 'Agg' backend for rendering to files
import matplotlib.pyplot as plt
from forms import CompanyForm, CategoryForm, TestCaseForm


app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'static/uploads/'

inch_to_mm = 25.4
custom_page_width = 16.55 * inch_to_mm  # Convert 16.55 inches to millimeters
custom_page_height = 11.7 * inch_to_mm  # Convert 11.7 inches to millimeters
custom_page_size = (custom_page_width, custom_page_height)

# Load the Excel file with test cases
file_path = 'test_cases/test_cases_data.xlsx'  # Update this path to where your file is located
xls = pd.ExcelFile(file_path)

# Extract categories from sheet names
categories = xls.sheet_names

def get_test_cases(category):
    """Retrieve test cases for a given category."""
    df = pd.read_excel(xls, sheet_name=category)
    df_cleaned = df.dropna(how='all').reset_index(drop=True)
    
    try:
        df_cleaned.columns = ['Test ID', 'Test Case', 'Test Data', 'Steps', 'Expected Result', 'Actual Result', 'Status']
        df_cleaned = df_cleaned.dropna(subset=['Test Case'])
        return df_cleaned['Test Case'].tolist()
    except ValueError:
        return []

def get_system_time():
    """Get the current system time in the desired format."""
    return datetime.now().strftime('%d/%m/%Y %H:%M:%S UTC')

def get_next_run_id():
    """Generate the next run ID based on the number of PDF reports generated."""
    files = os.listdir(app.config['UPLOAD_FOLDER'])
    pdf_count = len([file for file in files if file.endswith('.pdf')])
    return pdf_count + 1

def generate_summary_graph(success_count, fail_count):
    """Generate a pie chart summary of test results."""
    labels = 'Passed', 'Failed'
    sizes = [success_count, fail_count]
    colors = ['#00a65a', '#f56954']

    plt.figure(figsize=(4, 4))
    plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

    graph_path = os.path.join(app.config['UPLOAD_FOLDER'], 'summary_graph.png')
    plt.savefig(graph_path, bbox_inches='tight')
    plt.close()
    return graph_path

@app.route('/', methods=['GET', 'POST'])
def company():
    form = CompanyForm()
    if form.validate_on_submit():
        session['company_name'] = form.company_name.data
        return redirect(url_for('select_categories'))
    return render_template('company.html', form=form)

@app.route('/select_categories', methods=['GET', 'POST'])
def select_categories():
    form = CategoryForm()
    form.categories.choices = [(cat, cat) for cat in categories]
    form.tester_name.choices = [
        ('nipjyoti.saikia', 'Nipjyoti Saikia'),
        ('kongkona.das', 'Kongkona Das'),
        ('sanjay.kumar.singha', 'Sanjay Kumar Singha'),
        ('jonak.das', 'Jonak Das'),
        ('mushtaq.rejowan', 'Mushtaq Rejowan'),
        ('jitul.kumar.lahon', 'Jitul Kumar Lahon'),
        ('kumar.ankit', 'Kumar Ankit'),
        ('adin.ankur.saikia', 'Adin Ankur Saikia'),
        ('raktim.kakati', 'Raktim Kakati'),
        ('ankur.duarah', 'Ankur Duarah'),
        ('meghna.dutta', 'Meghna Dutta')
    ]

    if request.method == 'POST' and form.validate_on_submit():
        selected_categories = request.form.getlist('categories')
        tester_name = request.form.get('tester_name')
        selected_browsers = request.form.getlist('browsers')
        selected_devices = request.form.getlist('devices')
        environment = form.environment.data

        session['selected_categories'] = selected_categories
        session['tester_name'] = tester_name
        session['selected_browsers'] = selected_browsers
        session['selected_devices'] = selected_devices
        session['environment'] = environment
        return redirect(url_for('select_test_cases'))

    return render_template('select_categories.html', form=form)



@app.route('/select_test_cases', methods=['GET', 'POST'])
def select_test_cases():
    selected_categories = session.get('selected_categories', [])
    form = TestCaseForm()

    test_cases_dict = {}
    for category in selected_categories:
        test_cases_dict[category] = get_test_cases(category)

    if request.method == 'POST':
        selected_test_cases = {}
        for category in selected_categories:
            selected_test_cases[category] = request.form.getlist(f'test_cases_{category}')
        session['selected_test_cases'] = selected_test_cases
        return redirect(url_for('generate_report'))

    return render_template('select_test_cases.html', form=form, test_cases_dict=test_cases_dict)


def clean_text(text):
    text = text.replace('\u2019', "'").replace('\u2018', "'")
    text = text.replace('\u201C', '"').replace('\u201D', '"')
    text = text.replace('\u2013', '-').replace('\u2014', '-')
    return text

class PDF(FPDF):
    def header(self):
        if self.page_no() == 1:
            logo_path = 'static/logo.png'
            self.image(logo_path, x=(custom_page_width / 2) - 20, y=60, w=40)  # Centering the logo
        else:
            logo_path = 'static/logo.png'
            self.image(logo_path, x=custom_page_width - 50, y=10, w=40)  # Top-right corner logo
            self.set_font('DejaVu', 'B', 18)
            self.set_text_color(23, 22, 112)
            self.set_xy(10, 20)
            self.cell(0, 10, "QA Test Completion Certificate", ln=True, align='L')

    def footer(self):
        self.set_y(-15)
        self.set_font('DejaVu', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def add_table_row(self, data, column_widths, row_height=8):
        max_lines_test_case = len(self.multi_cell(column_widths[1], row_height, data[1], border=0, align='L', split_only=True))
        actual_row_height = row_height * max_lines_test_case
        if self.get_y() + actual_row_height > self.h - 30:
            self.add_page()

        self.set_font('DejaVu', '', 10)
        self.set_text_color(0, 128, 0)
        self.cell(column_widths[0], actual_row_height, data[0], border=1, align='C')  # Test Case ID
        self.set_text_color(0, 0, 0)

        x, y = self.get_x(), self.get_y()
        self.multi_cell(column_widths[1], row_height, data[1], border=1, align='L')  # Test Cases (wrap text)
        self.set_xy(x + column_widths[1], y)

        self.cell(column_widths[2], actual_row_height, data[2], border=1, align='L')  # Test Steps
        self.cell(column_widths[3], actual_row_height, data[3], border=1, align='L')  # Device Used
        self.cell(column_widths[4], actual_row_height, data[4], border=1, align='L')  # Expected Result
        self.cell(column_widths[5], actual_row_height, data[5], border=1, align='L')  # Actual Result
        self.cell(column_widths[6], actual_row_height, data[6], border=1, align='L')  # Status
        self.ln(actual_row_height)

@app.route('/generate_report', methods=['GET'])
def generate_report():
    company_name = session.get('company_name', 'Unknown_Company').replace(" ", "").lower()
    tester_name = session.get('tester_name', 'Unknown Tester')
    selected_test_cases = session.get('selected_test_cases', {})
    selected_devices = session.get('selected_devices', [])
    environment = session.get('environment', 'uat')

    environment_url = f"https://{company_name}.vantagecircle.com" if environment == 'production' else 'https://api.vantagecircle.co.in'
    test_date = session.get('test_date', 'Unknown Date')

    pdf = PDF(format=(custom_page_width, custom_page_height))
    pdf.add_font('DejaVu', '', 'dejavu-sans/DejaVuSans.ttf', uni=True)
    pdf.add_font('DejaVu', 'B', 'dejavu-sans/DejaVuSans-Bold.ttf', uni=True)
    pdf.add_font('DejaVu', 'I', 'dejavu-sans/DejaVuSans-Oblique.ttf', uni=True)

    pdf.add_page()
    pdf.set_xy(0, 110)
    pdf.set_font('DejaVu', 'B', 36)
    pdf.set_text_color(0, 128, 0)
    pdf.cell(0, 10, 'QA Test Completion Certificate', ln=True, align='C')

    pdf.add_page()
    pdf.set_font('DejaVu', 'B', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f"Name: {company_name.replace('_', ' ').title()}", ln=True)
    pdf.cell(0, 10, f"Triggered by: {tester_name}", ln=True)
    pdf.cell(0, 10, f"Environment: {environment.capitalize()}", ln=True)
    pdf.cell(0, 10, f"URL: {environment_url}", ln=True)
    pdf.cell(0, 10, f"Result: SUCCESS", ln=True)
    pdf.cell(0, 10, f"Browsers: {', '.join(session.get('selected_browsers', []))}", ln=True)
    pdf.cell(0, 10, f"Date of Test: {test_date}", ln=True)

    pdf.ln(10)

    # Column widths for the table
    column_widths = [20, 80, 30, 80, 50, 50, 20]

    # Table Header
    pdf.set_font('DejaVu', 'B', 10)
    header_data = ["TC ID", "Test Cases", "Test Steps", "Device Used", "Expected Result", "Actual Result", "Status"]
    pdf.add_table_row(header_data, column_widths, row_height=8)

    # Add test case rows
    pdf.set_font('DejaVu', '', 10)
    for i, (category, cases) in enumerate(selected_test_cases.items(), start=1):
        for j, test_case in enumerate(cases, start=1):
            if isinstance(test_case, str):
                # Ensure test_case is a dictionary-like structure
                test_case_data = {
                    'description': test_case,
                    'steps': 'N/A',
                    'expected': 'N/A',
                    'actual': 'N/A',
                    'status': 'Pass'
                }
            else:
                test_case_data = test_case
            
            row_data = [
                str(j),  # Test Case ID
                test_case_data.get('description', 'N/A'),  # Test Cases
                test_case_data.get('steps', 'N/A'),  # Test Steps
                ', '.join(selected_devices),  # Device Used
                test_case_data.get('expected', 'N/A'),  # Expected Result
                test_case_data.get('actual', 'N/A'),  # Actual Result
                test_case_data.get('status', 'Pass')  # Status
            ]
            pdf.add_table_row(row_data, column_widths, row_height=8)

    pdf_file = f"{company_name}_QA_Report_{get_next_run_id()}.pdf"
    pdf.output(os.path.join(app.config['UPLOAD_FOLDER'], pdf_file))

    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], pdf_file), as_attachment=True)

def get_next_run_id():
    # Dummy function to get next run ID
    return 1

if __name__ == '__main__':
    app.run(debug=True)