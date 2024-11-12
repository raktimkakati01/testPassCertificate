from flask_wtf import FlaskForm
from wtforms import StringField, SelectMultipleField, SubmitField, SelectField
from wtforms.validators import DataRequired

class CompanyForm(FlaskForm):
    company_name = StringField('Company Name', validators=[DataRequired()])
    submit = SubmitField('Next')

class CategoryForm(FlaskForm):
    categories = SelectMultipleField('Categories', choices=[], validators=[DataRequired()])
    tester_name = SelectField('Tester Name', choices=[], validators=[DataRequired()])
    browsers = SelectMultipleField('Browsers', choices=[
        ('chrome', 'Google Chrome'),
        ('firefox', 'Mozilla Firefox'),
        ('safari', 'Safari'),
        ('edge', 'Microsoft Edge'),
        ('opera', 'Opera'),
        ('ie', 'Internet Explorer')
    ], validators=[DataRequired()])
    
    # Add a new field for devices
    devices = SelectMultipleField('Devices', choices=[
        ('windows', 'Windows'),
        ('macbook', 'MacBook'),
        ('android', 'Android'),
        ('ios', 'iOS')
    ], validators=[DataRequired()])
    
    environment = SelectField('Environment', choices=[
        ('uat', 'UAT'),
        ('production', 'Production')
    ], validators=[DataRequired()])
    
    submit = SubmitField('Next')



class TestCaseForm(FlaskForm):
    submit = SubmitField('Generate Report')
