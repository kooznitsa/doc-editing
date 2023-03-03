from flask_wtf import FlaskForm
from wtforms import MultipleFileField, SubmitField
from wtforms.validators import InputRequired, Length


class UploadForm(FlaskForm):
    doc_files = MultipleFileField('', render_kw={'multiple': True}, validators=[
        InputRequired(message='At least one file is required.'), 
        Length(max=10, message='A maximum of 10 files are allowed.'),
    ])
    submit1 = SubmitField('Upload files')


class EditForm(FlaskForm):
    submit2 = SubmitField('Edit files')


class DownloadForm(FlaskForm):
    submit3 = SubmitField('Download files')