from flask_wtf import FlaskForm # type: ignore
from wtforms import MultipleFileField, SelectField, StringField, SubmitField # type: ignore
from wtforms.validators import InputRequired, Length # type: ignore

from utils import get_current_date, FORMATS


class UploadForm(FlaskForm):
    doc_files = MultipleFileField('', render_kw={'multiple': True}, validators=[
        InputRequired(message='At least one file is required.'), 
        Length(max=10, message='A maximum of 10 files are allowed.'),
    ])
    submit1 = SubmitField('Upload files')


class EditForm(FlaskForm):
    start_text = StringField('Add start text', 
            validators=[Length(max=100, message='A maximum of 100 characters are allowed.')])
    date_format = SelectField(u'Choose date format', choices=[(f, get_current_date(f)) for f in FORMATS])
    submit2 = SubmitField('Edit files')


class DownloadForm(FlaskForm):
    submit3 = SubmitField('Download')