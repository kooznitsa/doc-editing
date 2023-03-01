import os
import time
import uuid
from flask import Flask, render_template, send_from_directory
from werkzeug.utils import secure_filename
from flask_bootstrap import Bootstrap

from .forms import UploadForm, EditForm
from .formatting import Formatting, edit_docs


app = Flask(
    __name__, 
    template_folder='../templates', 
    static_folder='../static'
)

app.config.update(
    ENV='development',
    DEBUG=True,
    SECRET_KEY='top secret!',
)

bootstrap = Bootstrap(app)

ALLOWED_EXTENSIONS = {'doc', 'docx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def upload_files():
    doc = None
    upload_form = UploadForm()
    filenames = {}

    if upload_form.validate_on_submit():
        temp_path = os.path.join(app.static_folder, 'uploads/' + str(uuid.uuid4()))
        os.mkdir(temp_path)

        for file in upload_form.doc_files.data:
            file_filename = secure_filename(file.filename)
            if allowed_file(file_filename):
                doc = f'{temp_path}/{file_filename}'
                file.save(os.path.join(app.static_folder, doc))
                filenames[file_filename] = 'success'
            else:
                filenames[f'{file_filename} is not a .doc(x) file'] = 'error'

    return render_template('upload_files.html', upload_form=upload_form, doc=doc, 
                           title='File Upload', filenames=filenames)


if __name__ == '__main__':
    app.run(port=5000, debug=True)