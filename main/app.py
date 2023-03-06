import os
import time
import uuid
import concurrent.futures
from flask import Flask, render_template, redirect, session, url_for, send_from_directory
from flask_bootstrap import Bootstrap

from config import TEMPLATE_FOLDER, STATIC_FOLDER, PORT
from forms import UploadForm, EditForm, DownloadForm
from file import File


app = Flask(
    __name__, 
    template_folder=TEMPLATE_FOLDER, 
    static_folder=STATIC_FOLDER,
)

app.config.update(
    ENV='development',
    DEBUG=True,
    SECRET_KEY='secret',
    SESSION_PERMANENT=False,
    SESSION_TYPE='filesystem',
)

bootstrap = Bootstrap(app)


file = File(uuid.uuid4())


@app.route('/')
def index():
    return redirect(url_for('upload'))


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    upload_form = UploadForm()
    success_message = session.get('success_message', None)

    filenames = {}
    if upload_form.submit1.data and upload_form.validate_on_submit():
        file.create_directories()

        with concurrent.futures.ThreadPoolExecutor() as tpe:
            filenames = dict(tpe.map(file.upload_file, upload_form.doc_files.data))

    session.update({
        'filenames': filenames,
    })

    return render_template('upload.html', 
                            upload_form=upload_form, 
                            title='Upload files',
                            success_message=success_message)


@app.route('/edit', methods=['GET', 'POST'])
def edit():
    edit_form = EditForm()
    filenames = session.get('filenames', None)
    success_message = ''

    if edit_form.submit2.data and edit_form.validate_on_submit():
        with concurrent.futures.ProcessPoolExecutor() as ppe:
            t1 = time.perf_counter()
            ppe.map(file.edit_file, os.listdir(file.input_path))
            t2 = time.perf_counter()
            success_message = f'Editing finished in {round(t2 - t1, 2)} seconds.'

    session.update({ 
        'success_message': success_message,
    })

    return render_template('edit.html', 
                            edit_form=edit_form, 
                            title='Edit files',
                            filenames=filenames,)


@app.route('/download/', methods=['GET', 'POST'])
def download():
    download_form = DownloadForm()
    success_message = session.get('success_message', None)

    if download_form.submit3.data and download_form.validate_on_submit():
        return file.download_files()

    session.update({
        'success_message': 'Files have been downloaded.',
    })

    return render_template('download.html', 
                            download_form=download_form, 
                            title='Download files',
                            success_message=success_message,)


if __name__ == '__main__':
    app.run(port=PORT, threaded=True, debug=True)