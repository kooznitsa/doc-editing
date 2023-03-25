import concurrent.futures
import os
import time
import uuid

from flask import Flask, redirect, render_template, session, url_for
from flask_bootstrap import Bootstrap # type: ignore
from queue import Queue
from threading import Thread
from werkzeug.wrappers.response import Response

from config import STATIC_FOLDER, TEMPLATE_FOLDER, PORT
from file import File
from forms import EditForm, DownloadForm, UploadForm
from utils import countdown, execute_queue


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

app.url_map.strict_slashes = False
bootstrap = Bootstrap(app)
file = File(uuid.uuid4())
queue: Queue = Queue()


@app.route('/')
def index() -> Response:
    return redirect(url_for('upload'))


@app.route('/upload', methods=['GET', 'POST'])
def upload() -> str | Response:
    upload_form = UploadForm()
    filenames = {}
    active_upload, active_edit, active_download = True, False, False

    if upload_form.submit1.data and upload_form.validate_on_submit():
        file.create_directories()

        with concurrent.futures.ThreadPoolExecutor() as tpe:
            filenames = dict(tpe.map(file.upload_file, upload_form.doc_files.data))
            session['filenames'] = filenames

        return redirect(url_for('edit'))

    return render_template('upload.html', 
                            upload_form=upload_form, 
                            title='Upload files',
                            active_upload=active_upload, 
                            active_edit=active_edit, 
                            active_download=active_download,)


@app.route('/edit', methods=['GET', 'POST'])
def edit() -> str | Response:
    edit_form = EditForm()
    filenames = session.get('filenames', None)
    active_upload, active_edit, active_download = False, True, False

    if edit_form.submit2.data and edit_form.validate_on_submit():
        t1 = time.perf_counter()

        files = os.listdir(file.input_path)
        language = edit_form.language.data
        date_format = edit_form.date_format.data
        start_text = edit_form.start_text.data

        for f in files:
            file.edit_file(f, language, date_format, start_text)

        t2 = time.perf_counter()
        session['success_message'] = f'Editing finished in {round(t2 - t1, 2)} seconds.'

        return redirect(url_for('download'))

    return render_template('edit.html', 
                            edit_form=edit_form, 
                            title='Edit files',
                            filenames=filenames,
                            active_upload=active_upload, 
                            active_edit=active_edit, 
                            active_download=active_download,)


Thread(target=execute_queue, args=(queue,), daemon=True).start()

@app.route('/download', methods=['GET', 'POST'])
def download() -> str | Response:
    """Delete temporary folders, put into queue 
    the process of deleting ZIP, download ZIP, 
    then delete it from project after 10 sec.
    """
    download_form = DownloadForm()
    success_message = session.get('success_message', None)
    active_upload, active_edit, active_download = False, False, True

    if download_form.submit3.data and download_form.validate_on_submit():
        download = file.download_files()
        file.delete_directory(file.input_path)
        file.delete_directory(file.output_path)
        
        queue.put(lambda: countdown(10))
        queue.put(lambda: file.delete_zip())

        session.clear()

        return download
    
    return render_template('download.html', 
                            download_form=download_form, 
                            title='Download files',
                            success_message=success_message,
                            filename=file.archive_name,
                            active_upload=active_upload, 
                            active_edit=active_edit, 
                            active_download=active_download,)


if __name__ == '__main__':
    app.run(port=PORT, threaded=True, debug=True)