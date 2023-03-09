import os
import time
import uuid
import concurrent.futures
from flask import Flask, redirect, render_template, session, url_for
from queue import Queue, Empty
from threading import Thread
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

queue = Queue()


@app.route('/')
def index():
    return redirect(url_for('upload'))


@app.route('/upload/', methods=['GET', 'POST'])
def upload():
    upload_form = UploadForm()
    filenames = {}

    if upload_form.submit1.data and upload_form.validate_on_submit():
        file.create_directories()

        with concurrent.futures.ThreadPoolExecutor() as tpe:
            filenames = dict(tpe.map(file.upload_file, upload_form.doc_files.data))
            session['filenames'] = filenames

        return redirect(url_for('edit'))

    return render_template('upload.html', 
                            upload_form=upload_form, 
                            title='Upload files',)


@app.route('/edit/', methods=['GET', 'POST'])
def edit():
    edit_form = EditForm()
    filenames = session.get('filenames', None)

    if edit_form.submit2.data and edit_form.validate_on_submit():
        with concurrent.futures.ProcessPoolExecutor() as ppe:
            t1 = time.perf_counter()
            ppe.map(file.edit_file, os.listdir(file.input_path))
            t2 = time.perf_counter()
            session['success_message'] = f'Editing finished in {round(t2 - t1, 2)} seconds.'
        return redirect(url_for('download'))

    return render_template('edit.html', 
                            edit_form=edit_form, 
                            title='Edit files',
                            filenames=filenames,)


def countdown(seconds):
    while seconds >= 0:
        m, s = divmod(seconds, 60)
        timer = f'{m:02d}:{s:02d}'
        print('Time until file gets deleted:', timer, end='\r')
        time.sleep(1)
        seconds -= 1

def execute_queue():
    while True:
        try:
            q = queue.get()
            q()
        except Empty:
            pass
        
Thread(target=execute_queue, daemon=True).start()

@app.route('/download/', methods=['GET', 'POST'])
def download():
    download_form = DownloadForm()
    success_message = session.get('success_message', None)

    if download_form.submit3.data and download_form.validate_on_submit():
        download = file.download_files()
        file.delete_directory(file.input_path)
        file.delete_directory(file.output_path)
        
        queue.put(lambda: countdown(10))
        queue.put(lambda: file.delete_zip())

        return download
    
    return render_template('download.html', 
                            download_form=download_form, 
                            title='Download files',
                            success_message=success_message,)


if __name__ == '__main__':
    app.run(port=PORT, threaded=True, debug=True)