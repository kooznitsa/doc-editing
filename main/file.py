import os
import zipfile
from werkzeug.utils import secure_filename
from flask import send_from_directory
import win32com.client

from config import STATIC_FOLDER, ALLOWED_EXTENSIONS


class File(object):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    
    def __init__(self, uuid_name):
        self.input_path = os.path.join(STATIC_FOLDER, f'input-{uuid_name}/')
        self.output_path = os.path.join(STATIC_FOLDER, f'output-{uuid_name}/')

    def create_directories(self):
        os.mkdir(self.input_path)
        os.mkdir(self.output_path)

    def allowed_file(self, filename):
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    
    def upload_file(self, file):
        filename = secure_filename(file.filename)
        if self.allowed_file(filename):
            file.save(os.path.join(self.input_path, filename))
            return filename, 'success'
        return f'{filename} is not a .doc(x) file', 'error'
    
    def edit_file(self, file):
        doc = self.word.Documents.Open(os.path.abspath(self.input_path + file))

        def show_changes():
            doc.Activate()
            self.word.ActiveDocument.TrackRevisions = True
            doc.ShowRevisions = 0
            
        def add_start_text():
            start_text = 'Перевод с английского языка на русский язык\n'

            fline = doc.Range(0, 0)
            fline.InsertBefore(start_text)
            fline.Font.Name = 'Times New Roman'
            fline.Font.Size = 11
            fline.Font.Italic = True
            fline.Font.Underline = 2
            fline.Font.Bold = False
            fline.Paragraphs.Alignment = win32com.client.constants.wdAlignParagraphRight

        show_changes()
        add_start_text()

        doc.SaveAs(os.path.abspath(self.output_path + file))
        doc.Close()
        self.word.Quit()

    def download_files(self):
        archive_name = 'edited_files.zip'

        zipfolder = zipfile.ZipFile(
            os.path.join(STATIC_FOLDER, archive_name), 
            'w', 
            compression=zipfile.ZIP_STORED
        )

        for root, dirs, files in os.walk(self.output_path):
            for file in files:
                zipfolder.write(os.path.join(self.output_path, file), arcname=file)
        zipfolder.close()

        return send_from_directory(STATIC_FOLDER, archive_name, as_attachment=True)
    
    def delete_directories(self):
        ...