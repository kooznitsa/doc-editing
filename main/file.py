import os
import shutil
import zipfile
import uuid
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage
from werkzeug.wrappers.response import Response
from flask import send_from_directory
import win32com.client

from config import STATIC_FOLDER, ALLOWED_EXTENSIONS


class File(object):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    
    def __init__(self, uuid_name: uuid.UUID) -> None:
        self.input_path = os.path.join(STATIC_FOLDER, f'input-{uuid_name}/')
        self.output_path = os.path.join(STATIC_FOLDER, f'output-{uuid_name}/')
        self.archive_name = f'edited-{uuid_name}.zip'

    def create_directories(self) -> None:
        os.mkdir(self.input_path)
        os.mkdir(self.output_path)

    def allowed_file(self, filename: str) -> bool:
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    
    def upload_file(self, file: FileStorage) -> tuple[str, str]:
        filename = secure_filename(file.filename) \
                if file.filename else secure_filename(str(uuid.uuid4()))
        if self.allowed_file(filename):
            file.save(os.path.join(self.input_path, filename))
            return filename, 'success'
        return f'{filename} is not a .doc(x) file', 'error'
    
    def edit_file(self, file: str) -> None:
        doc = self.word.Documents.Open(os.path.abspath(self.input_path + file))

        def show_changes() -> None:
            doc.Activate()
            self.word.ActiveDocument.TrackRevisions = True
            doc.ShowRevisions = 0
            
        def add_start_text() -> None:
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

    def download_files(self) -> Response:
        zipfolder = zipfile.ZipFile(
            os.path.join(STATIC_FOLDER, self.archive_name), 
            'w', compression=zipfile.ZIP_STORED
        )

        for root, dirs, files in os.walk(self.output_path):
            for file in files:
                zipfolder.write(os.path.join(self.output_path, file), arcname=file)
        zipfolder.close()

        return send_from_directory(STATIC_FOLDER, self.archive_name, as_attachment=True)
    
    def delete_directory(self, path: str) -> None:
        try:
            shutil.rmtree(path)
            print(f'Directory {path} removed successfully')
        except OSError as error:
            print(f'Directory {path} cannot be removed: {error}')

    def delete_zip(self) -> None:
        try:
            os.remove(os.path.join(STATIC_FOLDER, self.archive_name))
            print(f'File {self.archive_name} removed successfully')
        except OSError as error:
            print(f'File {self.archive_name} cannot be removed: {error}')