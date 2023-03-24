import os
import shutil
from typing import Optional
import uuid
import zipfile

from flask import send_from_directory
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename
from werkzeug.wrappers.response import Response

from config import ALLOWED_EXTENSIONS, STATIC_FOLDER
from docfile import DocFile
from utils import REPL_DICT


class File(object):
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
        return f'{filename} is not a .DOCX file', 'error'
    
    def edit_file(self, 
                  file: str, 
                  language: str,
                  date_format: str,
                  start_text: Optional[str] = None) -> None:
        doc = DocFile(self.input_path, self.output_path, file)

        for regex, replace_str in REPL_DICT.items():
            doc.replace_text(regex, replace_str, language, date_format)

        doc.add_start_text(start_text)
        doc.save_file()
    
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