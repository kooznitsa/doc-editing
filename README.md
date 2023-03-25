# Automated Editing/Formatting of MS Word Files

## Project Description
The web app works with .DOCX files and performs the following actions on corresponding endpoints:
- /upload: uploading up to 10 .DOCX files
- /edit: automated files editing (adding text, text replacement, date formatting)
- /download: zipping output files into a .ZIP file and downloading the result file

## Tech Stack
<img src="https://img.shields.io/badge/Python-d93b32?style=for-the-badge&logo=python&logoColor=black"/> <img src="https://img.shields.io/badge/Flask-fc884d?style=for-the-badge&logo=Flask&logoColor=black"/> <img src="https://img.shields.io/badge/HTML5-96a4a5?style=for-the-badge&logo=HTML5&logoColor=black"/> <img src="https://img.shields.io/badge/CSS3-96a4a5?style=for-the-badge&logo=CSS3&logoColor=black"/> <img src="https://img.shields.io/badge/Bootstrap-96a4a5?style=for-the-badge&logo=Bootstrap&logoColor=black"/>

## Project Structure

### Upload
![upload.png](https://raw.githubusercontent.com/kooznitsa/doc-editing/main/screenshots/upload.png)

### Edit
![edit.png](https://raw.githubusercontent.com/kooznitsa/doc-editing/main/screenshots/edit.png)

### Download
![download.png](https://raw.githubusercontent.com/kooznitsa/doc-editing/main/screenshots/download.png)

## Run App
1. Clone the repo
    ```sh
   git clone https://github.com/kooznitsa/doc-editing.git
    ```

2. Create and activate virtual environment
   ```sh
   py -m venv venv
   venv\Scripts\activate
   ``

3. Select virtualenv which you created recently

4. Install  packages
   ```sh
   py -m pip install -r requirements.txt
   ```

5. Go to the main directory
   ```sh
   cd main
   ```

6. Run app.py file
   ```sh
   py app.py
   ```