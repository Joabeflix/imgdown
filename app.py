from flask import Flask, render_template, request, redirect, flash, send_from_directory
import pandas as pd
import requests
import os
from zipfile import ZipFile

app = Flask(__name__)
app.secret_key = 'super_secret_key'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename == '':
            flash('Nenhum arquivo selecionado', 'error')
            return redirect(request.url)
        if file and file.filename.endswith('.xlsx'):
            tmp_dir = os.path.join(app.root_path, 'tmp')
            if not os.path.exists(tmp_dir):
                os.makedirs(tmp_dir)

            path = os.path.join(tmp_dir, file.filename)
            file.save(path)
            zip_path = process_excel_file(path, tmp_dir)
            os.remove(path)  # Remove o arquivo Excel após processamento
            if zip_path and os.path.exists(zip_path):
                return send_from_directory(directory=os.path.dirname(zip_path), path=os.path.basename(zip_path), as_attachment=True)
            else:
                flash('Falha ao preparar o arquivo para download.', 'error')
                return redirect(request.url)
        else:
            flash('Formato de arquivo não suportado.', 'error')
            return redirect(request.url)
    return render_template('index.html')

def process_excel_file(excel_file, tmp_dir):
    data = pd.read_excel(excel_file)
    zip_filename = "images.zip"
    zip_path = os.path.join(tmp_dir, zip_filename)
    
    with ZipFile(zip_path, 'w') as zipf:
        for index, row in data.iterrows():
            image_url = row['Link']
            image_name = f"{row['Nome Img']}.jpg"
            image_path = os.path.join(tmp_dir, image_name)
            
            try:
                response = requests.get(image_url)
                response.raise_for_status()
                with open(image_path, 'wb') as f:
                    f.write(response.content)
                zipf.write(image_path, arcname=image_name)
                os.remove(image_path)
            except Exception as e:
                flash(f"Erro ao baixar a imagem {image_url}: {e}", 'error')
        
        # Check if any file was written to the zip
        if zipf.filelist:
            zipf.close()
            return zip_path
        else:
            zipf.close()
            os.remove(zip_path)
            return None

if __name__ == '__main__':
    app.run()