from operator import index
import os
from flask import Flask, flash, request, redirect, url_for,render_template,send_from_directory,send_file, jsonify
from werkzeug.utils import secure_filename
from extractor import *

import nltk 
try : 
    from nltk.corpus import stopwords
    stop = stopwords.words('english')
except :
    nltk.download('stopwords')
    nltk.download('punkt')
    nltk.download('averaged_perceptron_tagger')
    nltk.download('maxent_ne_chunker')
    nltk.download('words')



UPLOAD_FOLDER = os.getcwd()
ALLOWED_EXTENSIONS = {'docx', 'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
csv_file=None
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        print(request.files['file'])
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filetosend=UPLOAD_FOLDER+"\\"+file.filename
            details = extractDataPoints(filetosend, 'pdf')
            print(details)
            for i in details:
                if type(details[i]) == set :
                    details[i] = list(details[i])
            
            details_u = {}
            for i in details:
                L = []; L.append(details[i])
                details_u[i] = L
            
            df = pd.read_csv("resumes.csv")
            df1 = pd.DataFrame(details_u)
            df = df.append(df1, ignore_index = True)
            #df = df.drop_duplicates()

            df.to_csv("resumes.csv",header = True, index=False)
            
            #return jsonify(details)
            return render_template("table.html",result = details)
            
    return render_template('index.html')

@app.route('/download/<path>',methods=['GET', 'POST'])
def download_file(path):
    if request.method=='POST':
        print(path)
        return send_file(path,as_attachment=True)     # as_attachment=True do not change the format if false changes csv to xls format
    #return render_template('response.html',path=path)
    return render_template('response.html')

@app.route('/table')
def table():
    
    # converting csv to html
    data = pd.read_csv('resumes.csv')
    return render_template('resumes.html', tables=[data.to_html()], titles=[''])


if __name__ == "__main__":
    csv_file=None
    app.run(debug=True)