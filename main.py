from flask import Flask,render_template,request,flash
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from spire.doc import *
from spire.doc.common import *
from pdf2image import convert_from_path
from PIL import Image
from io import BytesIO
import os
import cv2
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'webp', 'png', 'jpg', 'jpeg', 'gif','pdf','docx'}

app = Flask(__name__)
app.secret_key='super secret key'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
@app.route("/")
def helloworld():
    return render_template("index.html")
def pdf_to_grayscale(input_pdf_path, output_pdf_path):
    images = convert_from_path(input_pdf_path)

    grayscale_images = []
    for image in images:
        grayscale_image = image.convert('L')  # Convert to grayscale
        grayscale_images.append(grayscale_image)

    grayscale_images[0].save(output_pdf_path, save_all=True, append_images=grayscale_images[1:])
    return output_pdf_path
def wordtopdf(filename):
    document=Document()
    document.LoadFromFile(f'uploads/{filename}')
    newfilename=f"static/{filename.split('.')[0]}.pdf"
    document.SaveToFile(newfilename,FileFormat.PDF)
    document.Close()
    return newfilename
def pdftoword(filename,newfilename):
    cv=Converter(filename)
    cv.convert(newfilename)
    cv.close()
    return newfilename

def processIMage(filename,operation):
    print(f"this is a operation {operation} and filename is {filename}")
    img=cv2.imread(f"uploads/{filename}")
    match operation:
        case "cgray":
            a=filename.split('.')
            if a[len(a)-1] == "pdf":
                newfilename=f"static/{filename.split('.')[0]}.pdf"
                new=pdf_to_grayscale(f'uploads/{filename}',newfilename)
                return new
            elif a[len(a)-1]=="docx":
                newfilename=f"static/{filename.split('.')[0]}.pdf"
                # newfile=f"static/{filename.split('.')[0]}.docx"
                new=wordtopdf(filename)
                new=pdf_to_grayscale(new,newfilename)
                # new=pdftoword(new,newfile)
                return new
            else:
                imgprocess=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
                cv2.imwrite(newfilename,imgprocess)
                return newfilename
        case "cjpg":
            newfilename=f"static/{filename.split('.')[0]}.jpg"
            cv2.imwrite(newfilename,img)
            return newfilename
        # case "cwebp":
        #     newfilename=f"static/{filename.split('.')[0]}.jpg"
        #     cv2.imwrite(newfilename,img)
        #     return newfilename
        case "cpng":
            newfilename=f"static/{filename.split('.')[0]}.png"
            cv2.imwrite(newfilename,img)
            return newfilename
        case "cword":
            newfilename=f"static/{filename.split('.')[0]}.docx"
            filename=f"uploads/{filename}"
            new=pdftoword(filename,newfilename)
            return new
        case "cpdf":
            newfilename=wordtopdf(filename)
            return newfilename
@app.route("/about")
def about():
    return render_template("about.html")
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/edit",methods=['GET','POST'])
def edit():
    if request.method == 'POST':
        # check if the post request has the file part
        operations=request.form.get("operation")
        if 'file' not in request.files:
            flash('No file part')
            return "error"
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return "please select file"
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            new=processIMage(filename,operations)
            flash(f"your image is processed <a href='/{new}' target='_blank'>here</a>")
            return render_template("index.html")
    return render_template("index.html")
if __name__=="__main__":
    app.run(debug=True,port=7000)