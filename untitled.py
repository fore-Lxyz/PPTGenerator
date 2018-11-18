from flask import Flask,render_template,request,redirect,url_for,send_file,make_response
from werkzeug.utils import secure_filename
import os
from autoPPT import autoCreatPPT


app = Flask(__name__)


@app.route('/',methods=['POST', 'GET'])
def upload():
    if request.method == 'POST':

        f = request.files['file']
        basepath = os.path.dirname(__file__)  # 当前文件所在路径
        upload_path = os.path.join(basepath, 'static/uploads',
                                   secure_filename(f.filename))
        f.save(upload_path)
        autoCreatPPT(upload_path)
        download_path=os.path.join(basepath, 'static/ppt',
                                                        secure_filename("test.pptx"))

        response = make_response(send_file(download_path, as_attachment=True))

        newName="autoPPT.pptx"
        response.headers["Content-Disposition"] = "attachment; filename={}".format(newName.encode().decode('latin-1'))
        return response



    return render_template('index.html')





if __name__ == '__main__':
    app.run()
