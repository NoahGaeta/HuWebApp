from tempfile import mkdtemp
from excel_sheet_integrator import LocalExcelSheetIntegrator

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    safe_join,
    send_file,
)
from werkzeug.exceptions import NotFound, BadRequest
from werkzeug.utils import secure_filename


UPLOAD_FOLDER = mkdtemp()

NO_FILE_MSG = "Missing file in request"
NO_FILE_SELECTED_MSG = "No file selected for uploading"
SUCCESSFUL_UPLOAD = "File successfully uploaded"

app = Flask(__name__)
app.secret_key = "setec astronomy"
app.config["SESSION_TYPE"] = "filesystem"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/upload", methods=["POST"])
def upload_file_locally():
    if request.method == "POST":
        file_path = handle_local_file(request.files["file"])
        master_file = request.form["master_sheet"]

        excel_sheet_integrator = LocalExcelSheetIntegrator(file_path)
        excel_sheet_integrator.create_new_excel_sheet(master_file)

        return redirect("/download/{0}".format(file_path))
    return redirect('/')


@app.route("/download/<path:file_path>", methods=["GET"])
def download_file(file_path):
    try:
        return send_file(file_path)
    except (NotFound, BadRequest):
        return redirect("/")


def handle_local_file(file):
    file_name = get_file_name(file)

    file_path = safe_join(app.config["UPLOAD_FOLDER"], file_name)
    file.save(file_path)
    flash(SUCCESSFUL_UPLOAD)
    return file_path


def get_file_name(file):
    return secure_filename(file.filename)


if __name__ == "__main__":
    app.run(debug=False, port=8080)
