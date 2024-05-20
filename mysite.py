import sys
from flask import Flask, render_template, send_file, request, render_template_string
from prepare import write_log_ove, write_log_main

# DEBUG = True
app = Flask(__name__)


# app.config.from_object(__name__)


@app.route("/", methods=['post', 'get'])
def index():
    message = 'Введите данные для запроса эфирной справки'
    filename = False
    if request.method == 'POST':
        channel = request.form.get('channel')
        partner = request.form.get('partner')
        data1 = request.form.get('data1')
        data2 = request.form.get('data2')
        type = request.form.get('type')
        col = request.form.get('col')
        if type == '1':
            message, filename = write_log_ove(channel, partner, col, data1, data2)
        if type == '2':
            message, filename = write_log_main(channel, partner, col, data1, data2)
        return render_template('index.html', message=message, filename=filename)
    else:
        return render_template('index.html', message=message, filename=False)


@app.route("/download<path:filename>", methods=['GET', 'POST'])
def Download_File(filename):
    link = 'C:\\FlaskApp\\App\\file\\' + filename
    return send_file(link, as_attachment=True)


if __name__ == "__main__":
    app.run()
