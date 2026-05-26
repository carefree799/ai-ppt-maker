import os
import json
from flask import Flask, render_template, request

app = Flask(__name__)
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'output')
app.config['SECRET_KEY'] = os.urandom(24)

os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)