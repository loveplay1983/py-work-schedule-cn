from flask import Flask, render_template, request, send_file
import io
from datetime import datetime
import calendar
from common import generate_schedule, coworkers, MONTH_NAMES

app = Flask(__name__)

@app.route('/')
def index():
    current_year = datetime.now().year
    current_month = datetime.now().month
    return render_template('index.html', current_year=current_year, current_month=current_month, month_name=MONTH_NAMES)

@app.route('/generate', methods=['POST'])
def generate():
    year = int(request.form['year'])
    month = int(request.form['month'])
    wb = generate_schedule(year, month, coworkers)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f"schedule_{year}_{month}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)