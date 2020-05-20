from flask import Flask, request
from export_excel import export_excel

app = Flask(__name__)


@app.route('/api/export-report', methods=['POST'])
def export_report():
    req_data = request.get_json()
    report_data = {
        "logo_url": req_data['logo_url'],
        "title": req_data['title'],
        "range_time": req_data['range_time'],
        "total": req_data['total'],
        "tables": req_data['tables']
    }

    report = export_excel(report_data)

    return report


if __name__ == '__main__':
    app.run(debug=True, port=5000)
