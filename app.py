import os
from flask import Flask, render_template, request, jsonify, send_file
from num_to_words import num_to_words_inr
from docx_generator import generate_sparsh_docx

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/num_to_words', methods=['POST'])
def convert_num():
    data = request.get_json() or {}
    amt = data.get('amount', 0)
    words = num_to_words_inr(amt)
    return jsonify({'amount': amt, 'words': words})

@app.route('/api/generate_docx', methods=['POST'])
def generate_docx():
    data = request.get_json() or {}
    tot_amt = sum(float(v.get('amount', 0) or 0) for v in data.get('vendors', []))
    data['totalAmount'] = tot_amt
    data['totalAmountWords'] = num_to_words_inr(tot_amt)

    stream = generate_sparsh_docx(data)
    prc_clean = str(data.get('prcNo', 'Bill')).replace('/', '_').replace('\\', '_')
    filename = f"SPARSH_Bill_{prc_clean}.docx"

    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Starting SNA SPARSH Bill Generator on http://127.0.0.1:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
