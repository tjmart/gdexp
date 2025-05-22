from flask import Flask, render_template, request, jsonify
from datetime import datetime

app = Flask(__name__)

# 샘플 데이터 - 실제 구현시에는 데이터베이스를 사용해야 합니다
WORKS = [
    {
        'id': 1,
        'title': '작품 1',
        'description': '작품 설명 1',
        'image_url': 'images/work1.jpg',
        'created_date': '2024-01-01'
    }
]

EXHIBITIONS = [
    {
        'id': 1,
        'title': '봄 기획전',
        'description': '봄을 맞이하는 특별 전시회',
        'start_date': '2024-03-01',
        'end_date': '2024-03-31',
        'location': '예고을 갤러리'
    }
]

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/works')
def get_works():
    return jsonify(WORKS)

@app.route('/api/exhibitions')
def get_exhibitions():
    return jsonify(EXHIBITIONS)

@app.route('/contact', methods=['POST'])
def submit_contact():
    data = request.json
    # 여기에 이메일 전송 또는 데이터베이스 저장 로직을 구현
    return jsonify({'status': 'success'})

if __name__ == '__main__':
    app.run(debug=True) 