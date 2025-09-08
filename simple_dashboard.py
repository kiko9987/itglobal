from flask import Flask, render_template, redirect, url_for

app = Flask(__name__, template_folder='dashboard/templates', static_folder='dashboard/static')
app.config['SECRET_KEY'] = 'simple-key'

@app.route('/')
def index():
    return redirect(url_for('login_page'))

@app.route('/login')
def login_page():
    try:
        return render_template('login.html')
    except Exception as e:
        return f"""
        <h1>IT Global 대시보드</h1>
        <p>로그인 페이지를 로드하는 중 오류 발생:</p>
        <p>{str(e)}</p>
        <a href="/projects">프로젝트 페이지로 이동</a>
        """

@app.route('/projects')  
def projects():
    return """
    <h1>프로젝트 관리</h1>
    <p>프로젝트 목록이 여기에 표시됩니다.</p>
    <p>서버가 정상적으로 작동하고 있습니다!</p>
    """

if __name__ == '__main__':
    print("=" * 50)
    print("IT Global 대시보드 (간단 버전)")  
    print("URL: http://localhost:5000")
    print("=" * 50)
    app.run(host='127.0.0.1', port=5000, debug=False)