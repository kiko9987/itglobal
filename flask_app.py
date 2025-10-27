from flask import Flask

app = Flask(__name__)

@app.route('/')
def hello():
    return '''
    <h1>IT Global Dashboard</h1>
    <p>Flask server is working!</p>
    <p>Port: 5001</p>
    '''

@app.route('/test')
def test():
    return {'status': 'ok', 'message': 'API working'}

if __name__ == '__main__':
    print("Starting Flask server on port 5001...")
    app.run(host='0.0.0.0', port=5001, debug=True, use_reloader=False)