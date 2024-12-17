from flask import Flask, jsonify, request
from flask_cors import CORS
import os
import backend
#______________________________________________________________________________________________________________________#
# Basics
app = Flask(__name__)
CORS(app, resources={r'/api/*': {'origins': "*"}})

# Funções diversas
def start_flask():
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)
#______________________________________________________________________________________________________________________#
# Rotas da API

@app.route('/api/login', methods=['POST', 'GET'])
def loginWrite():
    if request.method == 'POST':
        data = request.json
        user = data['user']
        password = data['password']
        if not user or not password:
            return jsonify({'Success':False, 'Message':'Usuário ou senha inválidos'}), 400
        backend.txtLoginWrite(user, password)
        return jsonify({'Success':True}), 200
    elif request.method == 'GET':
        user, password = backend.txtLogin()
        if not user or not password:
            return jsonify({'Success': False, 'Message': 'Login.txt está vazio'}), 401
        else:
            return jsonify({'Success': True, 'Message': 'Login.txt não está vazio'}), 200
    return jsonify({'Success': False, 'Message':'Erro ao tentar se conectar'}), 400

@app.route('/api/generate-daily', methods=['GET'])
def generate_dr():
    Success = backend.daily_report()
    if Success:
        return jsonify({'Success': True}), 200
    return jsonify({'Success': False, 'Message':'Algo deu errado'}), 400

@app.route('/api/send-email', methods=['GET'])
def send_email():
    Success = backend.send_mail()
    if Success:
        return jsonify({'Success': True}), 200
    return jsonify({'Success': False, 'Message':'Algo deu errado'}), 400

# Rota de Desligamento
@app.route('/shutdown', methods=['POST'])
def shutdown_flask():
    os._exit(0)