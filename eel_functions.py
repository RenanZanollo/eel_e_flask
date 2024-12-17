# Arquivo feito unicamente para organizar as funções que serão tecnicamente usadas apenas no código HTML

import requests
import eel
import backend


@eel.expose
def login_request(username, password):
    try:
        response = requests.post('http://127.0.0.1:5000/api/login', json={'user': username, 'password': password}).json()
        if response['Success']:
            return response['Success']
        else:
            return False
    except requests.exceptions.RequestException:
        return False

@eel.expose
def get_table():
    return backend.html

@eel.expose
def get_send_to():
    if backend.CC:
        return f'''Email será enviado para: {backend.To} com cópia para {backend.CC}'''
    return f'Email será enviado para: {backend.To}'

@eel.expose
def get_error():
    return backend.exception