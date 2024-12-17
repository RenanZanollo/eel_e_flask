import eel
import threading
from flask_app import start_flask
import requests
import os
import eel_functions

path = os.getcwd() + '\\'

#______________________________________________________________________________________________________________________#
# Funções diversas

# Função para fechar o servidor ao fechar o eel
@eel.expose
def on_close_callback(route, sockets):
    try:
        requests.post('http://127.0.0.1:5000/shutdown')
    except requests.exceptions.RequestException:
        pass
    os._exit(0)

def start_eel():
    eel.init('web')
    eel.start('index.html', size=(315,460), block=True, mode='chrome', close_callback=on_close_callback,
              fullscreen=False, position=(580, 180))


#______________________________________________________________________________________________________________________#
# Main
if __name__ == '__main__':
    threading.Thread(target=start_flask).start()
    if not os.path.exists(path + 'login.txt'):
        with open(path + 'login.txt', 'w') as file:
            file.write(';')
    start_eel()