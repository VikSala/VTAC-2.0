import xmlrpc.client

url = 'https://b2b.optimaluz.com/'
db = 'odoo0'
username = 'admin'
password = '1324'

common = None
uid = None
models = None

def conectar():
    global common, uid, models
    try:
        common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common', allow_none=True)
        uid = common.authenticate(db, username, password, {})
        if uid:
            print(f'Conectado como {username} (uid: {uid})')
            models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object', allow_none=True)
        else:
            print('Error de autenticación')
        return uid
    except Exception as e:
        print(f'❌ Error en la conexión: {e}')
        return None
