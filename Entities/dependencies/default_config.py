from typing import Dict
from getpass import getuser

default:Dict[str, Dict[str,object]] = {
    'log': {
        'hostname': 'Patrimar-RPA',
        'port': '80',
        'token': 'Central-RPA'
    },
    'credenciais': {
        'api': 'ClickCompliance'
    }
}