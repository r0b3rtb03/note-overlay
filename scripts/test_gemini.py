import os, json
from dotenv import load_dotenv
import requests

load_dotenv('.env')
key = os.getenv('GEMINI_API_KEY')
url = os.getenv('GEMINI_API_URL') or f'https://generativelanguage.googleapis.com/v1/models/text-bison-001:generate?key={key}'
print('URL:', url)
try:
    payload = {"prompt": {"text": "Say hello"}, "maxOutputTokens": 10}
    r = requests.post(url, json=payload, timeout=15)
    print('Status:', r.status_code)
    try:
        print('Response:', json.dumps(r.json(), indent=2))
    except Exception:
        print('Response text:', r.text)
except Exception as e:
    print('Request error:', str(e))
