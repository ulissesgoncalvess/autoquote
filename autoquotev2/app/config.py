import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    SECRET_KEY = os.getenv('SECRET_KEY', 'AutoQuote-Secret-Key-2024')
    SESSION_TYPE = 'filesystem'
    LOG_FOLDER = 'logs'
    LOG_FILE = 'autoquote.log'