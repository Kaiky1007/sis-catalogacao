import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

app.config['SECRET_KEY'] = 'chave_super_secreta_do_museu'
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL', 'postgresql://usuario:senha@db:5432/fichas_db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

from .models import db
db.init_app(app)

with app.app_context():
    from .models import Ficha 
    db.create_all()
    print("Banco de dados conectado e tabelas verificadas.")

from .routes import *

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
