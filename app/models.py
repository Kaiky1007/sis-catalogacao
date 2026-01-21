from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.dialects.postgresql import JSONB

db = SQLAlchemy()

class Ficha(db.Model):
    __tablename__ = 'fichas'

    id = db.Column(db.Integer, primary_key=True)
    numero_ficha = db.Column(db.String(50), unique=True)
    avaliacao = db.Column(db.Integer)
    
    autor = db.Column(db.String(200))
    titulo = db.Column(db.String(200))
    registro = db.Column(db.String(100))
    n_chamada = db.Column(db.String(100))
    secao_guarda = db.Column(db.String(100))
    data_obra = db.Column(db.String(100))
    paginas = db.Column(db.String(50))
    dimensoes = db.Column(db.String(100))

    especificacao_material = db.Column(JSONB) 
    tipo_suporte = db.Column(JSONB)
    estado_conservacao = db.Column(JSONB)
    deterioracoes = db.Column(JSONB)
    
    tratamento_planos = db.Column(JSONB)
    tratamento_volumes = db.Column(JSONB)

    observacoes = db.Column(db.Text)
    tecnico_nome = db.Column(db.String(150))
    data_preenchimento = db.Column(db.Date)

    imagens = db.relationship('Imagem', backref='ficha', cascade='all, delete-orphan')


class Imagem(db.Model):
    __tablename__ = 'imagens'
    
    id = db.Column(db.Integer, primary_key=True)
    
    caminho = db.Column(db.String(255), nullable=False)
    
    ficha_id = db.Column(db.Integer, db.ForeignKey('fichas.id'), nullable=False)
    
    def __repr__(self):
        return f'<Imagem {self.caminho}>'
