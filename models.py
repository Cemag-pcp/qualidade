from extensions import db   # usa a instância criada em extensions
from sqlalchemy import UniqueConstraint

class Item(db.Model):
    __tablename__ = "item"  # nome real da tabela
    id   = db.Column(db.Integer, primary_key=True)
    carreta = db.Column(db.String(120))
    codigo  = db.Column(db.String(120))
    desc    = db.Column(db.Text)
    item_code = db.Column(db.String(120))
    qt        = db.Column(db.Integer)
    item_description = db.Column(db.Text)
    tipo   = db.Column(db.String(50))
    
    __table_args__ = (
        UniqueConstraint('carreta', 'item_code', name='uix_carreta_item_code'),
    )

    def as_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}

class Recurso(db.Model):
    __tablename__ = "recurso"  # nome real da tabela

    id   = db.Column(db.Integer, primary_key=True)
    carreta = db.Column(db.String(120))
    desc_carreta = db.Column(db.Text)
    recurso  = db.Column(db.String(120))
    descricao_recurso = db.Column(db.Text)
    qt = db.Column(db.Float)

    def as_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}
    
    
