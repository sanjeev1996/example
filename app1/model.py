from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
SQLALCHEMY_DATABASE_URI = "postgres://rpcsnptislezje:86d21628161c84613c6728214fab05bd7bb80d1e8bfae32e5cbd77967b4f5a58@ec2-54-225-190-241.compute-1.amazonaws.com:5432/d118bdvdfhh2s5"
app.config['SQLALCHEMY_DATABASE_URI']=SQLALCHEMY_DATABASE_URI


db = SQLAlchemy(app)

class Users(db.Model):
    __tablename__ = "my_table1"
    id = db.Column('id', db.INT, primary_key=True)
    access_token = db.Column('access_token', db.VARCHAR)
    subdomain = db.Column('subdomain', db.VARCHAR)
    api_key = db.Column('api_key', db.VARCHAR)
    is_active = db.Column('is_active', db.INT)
    selected_field = db.Column('selected_field', db.VARCHAR)
    model_option = db.Column('model_option', db.VARCHAR)
    field_option = db.Column('field_option', db.VARCHAR)