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


user = Users.query.filter_by(access_token="b588c0cc074fd4dafd4692978bda7e43c34128922f48a3a818e0cbff3fe86e0d").first()
print(user.field_option)
user.subdomain = "d3v-sanjeevalgodom"
db.session.commit()
# new_token = Users(access_token="5334535", subdomain="sanjeev", selected_field="hallw", is_active=1,
#             field_option="thisjdf")
# db.session.add(new_token)
# db.session.commit()

#     def query_email(self, email):
#         user = Users.query.filter_by(email=email).first()
#         return user

# posts = Users.query.filter()    


# flag = 0
# users_obj = Users()
# data = request.get_json()
    
# hashed_password = generate_password_hash(data['password'], method='sha256')
# new_user = Users(name=data['name'], email=data['email'], password=hashed_password, created_at=datetime.datetime.utcnow(), is_active=1, is_deleted=0, is_admin=0, 
#                 updated_at=datetime.datetime.utcnow())
# db.session.add(new_user)

# new_user_id = users_obj.query_email(email = data['email']).id
# token = jwt.encode({'email': data['email']}, app.config['SECRET_KEY'], algorithm="HS256")
# new_token = UserApiKey(user_id=new_user_id, token=token.decode('UTF-8'), created_at=datetime.datetime.utcnow(),
#             updated_at=datetime.datetime.utcnow())
# db.session.add(new_token)

# new_user_access = UserAccessLimit(user_id=new_user_id, api_limit=api_per_user_limit, api_limit_count=0, add_limit=0,
#             updated_at=datetime.datetime.utcnow(), created_at=datetime.datetime.utcnow())
# db.session.add(new_user_access)
# flag = 1
# db.session.commit()


# user_accesss_limit_obj = UserAccessLimit()
# user_api_key_obj = UserApiKey()
# user_id = user_api_key_obj.fetch_token(token=request.headers['X-ACCESS-TOKEN']).user_id
# user_limit = user_accesss_limit_obj.access_limit(user_id=user_id)