from db.__all_models import Users, Notifications
from db.db_session import global_init, create_session
global_init()
db_sess = create_session()
users = db_sess.query(Users).all()
for user in users:
    print(user.id)