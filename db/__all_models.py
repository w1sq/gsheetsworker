from email.policy import default
import sqlalchemy
from .db_session import SqlAlchemyBase
from datetime import datetime

class Users(SqlAlchemyBase):
    __tablename__ = "users"

    id = sqlalchemy.Column(sqlalchemy.Integer, nullable=False, primary_key=True)
    muted_notifications = sqlalchemy.Column(sqlalchemy.String, default='')

class Notifications(SqlAlchemyBase):
    __tablename__ = 'notifications'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key = True, autoincrement = True)
    text = sqlalchemy.Column(sqlalchemy.String, default='')
    date_added = sqlalchemy.Column(sqlalchemy.DateTime, default=datetime.now())


class Limits(SqlAlchemyBase):
    __tablename__ = 'limits'

    id = sqlalchemy.Column(sqlalchemy.Integer, primary_key = True, autoincrement = True)
    warehouse = sqlalchemy.Column(sqlalchemy.Integer, nullable=False)
    type = sqlalchemy.Column(sqlalchemy.String, nullable=False)
    amount = sqlalchemy.Column(sqlalchemy.Integer, nullable=False)
    time_range = sqlalchemy.Column(sqlalchemy.DateTime, default = datetime.now())