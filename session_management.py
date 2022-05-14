import datetime
from peewee import *

db = SqliteDatabase('sessions.db')

class BaseModel(Model):
    class Meta:
        database = db


class Session(BaseModel):
    #the session id is an auto-incrementing integer
    session_id = AutoField()
    session_name = CharField()

    class Meta:
        db_table = 'sessions'


class Word(BaseModel):
    #the word id is an auto-incrementing integer
    word_id = AutoField()
    word = CharField()

    class Meta:
        db_table = 'words'


# Word class and Session clas have a many to many relationship
# the join table is called 'session_words'
class SessionWords(BaseModel):
    session_id = ForeignKeyField(Session, related_name='session_words')
    word_id = ForeignKeyField(Word, related_name='session_words')

    class Meta:
        db_table = 'session_words'
        primary_key = CompositeKey('session_id', 'word_id')


def initialize(database):
    database.connect()
    database.create_tables([Session, Word, SessionWords])
    return database


def save_word_in_session(session_name, word_name):
    try:
        session, _= Session.get_or_create(session_name=session_name)
        word, _ = Word.get_or_create(word=word_name)
        SessionWords.create(session_id=session.session_id, word_id=word.word_id)
    except IntegrityError:
        pass


def get_session_words(session):
    try:
        session_id = Session.get(Session.session_name == session).session_id
        return [word.word for word in Word.select().join(SessionWords).where(SessionWords.session_id == session_id)]
    except DoesNotExist:
        return []


def is_word_in_session(session, word):
    session_id = Session.get(Session.session_name == session).session_id
    word_id = Word.get(Word.word == word).word_id
    return SessionWords.select().where(SessionWords.session_id == session_id, SessionWords.word_id == word_id).exists()


if __name__ == '__main__':
    db = initialize(db)
    save_word_in_session('test', 'test_word')
    print(is_word_in_session('test', 'test_word'))
    print(get_session_words('test'))
