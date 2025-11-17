from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy.orm import DeclarativeBase
from sqlalchemy import Column, Integer, String
from sqlalchemy import delete


# Пользователь dmitriy, пароль: password, ip: localhost, порт: 5432, база данных: marketplaces
engine = create_engine('postgresql://dmitriy:passw@localhost:5432/market', echo=False)


class Base(DeclarativeBase):
    pass

# Создадим модель
class Books(Base):
    __tablename__ = 'books_program'
    id = Column(Integer, primary_key=True, autoincrement=True)
    title = Column(String)
    url = Column(String)
    price = Column(Integer)
    url_img = Column(String)


# создаем базу данных market и в ней таблицу books_program
Base.metadata.create_all(bind=engine)


def write_db(lst_data):
    """
    Удалим старые данные, чтобы они не дублировались.
    Новые данные из списка lst_data запишем в таблицу books_program
    """
    with Session(autoflush=False, bind=engine) as db:
        # Очистим таблицу books_program, если она существует
        print("++++++==================================================")
        for t in Base.metadata.sorted_tables:
            print(t.name)
            if t.name == 'books_program':
                db.execute(delete(t))
        print("++++++==================================================")
                
        for book in lst_data:
            db.add(Books(title = book['title'], url = book['url'], price = book['price'], url_img = book['img_url']))

        db.commit()
