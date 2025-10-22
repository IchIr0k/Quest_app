from sqlalchemy import Column, Integer, String, Boolean, Text, ForeignKey
from sqlalchemy.orm import relationship
from database import Base

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False)
    email = Column(String(120), unique=True, nullable=True)
    hashed_password = Column(String(255), nullable=False)
    is_admin = Column(Boolean, default=False)

    bookings = relationship("Booking", back_populates="user")

class Quest(Base):
    __tablename__ = "quests"
    id = Column(Integer, primary_key=True, index=True)
    title = Column(String(150), nullable=False)
    description = Column(Text, nullable=False)
    genre = Column(String(50), nullable=False)
    difficulty = Column(String(30), nullable=False)
    fear_level = Column(Integer, nullable=False)
    players = Column(Integer, nullable=False)
    price = Column(Integer, nullable=False, default=2000)  # Добавлено поле цены
    organizer_email = Column(String(120), nullable=False, default="alibi@mail.ru")  # Email организатора
    image_path = Column(String(255), nullable=True)

    bookings = relationship("Booking", back_populates="quest")

class Booking(Base):
    __tablename__ = "bookings"
    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"))
    quest_id = Column(Integer, ForeignKey("quests.id"))
    date_time = Column(String(50))

    user = relationship("User", back_populates="bookings")
    quest = relationship("Quest", back_populates="bookings")