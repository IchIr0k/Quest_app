from sqlalchemy.orm import Session
from sqlalchemy import and_
import models, schemas
from datetime import datetime

def get_quests(db: Session, skip: int = 0, limit: int = 12, filters: dict = None):
    query = db.query(models.Quest)

    if filters:
        # Поиск по тексту
        if filters.get("q"):
            q = f"%{filters['q']}%"
            query = query.filter(models.Quest.title.ilike(q))

        # Жанры
        if filters.get("genre"):
            if isinstance(filters["genre"], str):
                genres = [g.strip() for g in filters["genre"].split(",")]
            else:
                genres = filters["genre"]
            genre_filters = [models.Quest.genre.ilike(f"%{genre}%") for genre in genres]
            query = query.filter(and_(*genre_filters))

        # Сложность
        if filters.get("difficulty"):
            if isinstance(filters["difficulty"], str):
                difficulties = [d.strip() for d in filters["difficulty"].split(",")]
            else:
                difficulties = filters["difficulty"]
            query = query.filter(models.Quest.difficulty.in_(difficulties))

        # Уровень страха (>= выбранного)
        if filters.get("fear_level"):
            try:
                fear_level = int(filters["fear_level"])
                query = query.filter(models.Quest.fear_level >= fear_level)
            except ValueError:
                pass

        # Количество игроков (<= выбранного)
        if filters.get("players"):
            try:
                players = int(filters["players"])
                query = query.filter(models.Quest.players <= players)
            except ValueError:
                pass

    # Сортировка
    if filters and filters.get("sort"):
        sort = filters["sort"]
        if sort == "title_asc":
            query = query.order_by(models.Quest.title.asc())
        elif sort == "title_desc":
            query = query.order_by(models.Quest.title.desc())
        elif sort == "price_low":
            query = query.order_by(models.Quest.price.asc())
        elif sort == "price_high":
            query = query.order_by(models.Quest.price.desc())
    else:
        query = query.order_by(models.Quest.title.asc())

    return query.offset(skip).limit(limit).all()

def get_quest(db: Session, quest_id: int):
    return db.query(models.Quest).filter(models.Quest.id == quest_id).first()

def has_quest_bookings(db: Session, quest_id: int) -> bool:
    """Проверяет, есть ли у квеста активные бронирования"""
    booking_count = db.query(models.Booking).filter(
        models.Booking.quest_id == quest_id
    ).count()
    return booking_count > 0

def get_quest_bookings(db: Session, quest_id: int):
    """Получает все бронирования для конкретного квеста"""
    return db.query(models.Booking).filter(
        models.Booking.quest_id == quest_id
    ).join(models.User).order_by(models.Booking.date_time.desc()).all()

def delete_quest_bookings(db: Session, quest_id: int):
    """Удаляет все бронирования для квеста"""
    db.query(models.Booking).filter(models.Booking.quest_id == quest_id).delete()
    db.commit()
    return True

def delete_quest(db: Session, quest_id: int):
    quest = db.query(models.Quest).filter(models.Quest.id == quest_id).first()
    if quest:
        db.query(models.Booking).filter(models.Booking.quest_id == quest_id).delete()
        db.delete(quest)
        db.commit()
        return True
    return False

def get_booked_slots(db: Session, quest_id: int):
    """Получает все занятые слоты для квеста"""
    bookings = db.query(models.Booking).filter(models.Booking.quest_id == quest_id).all()
    return [booking.date_time for booking in bookings]

def get_booked_slots_for_date(db: Session, quest_id: int, date: str):
    """Получает занятые слоты для конкретной даты"""
    bookings = db.query(models.Booking).filter(
        models.Booking.quest_id == quest_id,
        models.Booking.date_time.like(f"{date}%")
    ).all()

    booked_times = [booking.date_time.split(" ")[1] for booking in bookings]
    return booked_times

def create_booking(db: Session, user_id: int, quest_id: int, date: str, timeslot: str):
    """Создает бронирование, если слот свободен"""
    date_time = f"{date} {timeslot}"

    # Проверяем, не занят ли слот этим же пользователем
    existing = db.query(models.Booking).filter(
        models.Booking.quest_id == quest_id,
        models.Booking.date_time == date_time,
        models.Booking.user_id == user_id
    ).first()

    if existing:
        return None  # Пользователь уже забронировал этот слот

    # Также проверяем, не занят ли слот вообще
    existing_any = db.query(models.Booking).filter(
        models.Booking.quest_id == quest_id,
        models.Booking.date_time == date_time
    ).first()

    if existing_any:
        return None  # Слот уже занят другим пользователем

    booking = models.Booking(
        user_id=user_id,
        quest_id=quest_id,
        date_time=date_time
    )

    db.add(booking)
    db.commit()
    db.refresh(booking)
    return booking

def get_user_bookings(db: Session, user_id: int):
    """Получает все бронирования пользователя"""
    return db.query(models.Booking).filter(
        models.Booking.user_id == user_id
    ).join(models.Quest).order_by(models.Booking.date_time.desc()).all()

def get_all_bookings(db: Session):
    """Получает все бронирования для администратора"""
    return db.query(models.Booking).join(models.User).join(models.Quest).order_by(
        models.Booking.date_time.desc()
    ).all()

def delete_booking(db: Session, booking_id: int):
    """Удаляет бронирование"""
    booking = db.query(models.Booking).filter(models.Booking.id == booking_id).first()
    if booking:
        db.delete(booking)
        db.commit()
        return True
    return False