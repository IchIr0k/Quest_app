import os
import shutil
import uuid
from typing import Optional, List
from datetime import datetime

from fastapi import (
    FastAPI, Request, Form, UploadFile, File, Depends, HTTPException
)
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy.orm import Session
from sqlalchemy.exc import IntegrityError

from database import engine, Base, SessionLocal
import models
import crud
from auth import hash_password, verify_password, get_db, get_current_user, require_admin
from schemas import QuestCreate
import uvicorn

# --- Подготовка директорий ---
os.makedirs("static/uploads", exist_ok=True)
os.makedirs("static/images", exist_ok=True)

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key="!secret_dev_change_me!")

# --- Статика и шаблоны ---
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# --- Создание таблиц ---
Base.metadata.create_all(bind=engine)


# --- Создание дефолтного админа ---
def create_default_admin():
    db = SessionLocal()
    try:
        admin = db.query(models.User).filter_by(username="admin").first()
        if not admin:
            a = models.User(
                username="admin",
                email="admin@example.com",
                hashed_password=hash_password("admin"),
                is_admin=True
            )
            db.add(a)
            db.commit()
            print("✅ Created default admin (username=admin password=admin). Change password immediately.")
    finally:
        db.close()


create_default_admin()


# --- Хелперы ---
def save_upload(file: UploadFile) -> str:
    """Сохраняет файл в static/uploads и возвращает относительный путь"""
    ext = os.path.splitext(file.filename)[1]
    safe_name = f"{uuid.uuid4().hex}{ext}"
    dest = os.path.join("static", "uploads", safe_name)
    with open(dest, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return f"uploads/{safe_name}"


# --- Маршруты ---
@app.get("/", response_class=HTMLResponse)
def index(request: Request, q: Optional[str] = None, genre: Optional[str] = None,
          difficulty: Optional[str] = None, sort: Optional[str] = None,
          skip: int = 0, db: Session = Depends(get_db)):
    filters = {}
    if q:
        filters["q"] = q
    if genre:
        filters["genre"] = genre
    if difficulty:
        filters["difficulty"] = difficulty
    if sort:
        filters["sort"] = sort

    quests = crud.get_quests(db, skip=skip, limit=6, filters=filters)
    try:
        user = get_current_user(request, db)
    except:
        user = None
    return templates.TemplateResponse("index.html", {
        "request": request,
        "quests": quests,
        "user": user,
        "skip": skip,
        "now": datetime.now
    })


@app.get("/api/quests", response_class=HTMLResponse)
def api_quests(request: Request, q: Optional[str] = None, genre: Optional[str] = None,
               difficulty: Optional[str] = None, skip: int = 0, limit: int = 6, db: Session = Depends(get_db)):
    filters = {}
    if q:
        filters["q"] = q
    if genre:
        filters["genre"] = genre
    if difficulty:
        filters["difficulty"] = difficulty
    quests = crud.get_quests(db, skip=skip, limit=limit, filters=filters)
    return templates.TemplateResponse("_quest_cards.html", {"request": request, "quests": quests})


@app.get("/quest/{quest_id}", response_class=HTMLResponse)
def quest_detail(request: Request, quest_id: int, db: Session = Depends(get_db)):
    quest = crud.get_quest(db, quest_id)
    if not quest:
        raise HTTPException(status_code=404, detail="Quest not found")

    # Получаем занятые слоты для этого квеста
    booked_slots = crud.get_booked_slots_for_date(db, quest_id, datetime.now().strftime('%Y-%m-%d'))

    try:
        user = get_current_user(request, db)
    except:
        user = None

    return templates.TemplateResponse("quest_detail.html", {
        "request": request,
        "quest": quest,
        "user": user,
        "booked_slots": booked_slots,
        "now": datetime.now
    })


@app.get("/api/available-slots")
def get_available_slots(quest_id: int, date: str, db: Session = Depends(get_db)):
    """API для получения занятых слотов"""
    booked_slots = crud.get_booked_slots_for_date(db, quest_id, date)
    return JSONResponse(booked_slots)



@app.post("/book")
def book(request: Request, quest_id: int = Form(...), date: str = Form(...), timeslot: str = Form(...),
         db: Session = Depends(get_db)):
    user = get_current_user(request, db)
    booking = crud.create_booking(db, user_id=user.id, quest_id=quest_id, date=date, timeslot=timeslot)
    if not booking:
        return JSONResponse({"success": False, "message": "Выбранный слот уже занят"}, status_code=400)
    return JSONResponse({"success": True, "message": "Бронь успешно создана"})


@app.get("/my-bookings", response_class=HTMLResponse)
def my_bookings(request: Request, db: Session = Depends(get_db)):
    """Страница с бронированиями пользователя"""
    user = get_current_user(request, db)
    bookings = crud.get_user_bookings(db, user.id)
    return templates.TemplateResponse("my_bookings.html", {
        "request": request,
        "user": user,
        "bookings": bookings,
        "now": datetime.now
    })


# --- Auth routes ---
@app.get("/login", response_class=HTMLResponse)
def login_get(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})


@app.post("/login")
def login_post(request: Request, username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    user = db.query(models.User).filter_by(username=username).first()
    if not user or not verify_password(password, user.hashed_password):
        return templates.TemplateResponse("login.html", {"request": request, "error": "Неверные учётные данные"})
    request.session["user_id"] = user.id
    return RedirectResponse("/", status_code=303)


@app.get("/register", response_class=HTMLResponse)
def register_get(request: Request):
    return templates.TemplateResponse("register.html", {"request": request})


@app.post("/register")
def register_post(request: Request, username: str = Form(...), email: str = Form(None), password: str = Form(...),
                  db: Session = Depends(get_db)):
    # Проверяем, не существует ли пользователь с таким username
    existing_username = db.query(models.User).filter_by(username=username).first()
    if existing_username:
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "Пользователь с таким именем уже существует"
        })

    # Если email указан, проверяем его уникальность
    if email:
        existing_email = db.query(models.User).filter_by(email=email).first()
        if existing_email:
            return templates.TemplateResponse("register.html", {
                "request": request,
                "error": "Пользователь с таким email уже существует"
            })

    # Проверяем длину пароля
    if len(password) < 4:
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "Пароль должен содержать минимум 4 символа"
        })

    # Проверяем длину имени пользователя
    if len(username) < 3:
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "Имя пользователя должно содержать минимум 3 символа"
        })

    try:
        u = models.User(
            username=username,
            email=email if email else None,  # Сохраняем как None если email пустой
            hashed_password=hash_password(password),
            is_admin=False
        )
        db.add(u)
        db.commit()
        db.refresh(u)
        request.session["user_id"] = u.id
        return RedirectResponse("/", status_code=303)

    except IntegrityError:
        db.rollback()
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "Произошла ошибка при создании пользователя. Попробуйте другое имя или email."
        })


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/", status_code=303)


# --- Admin routes ---
@app.get("/admin", response_class=HTMLResponse)
def admin_dashboard(request: Request, db: Session = Depends(get_db), user=Depends(require_admin)):
    quests = crud.get_quests(db, skip=0, limit=1000, filters={})
    return templates.TemplateResponse("admin_dashboard.html", {
        "request": request,
        "quests": quests,
        "user": user,
        "now": datetime.now
    })


@app.get("/admin/add", response_class=HTMLResponse)
def add_get(request: Request, user=Depends(require_admin)):
    return templates.TemplateResponse("add_quest.html", {"request": request, "user": user})


@app.post("/admin/add")
def add_post(request: Request,
             title: str = Form(...),
             description: str = Form(""),
             genres: List[str] = Form(...),
             difficulty: str = Form(""),
             fear_level: int = Form(0),
             players: int = Form(1),
             price: int = Form(2000),  # Добавлено
             image: Optional[UploadFile] = File(None),
             db: Session = Depends(get_db),
             user=Depends(require_admin)):
    image_path = None
    if image and image.filename:
        image_path = save_upload(image)

    genre_str = ", ".join(genres)

    new_quest = models.Quest(
        title=title,
        description=description,
        genre=genre_str,
        difficulty=difficulty,
        fear_level=fear_level,
        players=players,
        price=price,  # Добавлено
        image_path=image_path
    )

    db.add(new_quest)
    db.commit()
    db.refresh(new_quest)

    return RedirectResponse("/admin", status_code=303)


@app.get("/admin/edit/{quest_id}", response_class=HTMLResponse)
def edit_get(request: Request, quest_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    quest = crud.get_quest(db, quest_id)
    if not quest:
        raise HTTPException(404, "Квест не найден")
    return templates.TemplateResponse("edit_quest.html", {"request": request, "quest": quest, "user": user})


@app.post("/admin/edit/{quest_id}")
def edit_post(request: Request, quest_id: int,
              title: str = Form(...),
              description: str = Form(""),
              genres: List[str] = Form(...),
              difficulty: str = Form(""),
              fear_level: int = Form(0),
              players: int = Form(1),
              price: int = Form(2000),  # Добавлено
              image: Optional[UploadFile] = File(None),
              db: Session = Depends(get_db),
              user=Depends(require_admin)):
    quest = crud.get_quest(db, quest_id)
    if not quest:
        raise HTTPException(404, "Квест не найден")

    # Обновляем данные
    quest.title = title
    quest.description = description
    quest.genre = ", ".join(genres)
    quest.difficulty = difficulty
    quest.fear_level = fear_level
    quest.players = players
    quest.price = price  # Добавлено

    # Обновляем изображение если загружено новое
    if image and image.filename:
        if quest.image_path:
            old_path = os.path.join("static", quest.image_path)
            if os.path.exists(old_path):
                os.remove(old_path)
        quest.image_path = save_upload(image)

    db.commit()

    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/delete/{quest_id}")
def admin_delete(request: Request, quest_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    # Проверяем, есть ли бронирования у этого квеста
    has_bookings = crud.has_quest_bookings(db, quest_id)

    if has_bookings:
        # Если есть бронирования, возвращаем ошибку
        quests = crud.get_quests(db, skip=0, limit=1000, filters={})
        return templates.TemplateResponse("admin_dashboard.html", {
            "request": request,
            "quests": quests,
            "user": user,
            "error": f"Невозможно удалить квест: есть активные бронирования. Сначала удалите все бронирования."
        })

    # Если бронирований нет, удаляем квест
    q = crud.get_quest(db, quest_id)
    if q and q.image_path:
        file_path = os.path.join("static", q.image_path)
        if os.path.exists(file_path):
            os.remove(file_path)
    crud.delete_quest(db, quest_id)
    return RedirectResponse("/admin", status_code=303)


@app.get("/admin/bookings", response_class=HTMLResponse)
def admin_bookings(request: Request, quest_id: Optional[int] = None, db: Session = Depends(get_db),
                   user=Depends(require_admin)):
    """Страница управления бронированиями для администратора"""
    if quest_id:
        bookings = crud.get_quest_bookings(db, quest_id)
    else:
        bookings = crud.get_all_bookings(db)

    # Получаем все квесты для фильтра
    quests = crud.get_quests(db, skip=0, limit=1000, filters={})

    return templates.TemplateResponse("admin_bookings.html", {
        "request": request,
        "bookings": bookings,
        "quests": quests,
        "user": user,
        "now": datetime.now
    })


@app.post("/admin/delete/{quest_id}")
def admin_delete(request: Request, quest_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    # Проверяем, есть ли бронирования у этого квеста
    has_bookings = crud.has_quest_bookings(db, quest_id)

    if has_bookings:
        # Получаем информацию о квесте и его бронированиях
        quest = crud.get_quest(db, quest_id)
        quest_bookings = crud.get_quest_bookings(db, quest_id)
        quests = crud.get_quests(db, skip=0, limit=1000, filters={})

        return templates.TemplateResponse("admin_dashboard.html", {
            "request": request,
            "quests": quests,
            "user": user,
            "error": f"Невозможно удалить квест '{quest.title}': есть активные бронирования ({len(quest_bookings)} шт.).",
            "blocked_quest_id": quest_id,
            "quest_bookings": quest_bookings
        })

    # Если бронирований нет, удаляем квест
    q = crud.get_quest(db, quest_id)
    if q and q.image_path:
        file_path = os.path.join("static", q.image_path)
        if os.path.exists(file_path):
            os.remove(file_path)
    crud.delete_quest(db, quest_id)
    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/delete-quest-with-bookings/{quest_id}")
def admin_delete_quest_with_bookings(quest_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    """Удаляет квест вместе со всеми его бронированиями"""
    # Сначала удаляем все бронирования квеста
    crud.delete_quest_bookings(db, quest_id)

    # Затем удаляем сам квест
    q = crud.get_quest(db, quest_id)
    if q and q.image_path:
        file_path = os.path.join("static", q.image_path)
        if os.path.exists(file_path):
            os.remove(file_path)
    crud.delete_quest(db, quest_id)

    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/delete-all-bookings/{quest_id}")
def admin_delete_all_bookings(quest_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    """Удаляет все бронирования квеста без удаления самого квеста"""
    crud.delete_quest_bookings(db, quest_id)
    return RedirectResponse("/admin", status_code=303)

@app.post("/admin/delete-booking/{booking_id}")
def admin_delete_booking(booking_id: int, db: Session = Depends(get_db), user=Depends(require_admin)):
    """Удаление бронирования администратором"""
    success = crud.delete_booking(db, booking_id)
    if not success:
        return JSONResponse({"success": False, "message": "Бронирование не найдено"}, status_code=404)

    return RedirectResponse("/admin/bookings", status_code=303)

@app.get("/api/quest-has-bookings/{quest_id}")
def api_quest_has_bookings(quest_id: int, db: Session = Depends(get_db)):
    """API для проверки наличия бронирований у квеста"""
    has_bookings = crud.has_quest_bookings(db, quest_id)
    return JSONResponse({"has_bookings": has_bookings})


if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=5000,
        reload=True
    )