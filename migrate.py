from sqlalchemy import text
from database import engine, SessionLocal


def migrate_database():
    """Добавляет отсутствующие колонки в существующую базу данных"""
    db = SessionLocal()

    try:
        # Проверяем существование колонки price
        result = db.execute(text("""
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_name='quests' AND column_name='price'
        """))

        if not result.fetchone():
            print("Добавляем колонку price...")
            db.execute(text("ALTER TABLE quests ADD COLUMN price INTEGER DEFAULT 2000"))
            print("✅ Колонка price добавлена")
        else:
            print("✅ Колонка price уже существует")

        # Проверяем существование колонки created_at
        result = db.execute(text("""
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_name='quests' AND column_name='created_at'
        """))

        if not result.fetchone():
            print("Добавляем колонку created_at...")
            db.execute(text("ALTER TABLE quests ADD COLUMN created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()"))
            print("✅ Колонка created_at добавлена")
        else:
            print("✅ Колонка created_at уже существует")

        db.commit()
        print("✅ Миграция завершена успешно!")

    except Exception as e:
        db.rollback()
        print(f"❌ Ошибка миграции: {e}")
    finally:
        db.close()


if __name__ == "__main__":
    migrate_database()