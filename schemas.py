from pydantic import BaseModel, ConfigDict
from typing import Optional

class UserCreate(BaseModel):
    username: str
    email: Optional[str] = None
    password: str

    model_config = ConfigDict(from_attributes=True)

class UserOut(BaseModel):
    id: int
    username: str
    email: Optional[str]
    is_admin: bool

    model_config = ConfigDict(from_attributes=True)

class QuestBase(BaseModel):
    title: str
    description: str
    genre: str
    difficulty: str
    fear_level: int
    players: int

    model_config = ConfigDict(from_attributes=True)

class QuestCreate(QuestBase):
    pass

class QuestOut(QuestBase):
    id: int
    image_path: Optional[str] = None

    model_config = ConfigDict(from_attributes=True)

class BookingCreate(BaseModel):
    quest_id: int
    date: str
    timeslot: str

    model_config = ConfigDict(from_attributes=True)
