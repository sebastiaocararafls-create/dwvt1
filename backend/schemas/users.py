from pydantic import BaseModel, Field
from typing import Literal

Role = Literal["admin", "engenharia", "comercial"]

class UserCreate(BaseModel):
    username: str = Field(min_length=3, max_length=50)
    password: str = Field(min_length=6, max_length=128)
    role: Role = "engenharia"
    is_active: bool = True

class UserOut(BaseModel):
    id: int
    username: str
    role: str
    is_active: bool

    class Config:
        from_attributes = True