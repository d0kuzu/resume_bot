from datetime import datetime
from typing import Annotated

from sqlalchemy import text, BigInteger, ForeignKey, DateTime, func, Boolean
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

intpk = Annotated[int, mapped_column(primary_key=True, autoincrement=True)]
created_at_pk = Annotated[
    datetime, mapped_column(server_default=text("TIMEZONE('utc', now())"))]
updated_at_pk = Annotated[
    datetime, mapped_column(server_default=text("TIMEZONE('utc', now())"),
                            onupdate=text("TIMEZONE('utc', now())"))]


class Base(DeclarativeBase):
    __table_args__ = {'extend_existing': True}


class User(Base):
    __tablename__ = "users"

    id: Mapped[intpk]
    user_id: Mapped[int] = mapped_column(BigInteger, nullable=True)
    username: Mapped[str] = mapped_column(nullable=True)
    fullname: Mapped[str] = mapped_column(nullable=True)
    affiliate: Mapped[str] = mapped_column(nullable=True)
    city: Mapped[str] = mapped_column(nullable=True)
    country: Mapped[str] = mapped_column(nullable=True)
    """
    guest: null in fullname, affiliate, city, phone_number
    no_access: have all datas. wait access from admin
    user: have all user privileges
    admin: have all privileges
    """
    role: Mapped[str] = mapped_column(default="guest")
    lang: Mapped[str] = mapped_column(default="ru")
    phone_number: Mapped[str] = mapped_column(nullable=True)
    created_at: Mapped[created_at_pk]
    updated_at: Mapped[updated_at_pk]

    def get_null_columns(self):
        result = []
        if not self.fullname:
            result.append("fullname")
        if not self.affiliate:
            result.append("affiliate")
        if not self.city:
            result.append("city")
        if not self.country:
            result.append("country")
        if not self.phone_number:
            result.append("phone_number")
        result.append("lang")
        return result


class Subscription(Base):
    __tablename__ = "subscriptions"

    id: Mapped[intpk]
    user_id = mapped_column(BigInteger)
    date_start = mapped_column(DateTime)
    date_end = mapped_column(DateTime)
    created_at = mapped_column(DateTime, server_default=func.now())
    is_warn = mapped_column(Boolean, default=False)
