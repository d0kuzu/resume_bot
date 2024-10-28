from aiogram import Dispatcher
from aiogram.utils.i18n import SimpleI18nMiddleware
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from database.database import ORM
from services.telegram.handlers.gpt import gpt
from services.telegram.middlewares.data import DataMiddleware


class TgRegister:
    def __init__(self, dp: Dispatcher, orm: ORM):
        self.dp = dp
        self.orm = orm
        # self.i18n = i18n

    def register(self):
        self._register_handlers()
        self._register_middlewares()

    def _register_handlers(self):
        # home
        self.dp.include_routers(gpt.router)

    def _register_middlewares(self):
        scheduler = AsyncIOScheduler(timezone="Asia/Almaty")
        scheduler.start()
        middleware = DataMiddleware(self.orm, scheduler)
        # i18n_middleware = SimpleI18nMiddleware(self.i18n, "i18n", "i18n_middleware")

        self.dp.callback_query.middleware(middleware)
        # self.dp.callback_query.middleware(i18n_middleware)
        self.dp.message.middleware(middleware)
        # self.dp.message.middleware(i18n_middleware)
        self.dp.inline_query.middleware(middleware)
        # self.dp.inline_query.middleware(i18n_middleware)