import json
import os

from aiogram import Router, F
from aiogram.fsm.context import FSMContext
from aiogram.types import Message, FSInputFile

from services.doc.doc import fill_document
from .config import start_text
from database.database import ORM
from services.gpt.gpt import read_docx_with_tables, send_to_gpt

router = Router()

@router.message(F.document)
async def photo_analyze(message: Message):
    temp_path = f"data/tmp/"
    filename = f'{message.document.file_name}'
    await message.bot.download(file=message.document.file_id, destination=temp_path+filename)

    msg = await message.answer('Бот анализирует ваше резюме')

    text = read_docx_with_tables(temp_path+filename)

    answer: dict = send_to_gpt(start_text+'\n'+text)

    answer_text = ''
    for k, v in list(answer.items()):
        answer_text += f'{k} - {v} \n\n'

    try:
        await msg.delete()
        await message.answer(answer_text)
    except:
        ...

    os.remove(temp_path+filename)

    example_path = f"data/"
    output_file = temp_path+'FIRCaspian CV_заполненный.docx'
    fill_document(example_path+'FIRCaspian CV RU_ новый шаблон.docx', answer, output_file)

    document = FSInputFile(output_file)
    await message.bot.send_document(message.from_user.id, document)

    os.remove(output_file)


@router.message(F.text)
async def sui(message: Message):
    print('\n\n\n\n\n\n\n')
    print(message.text)
    print(message.from_user.first_name)
    print(message.from_user.last_name)
    print(message.from_user.id)
    print('\n\n\n\n\n\n\n')
