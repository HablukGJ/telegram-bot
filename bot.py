import asyncio
import os
import pandas as pd
import logging

from aiogram import Dispatcher, Bot, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.types import Message, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery

token = "8471510827:AAHZwNtBBKTRJ39p4FV7C4YPZ26gFhDugVs"
dp = Dispatcher()
bot = Bot(token=token)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

NAME, EMAIL, PHONE = range(3)
EXCEL_FILE = "users_data.xlsx"

def user_exists(user_id: str) -> bool:
    """Проверяет, зарегистрирован ли пользователь уже"""
    try:
        if not os.path.exists(EXCEL_FILE):
            return False

        df = pd.read_excel(EXCEL_FILE)
        user_id_str = str(user_id)
        return user_id_str in df['user_id'].astype(str).values

    except Exception as e:
        logger.error(f"❌ Ошибка при проверке пользователя: {e}")
        return False

def get_user_data(user_id: str) -> dict:
    """Получает данные пользователя по ID"""
    try:
        if not os.path.exists(EXCEL_FILE):
            return None

        df = pd.read_excel(EXCEL_FILE)
        user_id_str = str(user_id)
        user_row = df[df['user_id'].astype(str) == user_id_str]

        if not user_row.empty:
            return user_row.iloc[0].to_dict()
        return None

    except Exception as e:
        logger.error(f"❌ Ошибка при получении данных пользователя: {e}")
        return None

def save_to_excel(user_data: dict):
    """Сохранение данных в Excel с помощью pandas"""
    try:
        # Создаем DataFrame из данных пользователя
        df_new = pd.DataFrame([user_data])

        # Если файл существует, читаем его и добавляем новые данные
        if os.path.exists(EXCEL_FILE):
            df_existing = pd.read_excel(EXCEL_FILE)

            # Проверяем, нет ли уже пользователя с таким ID
            user_id_str = str(user_data['user_id'])
            if user_id_str in df_existing['user_id'].astype(str).values:
                # Обновляем существующую запись
                user_index = df_existing.index[df_existing['user_id'].astype(str) == user_id_str].tolist()
                for key, value in user_data.items():
                    df_existing.at[user_index[0], key] = value
                df_combined = df_existing
            else:
                # Добавляем нового пользователя
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        # Сохраняем в Excel с авто-шириной столбцов
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, sheet_name='Users')

            # Получаем лист для настройки ширины столбцов
            worksheet = writer.sheets['Users']

            # Автоматически подбираем ширину столбцов
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass

                # Устанавливаем ширину столбца с запасом
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        logger.info(f"✅ Данные сохранены в Excel: {user_data['full_name']}")
        return True

    except Exception as e:
        logger.error(f"❌ Ошибка при сохранении в Excel: {e}")
        return False

inline_events_1 = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text='Разговоры о важном с директором', callback_data='1_0')],
    [InlineKeyboardButton(text='круглый стол "Как привить ребенку интерес к чтению', callback_data='1_1')],
    [InlineKeyboardButton(text='мастер-класс по русскому языку', callback_data='1_2')],
    [InlineKeyboardButton(text='мастер-класс по истории', callback_data='1_3')],
    [InlineKeyboardButton(text='мастер-класс по китайскому языку', callback_data='1_4')],
    [InlineKeyboardButton(text='мастер-класс по ИЗО', callback_data='1_5')],
    [InlineKeyboardButton(text='мастер-класс по химии', callback_data='1_6')],
    [InlineKeyboardButton(text='мастер-класс по английскому языку', callback_data='1_7')],
    [InlineKeyboardButton(text='мастер-класс по информатике', callback_data='1_8')],
    [InlineKeyboardButton(text='мастер-класс по робототехники', callback_data='1_9')]
])

inline_events_2 = InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text='квиз "60 ярких лет"', callback_data='2_0')],
    [InlineKeyboardButton(text='круглый стол "Как привить ребенку интерес к чтению', callback_data='2_1')],
    [InlineKeyboardButton(text='Практическое занятие с психологом "Управление детскими эмоциями"', callback_data='2_2')],
    [InlineKeyboardButton(text='«Что не любит сердце? Риски болезни сердца и пути профилактики»', callback_data='2_3')],
    [InlineKeyboardButton(text='«Школьное путешествие: прихоть или стратегия', callback_data='2_4')],
    [InlineKeyboardButton(text='Пишем ЕГЭ', callback_data='2_5')],
])

block1 = ['Разговоры о важном с директором',
          'круглый стол "Как привить ребенку интерес к чтению',
          'мастер-класс по русскому языку',
          'мастер-класс по истории',
          'мастер-класс по китайскому языку',
          'мастер-класс по ИЗО',
          'мастер-класс по химии',
          'мастер-класс по английскому языку',
          'мастер-класс по информатике',
          'мастер-класс по робототехники']

block2 = ['квиз "60 ярких лет',
          'Круглый стол «Семейный совет» встреча с А.М.Тереховым',
          'Практическое занятие с психологом "Управление детскими эмоциями"',
          '«Что не любит сердце? Риски болезни сердца и пути профилактики»',
          '«Школьное путешествие: прихоть или стратегия',
          'Пишем ЕГЭ']

class Register(StatesGroup):
    id = State()
    username = State()
    telegram_name = State()
    full_name = State()
    email = State()
    phone = State()
    event_1 = State()
    event_2 = State()

@dp.callback_query(F.data.startswith("1_"))
async def event_1_select(callback: CallbackQuery, state: FSMContext):
    await state.update_data(event_1=block1[int(callback.data.split('_')[1])])
    await callback.message.answer(f'Теперь выбери мероприятие из второго блока: ', reply_markup=inline_events_2)
    await state.set_state(Register.event_2)


@dp.callback_query(F.data.startswith("2_"))
async def event_2_select(callback: CallbackQuery, state: FSMContext):
    await state.update_data(event_2=block2[int(callback.data.split('_')[1])])
    await callback.message.answer("Спасибо за выбор!")
    data = await state.get_data()
    save_to_excel({
        'user_id': data['id'],
        'username': data['username'],
        'telegram_name': data['telegram_name'],
        'full_name': data['full_name'],
        'email': data['email'],
        'phone': data['phone'],
        'event_1': data['event_1'],
        'event_2': data['event_2']
    })
    await state.clear()


@dp.message(CommandStart())
async def start_bot(message: Message, state: FSMContext):
    user = message.from_user
    await message.answer(f"Привет, {user.first_name}!\n"
                         f"Давай зарегистрируем тебя. Введи свое ФИО: ")
    await state.update_data(id=user.id, username=user.username, telegram_name=user.first_name)
    await state.set_state(Register.full_name)


@dp.message(Register.full_name)
async def collect_name(message: Message, state: FSMContext):
    await state.update_data(full_name=message.text)
    await message.answer("Теперь введи почту: ")
    await state.set_state(Register.email)


@dp.message(Register.email)
async def collect_email(message: Message, state: FSMContext):
    await state.update_data(email=message.text)
    await message.answer("Нам еще понадобится контактный номер: ")
    await state.set_state(Register.phone)


@dp.message(Register.phone)
async def collect_phone(message: Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer("Регистрация завершена! Выберите блок мероприятий: ", reply_markup=inline_events_1)
    await state.set_state(Register.event_1)


async def main():
     await dp.start_polling(bot)

if __name__ == '__main__':
    asyncio.run(main())