import os
import logging
import asyncio
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder, CommandHandler, CallbackQueryHandler,
    MessageHandler, ContextTypes, filters, ConversationHandler
)
from dotenv import load_dotenv
from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
import shutil
from datetime import datetime, date, timedelta
import calendar

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# Загрузка переменных окружения
load_dotenv()

# Переменные окружения для бота и почты
BOT_TOKEN = os.getenv("BOT_TOKEN")
EMAIL_LOGIN = os.getenv("EMAIL_LOGIN")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER") # Это теперь будет использоваться как почта для копии (CC) или резервная
TEMPLATE_PATH = "template.xlsx" # Убедитесь, что template.xlsx существует в той же директории

# === НАЧАЛО ИЗМЕНЕНИЙ ===
# Почты специалистов
SPECIALIST_EMAIL_ALESYA = "a.zabavskaya@vds.by"
SPECIALIST_EMAIL_DMITRY = "d.karp@vds.by"
CC_EMAIL = "bas@vds.by" # Почта для копии

# Состояния для ConversationHandler
# Обновлено количество состояний - добавлено SPECIALIST_SELECTION
PROJECT, OBJECT, NAME, UNIT, QUANTITY, MODULE, POSITION_DELIVERY_DATE, \
ATTACHMENT_CHOICE, FILE_INPUT, LINK_INPUT, \
CONFIRM_ADD_MORE, \
EDIT_MENU, SELECT_POSITION, EDIT_FIELD_SELECTION, EDIT_FIELD_INPUT, \
FINAL_CONFIRMATION, GLOBAL_DELIVERY_DATE_SELECTION, \
EDITING_UNIT, EDITING_MODULE, SPECIALIST_SELECTION = range(20) # Изменили range на 20
# === КОНЕЦ ИЗМЕНЕНИЙ ===


# Глобальные переменные для хранения данных пользователя и предварительно определенных списков
user_state = {}
projects = ["Stadler", "Мотели"]
objects = ["Мерке", "Аральск", "Атырау", "Каркаролинск", "Семипалатинск"]
modules = [f"{i+1}" for i in range(18)]
units = ["м", "м2", "м3", "шт", "компл", "л", "кг", "тн"]

def fill_excel(project, object_name, positions, user_full_name, telegram_id_or_username):
    """
    Заполняет Excel-файл данными, включая дату поставки для каждой позиции, проект, объект,
    а также информацию о пользователе, от которого пришла заявка.
    """
    today = datetime.today().strftime("%d.%m.%Y")
    sanitized_user_name = user_full_name.replace(" ", "_")
    filename = f"Заявка_{project}_{object_name}_{sanitized_user_name}_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    output_dir = "out"

    template_full_path = os.path.abspath(TEMPLATE_PATH)

    os.makedirs(output_dir, exist_ok=True)
    new_path = os.path.join(output_dir, filename)

    shutil.copy(template_full_path, new_path)
    wb = load_workbook(new_path)
    ws = wb.active

    ws['G2'] = today
    ws['G3'] = project
    ws['G4'] = object_name
    ws['E16'] = user_full_name
    ws['E17'] = telegram_id_or_username

    logger.info(f"Writing to Excel: G2={today}, G3={project}, G4={object_name}, E16={user_full_name}, E17={telegram_id_or_username}")


    row_start_data = 9
    for i, pos in enumerate(positions):
        row = row_start_data + i

        ws.cell(row=row, column=1).value = i + 1
        ws.cell(row=row, column=2).value = pos["name"]
        ws.cell(row=row, column=3).value = pos["unit"]
        ws.cell(row=row, column=4).value = pos["quantity"]
        ws.cell(row=row, column=5).value = pos.get("delivery_date", "Не указано")
        ws.cell(row=row, column=6).value = pos["module"]
        ws.cell(row=row, column=7).value = pos.get("link", "")
        logger.info(f"Writing position {i+1} to Excel: {pos}")


    wb.save(new_path)
    logger.info(f"Excel file saved to: {new_path}")
    return new_path

# === НАЧАЛО ИЗМЕНЕНИЙ ===
async def send_email(chat_id, project, object_name, positions, user_full_name, telegram_id_or_username, to_email, cc_email, context=None):
    """
    Отправляет сгенерированный Excel-файл по электронной почте,
    с возможностью прикрепления дополнительных файлов и ссылок, привязанных к позициям,
    а также информацией о пользователе.
    """
    msg = EmailMessage()
    msg["Subject"] = f"Заявка на снабжение: {project} - {object_name}"
    msg["From"] = EMAIL_LOGIN
    msg["To"] = to_email # Динамический получатель
    msg["Cc"] = cc_email # Копия

    email_body = "Во вложении заявка на снабжение.\n\n"
    email_body += f"Проект: {project}\n"
    email_body += f"Объект: {object_name}\n"
    email_body += f"От кого: {user_full_name}\n"
    email_body += f"Telegram ID: {telegram_id_or_username}\n\n"
    email_body += "Позиции:\n"

    files_to_attach = []
    links_in_email = []

    for i, p in enumerate(positions):
        pos_info = (
            f"{i+1}. Модуль: {p.get('module', 'N/A')} | Наименование: {p.get('name', 'N/A')} | "
            f"Ед.изм.: {p.get('unit', 'N/A')} | Количество: {p.get('quantity', 'N/A')} | "
            f"Дата поставки: {p.get('delivery_date', 'N/A')}"
        )
        if p.get('link'):
            pos_info += f" | Ссылка: {p['link']}"
            links_in_email.append(f"Позиция {i+1} ({p.get('name', 'N/A')}): {p['link']}")

        if p.get('file_data') and isinstance(p['file_data'], list):
            file_names = []
            for file_item in p['file_data']:
                file_names.append(file_item.get('file_name', 'N/A'))
                files_to_attach.append((i+1, file_item))
            if file_names:
                pos_info += f" | Файлы: {', '.join(file_names)}"
        email_body += pos_info + "\n"

    if links_in_email:
        email_body += "\nОтдельные ссылки для позиций:\n" + "\n".join(links_in_email) + "\n"

    msg.set_content(email_body)
    logger.info(f"Email body generated for chat_id {chat_id}: \n{email_body}")

    try:
        file_path = fill_excel(project, object_name, positions, user_full_name, telegram_id_or_username)
        with open(file_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=os.path.basename(file_path),
            )
        logger.info(f"Excel file '{file_path}' прикреплен к письму")
    except Exception as e:
        logger.error(f"Ошибка при создании или прикреплении Excel файла: {e}")

    if context:
        for pos_index, file_data in files_to_attach:
            try:
                file_id = file_data['file_id']
                file_name = file_data['file_name']
                mime_type = file_data['mime_type']

                telegram_file = await context.bot.get_file(file_id)
                file_bytes = await telegram_file.download_as_bytearray()

                msg.add_attachment(
                    file_bytes,
                    maintype=mime_type.split('/')[0],
                    subtype=mime_type.split('/')[1],
                    filename=f"Позиция_{pos_index}_{file_name}",
                )
                logger.info(f"Дополнительный файл '{file_name}' для позиции {pos_index} прикреплен к письму")
            except Exception as e:
                logger.error(f"Ошибка при скачивании или прикреплении файла '{file_name}' для позиции {pos_index}: {e}")
                msg.set_content(msg.get_content() + f"\n\nВнимание: Не удалось прикрепить файл '{file_name}' для позиции {pos_index} из-за ошибки: {e}")


    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_LOGIN, EMAIL_PASSWORD)
            server.send_message(msg)
        logger.info(f"Письмо успешно отправлено на {to_email} с копией {cc_email}") # Изменено логирование
        return True
    except Exception as e:
        logger.error(f"Ошибка при отправке письма: {e}")
        raise

# === КОНЕЦ ИЗМЕНЕНИЙ ===


# === Telegram Handlers ===

async def initial_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает первое сообщение от пользователя (или когда диалог неактивен)
    и предлагает кнопку "Создать заявку".
    """
    keyboard = [[KeyboardButton("Создать заявку")]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=False, resize_keyboard=True)
    await update.message.reply_text(
        "Привет! Я бот для создания заявок. Нажмите кнопку, чтобы начать.",
        reply_markup=reply_markup
    )

async def start_conversation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Начинает новый разговор (после нажатия кнопки "Создать заявку")
    и предлагает пользователю выбрать проект.
    """
    chat_id = update.effective_chat.id
    user = update.effective_user

    first_name = user.first_name if user.first_name else ""
    last_name = user.last_name if user.last_name else ""
    user_full_name = f"{first_name} {last_name}".strip()
    telegram_id_or_username = user.username if user.username else str(user.id)

    user_state[chat_id] = {
        "user_full_name": user_full_name,
        "telegram_id_or_username": telegram_id_or_username,
        "project": None,
        "object": None,
        "positions": [],
    }
    logger.info(f"User {user_full_name} ({telegram_id_or_username}) started conversation.")

    await update.message.reply_text("Начинаем создание заявки...", reply_markup=ReplyKeyboardRemove())

    keyboard = [[InlineKeyboardButton(p, callback_data=p)] for p in projects]
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Выберите проект:", reply_markup=reply_markup)
    return PROJECT

async def project_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор проекта и предлагает выбрать объект."""
    query = update.callback_query
    await query.answer()

    user_state[query.message.chat.id]["project"] = query.data
    logger.info(f"Chat {query.message.chat.id}: Project selected - {query.data}")

    keyboard = [[InlineKeyboardButton(o, callback_data=o)] for o in objects]
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text("Выберите объект:", reply_markup=reply_markup)
    return OBJECT

async def object_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор объекта и запрашивает наименование позиции."""
    query = update.callback_query
    await query.answer()

    user_state[query.message.chat.id]["object"] = query.data
    logger.info(f"Chat {query.message.chat.id}: Object selected - {query.data}")
    await query.edit_message_text("Введите наименование позиции:")
    return NAME

async def name_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает наименование позиции и предлагает выбрать единицу измерения."""
    user_state[update.effective_chat.id]["current"] = {"name": update.message.text, "file_data": []}
    logger.info(f"Chat {update.effective_chat.id}: Position name entered - {update.message.text}")

    keyboard = [[InlineKeyboardButton(u, callback_data=u)] for u in units]
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Выберите единицу измерения:", reply_markup=reply_markup)
    return UNIT

async def unit_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор единицы измерения и запрашивает количество."""
    query = update.callback_query
    await query.answer()

    user_state[query.message.chat.id]["current"]["unit"] = query.data
    logger.info(f"Chat {query.message.chat.id}: Unit selected - {query.data}")
    await query.edit_message_text("Введите количество:")
    return QUANTITY

async def quantity_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает количество и предлагает выбрать модуль. Добавлена базовая валидация."""
    chat_id = update.effective_chat.id
    try:
        quantity = float(update.message.text)
        user_state[chat_id]["current"]["quantity"] = quantity
        logger.info(f"Chat {chat_id}: Quantity entered - {quantity}")
    except ValueError:
        logger.warning(f"Chat {chat_id}: Invalid quantity format - '{update.message.text}'")
        await update.message.reply_text("Неверный формат количества. Пожалуйста, введите число (например, 5 или 3.5):")
        return QUANTITY

    buttons_per_row = 5
    keyboard = []
    current_row = []
    for i, m in enumerate(modules):
        current_row.append(InlineKeyboardButton(m, callback_data=m))
        if (i + 1) % buttons_per_row == 0 or (i + 1) == len(modules):
            keyboard.append(current_row)
            current_row = []
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("К какому модулю относится позиция?", reply_markup=reply_markup)
    return MODULE

async def module_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает выбор модуля и переходит к выбору даты поставки для текущей позиции.
    """
    query = update.callback_query
    await query.answer()

    chat_id = query.message.chat.id
    user_state[chat_id]["current"]["module"] = query.data
    logger.info(f"Chat {chat_id}: Module selected - {query.data}. Requesting delivery date for this position.")

    current_date = date.today()
    reply_markup = create_calendar_keyboard(current_date.year, current_date.month, prefix="POS_CAL_")
    await query.edit_message_text("Выберите желаемую дату поставки для этой позиции:", reply_markup=reply_markup)
    return POSITION_DELIVERY_DATE

async def process_position_calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает нажатия на кнопки календаря для выбора даты поставки отдельной позиции.
    """
    query = update.callback_query
    await query.answer()

    data = query.data
    chat_id = query.message.chat.id

    if data.startswith("POS_CAL_NAV_"):
        parts = data.split('_')
        year = int(parts[3])
        month = int(parts[4])

        if month > 12:
            month = 1
            year += 1
        elif month < 1:
            month = 12
            year -= 1

        reply_markup = create_calendar_keyboard(year, month, prefix="POS_CAL_")
        await query.edit_message_text("Выберите желаемую дату поставки для этой позиции:", reply_markup=reply_markup)
        return POSITION_DELIVERY_DATE

    elif data.startswith("POS_CAL_DATE_"):
        selected_date_str = data.replace("POS_CAL_DATE_", "")

        user_state[chat_id]["current"]["delivery_date"] = selected_date_str
        logger.info(f"Chat {chat_id}: Position delivery date selected - {selected_date_str}. Now asking about attachments.")

        keyboard = [
            [InlineKeyboardButton("Прикрепить файл", callback_data="attach_file")],
            [InlineKeyboardButton("Прикрепить ссылку", callback_data="attach_link")],
            [InlineKeyboardButton("Продолжить", callback_data="no_attachment")]
        ]
        keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text("Теперь вы можете прикрепить файл или ссылку к этой позиции:", reply_markup=reply_markup)
        return ATTACHMENT_CHOICE

    elif data == "POS_CAL_CANCEL":
        logger.info(f"Chat {chat_id}: Position calendar date selection cancelled.")
        if "current" in user_state[chat_id]:
            del user_state[chat_id]["current"]
        await query.edit_message_text("Выбор даты для позиции отменен. Вы можете добавить позицию снова или продолжить.")
        return await edit_menu_handler(update, context)

    return POSITION_DELIVERY_DATE

async def attachment_choice_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает выбор пользователя по прикреплению файла, ссылки или продолжению без вложений.
    """
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    data = query.data

    if data == "attach_file":
        await query.edit_message_text("Пожалуйста, **отправьте мне файл** (как документ) для этой позиции.",
                                      reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")]]))
        return FILE_INPUT
    elif data == "attach_link":
        await query.edit_message_text("Пожалуйста, **введите ссылку** для этой позиции.",
                                      reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")]]))
        return LINK_INPUT
    elif data == "no_attachment":
        user_state[chat_id]["positions"].append(user_state[chat_id]["current"])
        logger.info(f"Chat {chat_id}: Position added without attachments: {user_state[chat_id]['current']}")
        del user_state[chat_id]["current"]

        keyboard = [
            [InlineKeyboardButton("Да", callback_data="yes"), InlineKeyboardButton("Нет", callback_data="no")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text("Позиция добавлена. Добавить ещё позицию?", reply_markup=reply_markup)
        return CONFIRM_ADD_MORE
    else:
        await query.edit_message_text("Неизвестный выбор.")
        return ATTACHMENT_CHOICE

async def handle_file_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает получение файла и сохраняет его данные в текущую позицию.
    """
    chat_id = update.effective_chat.id
    if update.message.document:
        document = update.message.document
        if "file_data" not in user_state[chat_id]["current"]:
            user_state[chat_id]["current"]["file_data"] = []
        user_state[chat_id]["current"]["file_data"].append({
            'file_id': document.file_id,
            'file_name': document.file_name,
            'mime_type': document.mime_type
        })
        logger.info(f"Chat {chat_id}: File '{document.file_name}' attached to current position.")
        await update.message.reply_text(f"Файл '{document.file_name}' успешно прикреплен.")
    else:
        logger.warning(f"Chat {chat_id}: Expected document but received something else for file input.")
        await update.message.reply_text("Это не похоже на файл-документ. Пожалуйста, отправьте файл (документ).")
        return FILE_INPUT

    keyboard = [
        [InlineKeyboardButton("Прикрепить файл", callback_data="attach_file")],
        [InlineKeyboardButton("Прикрепить ссылку", callback_data="attach_link")],
        [InlineKeyboardButton("Продолжить", callback_data="no_attachment")]
    ]
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Что дальше?", reply_markup=reply_markup)
    return ATTACHMENT_CHOICE

async def handle_link_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает получение ссылки и сохраняет ее в текущую позицию.
    """
    chat_id = update.effective_chat.id
    link = update.message.text.strip()
    if link.startswith("http://") or link.startswith("https://"):
        user_state[chat_id]["current"]["link"] = link
        logger.info(f"Chat {chat_id}: Link '{link}' attached to current position.")
        await update.message.reply_text(f"Ссылка '{link}' успешно прикреплена.")
    else:
        logger.warning(f"Chat {chat_id}: Invalid link format for link input - '{link}'")
        await update.message.reply_text("Пожалуйста, введите корректную ссылку, начинающуюся с http:// или https://")
        return LINK_INPUT

    keyboard = [
        [InlineKeyboardButton("Прикрепить файл", callback_data="attach_file")],
        [InlineKeyboardButton("Прикрепить ссылку", callback_data="attach_link")],
        [InlineKeyboardButton("Продолжить", callback_data="no_attachment")]
    ]
    keyboard.append([InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Что дальше?", reply_markup=reply_markup)
    return ATTACHMENT_CHOICE

async def confirm_add_more_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает подтверждение добавления еще одной позиции.
    """
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    if query.data == "yes":
        logger.info(f"Chat {chat_id}: User wants to add more positions.")
        await query.edit_message_text("Введите наименование следующей позиции:")
        return NAME
    else:
        logger.info(f"Chat {chat_id}: User finished adding positions. Proceeding to final confirmation.")
        return await final_confirm_handler(update, context)


async def final_confirm_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Представляет окончательное подтверждение заявки и предлагает редактирование или отправку.
    """
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    data = user_state[chat_id]
    project = data["project"]
    object_name = data["object"]
    positions = data["positions"]

    summary = f"**Ваша заявка:**\nПроект: {project}\nОбъект: {object_name}\n\n**Позиции:**\n"
    for i, pos in enumerate(positions):
        summary += (
            f"{i+1}. {pos['name']} ({pos['quantity']} {pos['unit']}) "
            f"Модуль: {pos['module']} "
            f"Дата поставки: {pos.get('delivery_date', 'Не указано')}\n"
        )
        if pos.get('link'):
            summary += f"   Ссылка: {pos['link']}\n"
        if pos.get('file_data'):
            file_names = ", ".join([f['file_name'] for f in pos['file_data']])
            summary += f"   Файлы: {file_names}\n"
    summary += "\n**Все верно?**"

    # === НАЧАЛО ИЗМЕНЕНИЙ ===
    # Вместо прямой отправки, предлагаем выбрать специалиста
    keyboard = [
        [InlineKeyboardButton("Алеся Забавская", callback_data="specialist_alesya")],
        [InlineKeyboardButton("Дмитрий Карп", callback_data="specialist_dmitry")],
        [InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text(summary + "\n\nКому направить заявку?", reply_markup=reply_markup, parse_mode='Markdown')
    return SPECIALIST_SELECTION # Переход в новое состояние
    # === КОНЕЦ ИЗМЕНЕНИЙ ===


# === НАЧАЛО ИЗМЕНЕНИЙ ===
async def specialist_selection_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает выбор специалиста и отправляет заявку.
    """
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    data = user_state[chat_id]
    project = data["project"]
    object_name = data["object"]
    positions = data["positions"]
    user_full_name = data["user_full_name"]
    telegram_id_or_username = data["telegram_id_or_username"]

    to_email = ""
    specialist_name = ""

    if query.data == "specialist_alesya":
        to_email = SPECIALIST_EMAIL_ALESYA
        specialist_name = "Алесе Забавской"
    elif query.data == "specialist_dmitry":
        to_email = SPECIALIST_EMAIL_DMITRY
        specialist_name = "Дмитрию Карпу"
    else:
        await query.edit_message_text("Неизвестный выбор специалиста. Пожалуйста, попробуйте снова.", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")]]))
        return SPECIALIST_SELECTION # Остаемся в этом состоянии, если выбор некорректен

    try:
        # Вызываем send_email с выбранной почтой специалиста и общей копией
        success = await send_email(
            chat_id, project, object_name, positions, user_full_name,
            telegram_id_or_username, to_email, CC_EMAIL, context
        )
        if success:
            await query.edit_message_text(f"Заявка успешно отправлена {specialist_name} (копия {CC_EMAIL}). Спасибо!",
                                          reply_markup=ReplyKeyboardRemove())
        else:
            await query.edit_message_text("Произошла ошибка при отправке заявки. Пожалуйста, попробуйте еще раз.",
                                          reply_markup=ReplyKeyboardRemove())
    except Exception as e:
        logger.error(f"Ошибка при отправке заявки после выбора специалиста: {e}")
        await query.edit_message_text(f"Произошла критическая ошибка при отправке заявки: {e}. Пожалуйста, сообщите об этом администратору.",
                                      reply_markup=ReplyKeyboardRemove())
    
    del user_state[chat_id] # Очищаем состояние пользователя после завершения
    return ConversationHandler.END # Завершаем диалог

# === КОНЕЦ ИЗМЕНЕНИЙ ===


async def edit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает меню редактирования заявки."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    keyboard = [
        [InlineKeyboardButton("Редактировать позицию", callback_data="edit_position")],
        [InlineKeyboardButton("Удалить позицию", callback_data="delete_position")],
        [InlineKeyboardButton("Изменить проект", callback_data="edit_project")],
        [InlineKeyboardButton("Изменить объект", callback_data="edit_object")],
        [InlineKeyboardButton("Изменить глобальную дату поставки", callback_data="edit_global_delivery_date")],
        [InlineKeyboardButton("Все верно, продолжить", callback_data="confirm_final")],
        [InlineKeyboardButton("Отмена заявки", callback_data="cancel_dialog")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text("Что вы хотите сделать?", reply_markup=reply_markup)
    return EDIT_MENU

# Функции календаря - без изменений, но добавлены для полноты кода
def create_calendar_keyboard(year, month, prefix="CAL_"):
    """
    Создает инлайн-клавиатуру с календарем для выбора даты.
    """
    keyboard = []
    # Заголовки дней недели
    row = []
    for day_name in ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]:
        row.append(InlineKeyboardButton(day_name, callback_data="ignore"))
    keyboard.append(row)

    my_calendar = calendar.monthcalendar(year, month)
    for week in my_calendar:
        row = []
        for day in week:
            if day == 0:
                row.append(InlineKeyboardButton(" ", callback_data="ignore"))
            else:
                date_str = f"{year}-{month:02d}-{day:02d}"
                row.append(InlineKeyboardButton(str(day), callback_data=f"{prefix}DATE_{date_str}"))
        keyboard.append(row)

    # Кнопки навигации
    prev_month = month - 1
    prev_year = year
    if prev_month < 1:
        prev_month = 12
        prev_year -= 1

    next_month = month + 1
    next_year = year
    if next_month > 12:
        next_month = 1
        next_year += 1

    keyboard.append([
        InlineKeyboardButton(f"< {calendar.month_name[prev_month][:3]}", callback_data=f"{prefix}NAV_{prev_year}_{prev_month}"),
        InlineKeyboardButton(f"{calendar.month_name[month]} {year}", callback_data="ignore"),
        InlineKeyboardButton(f"{calendar.month_name[next_month][:3]} >", callback_data=f"{prefix}NAV_{next_year}_{next_month}")
    ])
    keyboard.append([InlineKeyboardButton("Отмена", callback_data=f"{prefix}CANCEL")]) # Кнопка отмены

    return InlineKeyboardMarkup(keyboard)

async def process_global_calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Обрабатывает нажатия на кнопки календаря для глобального выбора даты поставки.
    """
    query = update.callback_query
    await query.answer()

    data = query.data
    chat_id = query.message.chat.id

    if data.startswith("CAL_NAV_") or data.startswith("EDIT_CAL_NAV_"):
        parts = data.split('_')
        prefix = f"{parts[0]}_{parts[1]}_" # CAL_NAV_ or EDIT_CAL_NAV_ -> CAL_ or EDIT_CAL_
        year = int(parts[3])
        month = int(parts[4])

        if month > 12:
            month = 1
            year += 1
        elif month < 1:
            month = 12
            year -= 1

        reply_markup = create_calendar_keyboard(year, month, prefix=prefix)
        await query.edit_message_text("Выберите глобальную дату поставки:", reply_markup=reply_markup)
        return GLOBAL_DELIVERY_DATE_SELECTION

    elif data.startswith("CAL_DATE_") or data.startswith("EDIT_CAL_DATE_"):
        selected_date_str = data.replace("CAL_DATE_", "").replace("EDIT_CAL_DATE_", "")
        
        # Обновляем все позиции с глобальной датой поставки
        for pos in user_state[chat_id]["positions"]:
            pos["delivery_date"] = selected_date_str
        logger.info(f"Chat {chat_id}: Global delivery date set to {selected_date_str} for all positions.")
        
        await query.edit_message_text(f"Глобальная дата поставки установлена на: {selected_date_str}. Все позиции обновлены.",
                                      reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Продолжить", callback_data="confirm_final")]]))
        return FINAL_CONFIRMATION # Возвращаемся в FINAL_CONFIRMATION
    elif data == "CAL_CANCEL" or data == "EDIT_CAL_CANCEL":
        await query.edit_message_text("Выбор глобальной даты отменен. Продолжить без глобальной даты.",
                                      reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Продолжить", callback_data="confirm_final")]]))
        return FINAL_CONFIRMATION # Возвращаемся в FINAL_CONFIRMATION
    
    return GLOBAL_DELIVERY_DATE_SELECTION

async def edit_field_input_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает введенные данные для редактируемого поля."""
    chat_id = update.effective_chat.id
    current_edit = user_state[chat_id].get("editing_position_index_and_field")
    
    if not current_edit:
        await update.message.reply_text("Ошибка: нет активного редактирования. Пожалуйста, начните сначала.")
        return ConversationHandler.END

    pos_index, field_to_edit = current_edit
    new_value = update.message.text

    if field_to_edit == "name":
        user_state[chat_id]["positions"][pos_index]["name"] = new_value
        await update.message.reply_text(f"Наименование позиции {pos_index+1} изменено на '{new_value}'.")
    elif field_to_edit == "quantity":
        try:
            quantity = float(new_value)
            user_state[chat_id]["positions"][pos_index]["quantity"] = quantity
            await update.message.reply_text(f"Количество позиции {pos_index+1} изменено на '{quantity}'.")
        except ValueError:
            await update.message.reply_text("Неверный формат количества. Пожалуйста, введите число.")
            return EDIT_FIELD_INPUT # Остаемся в этом состоянии
    elif field_to_edit == "link":
        if new_value.startswith("http://") or new_value.startswith("https://"):
            user_state[chat_id]["positions"][pos_index]["link"] = new_value
            await update.message.reply_text(f"Ссылка позиции {pos_index+1} изменена на '{new_value}'.")
        else:
            await update.message.reply_text("Неверный формат ссылки. Пожалуйста, введите ссылку, начинающуюся с http:// или https://")
            return EDIT_FIELD_INPUT # Остаемся в этом состоянии
    else:
        await update.message.reply_text("Неизвестное поле для редактирования.")
        
    del user_state[chat_id]["editing_position_index_and_field"] # Очищаем состояние редактирования
    return await edit_menu_handler(update, context) # Возвращаемся в меню редактирования


async def select_position_to_edit_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает список позиций для выбора редактирования."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    if query.data == "edit_position" or query.data == "delete_position":
        context.user_data["action_type"] = query.data # Сохраняем тип действия
    
    positions = user_state[chat_id]["positions"]
    if not positions:
        await query.edit_message_text("Нет позиций для редактирования/удаления.", 
                                      reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Назад в меню", callback_data="back_to_edit_menu")]]))
        return EDIT_MENU # Или другое подходящее состояние

    keyboard = []
    for i, pos in enumerate(positions):
        keyboard.append([InlineKeyboardButton(f"{i+1}. {pos['name']}", callback_data=f"select_pos_{i}")])
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")]) # Или "Назад"
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text("Выберите позицию:", reply_markup=reply_markup)
    return SELECT_POSITION

async def edit_field_selection_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Позволяет выбрать поле для редактирования выбранной позиции."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    if query.data.startswith("select_pos_"):
        pos_index = int(query.data.replace("select_pos_", ""))
        user_state[chat_id]["editing_position_index"] = pos_index # Сохраняем индекс редактируемой позиции

        action_type = context.user_data.get("action_type")

        if action_type == "delete_position":
            del user_state[chat_id]["positions"][pos_index]
            await query.edit_message_text(f"Позиция {pos_index+1} удалена.")
            if not user_state[chat_id]["positions"]:
                await query.message.reply_text("Все позиции удалены. Начните новую заявку, или отмените текущую.")
                return ConversationHandler.END # Если позиций не осталось, завершаем
            return await edit_menu_handler(update, context) # Возвращаемся в меню редактирования

        keyboard = [
            [InlineKeyboardButton("Наименование", callback_data="edit_name")],
            [InlineKeyboardButton("Единица измерения", callback_data="edit_unit")],
            [InlineKeyboardButton("Количество", callback_data="edit_quantity")],
            [InlineKeyboardButton("Модуль", callback_data="edit_module")],
            [InlineKeyboardButton("Дата поставки", callback_data="edit_delivery_date")],
            [InlineKeyboardButton("Ссылка", callback_data="edit_link")],
            [InlineKeyboardButton("Отмена", callback_data="cancel_dialog")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(f"Что вы хотите отредактировать в позиции {pos_index+1}?", reply_markup=reply_markup)
        return EDIT_FIELD_SELECTION
    else:
        await query.edit_message_text("Неизвестный выбор позиции. Пожалуйста, попробуйте еще раз.")
        return SELECT_POSITION

async def process_field_to_edit_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Запрашивает новое значение для выбранного поля."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    pos_index = user_state[chat_id]["editing_position_index"]
    field_to_edit = query.data.replace("edit_", "")

    user_state[chat_id]["editing_position_index_and_field"] = (pos_index, field_to_edit) # Сохраняем поле для редактирования

    if field_to_edit == "unit":
        keyboard = [[InlineKeyboardButton(u, callback_data=f"edit_unit_{u}")] for u in units]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(f"Выберите новую единицу измерения для позиции {pos_index+1}:", reply_markup=reply_markup)
        return EDITING_UNIT
    elif field_to_edit == "module":
        buttons_per_row = 5
        keyboard = []
        current_row = []
        for i, m in enumerate(modules):
            current_row.append(InlineKeyboardButton(m, callback_data=f"edit_module_{m}"))
            if (i + 1) % buttons_per_row == 0 or (i + 1) == len(modules):
                keyboard.append(current_row)
                current_row = []
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(f"Выберите новый модуль для позиции {pos_index+1}:", reply_markup=reply_markup)
        return EDITING_MODULE
    elif field_to_edit == "delivery_date":
        current_date = date.today()
        reply_markup = create_calendar_keyboard(current_date.year, current_date.month, prefix="POS_CAL_") # Используем тот же префикс
        await query.edit_message_text(f"Выберите новую дату поставки для позиции {pos_index+1}:", reply_markup=reply_markup)
        return POSITION_DELIVERY_DATE # Возвращаемся к обработчику календаря позиций
    else:
        await query.edit_message_text(f"Введите новое значение для поля '{field_to_edit}' позиции {pos_index+1}:")
        return EDIT_FIELD_INPUT

async def process_edited_unit_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор новой единицы измерения при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    pos_index, _ = user_state[chat_id]["editing_position_index_and_field"]
    new_unit = query.data.replace("edit_unit_", "")
    user_state[chat_id]["positions"][pos_index]["unit"] = new_unit
    await query.edit_message_text(f"Единица измерения позиции {pos_index+1} изменена на '{new_unit}'.")
    del user_state[chat_id]["editing_position_index_and_field"]
    return await edit_menu_handler(update, context)

async def process_edited_module_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор нового модуля при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    pos_index, _ = user_state[chat_id]["editing_position_index_and_field"]
    new_module = query.data.replace("edit_module_", "")
    user_state[chat_id]["positions"][pos_index]["module"] = new_module
    await query.edit_message_text(f"Модуль позиции {pos_index+1} изменен на '{new_module}'.")
    del user_state[chat_id]["editing_position_index_and_field"]
    return await edit_menu_handler(update, context)

async def handle_edit_project(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает запрос на изменение проекта."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    keyboard = [[InlineKeyboardButton(p, callback_data=f"edit_project_{p}")] for p in projects]
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text("Выберите новый проект:", reply_markup=reply_markup)
    return PROJECT # Возвращаемся в состояние PROJECT, но с новым callback_data

async def handle_edit_object(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает запрос на изменение объекта."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    keyboard = [[InlineKeyboardButton(o, callback_data=f"edit_object_{o}")] for o in objects]
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.edit_message_text("Выберите новый объект:", reply_markup=reply_markup)
    return OBJECT # Возвращаемся в состояние OBJECT, но с новым callback_data

async def process_edited_project_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор нового проекта при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    new_project = query.data.replace("edit_project_", "")
    user_state[chat_id]["project"] = new_project
    await query.edit_message_text(f"Проект изменен на: {new_project}.")
    return await edit_menu_handler(update, context)

async def process_edited_object_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает выбор нового объекта при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    new_object = query.data.replace("edit_object_", "")
    user_state[chat_id]["object"] = new_object
    await query.edit_message_text(f"Объект изменен на: {new_object}.")
    return await edit_menu_handler(update, context)

async def handle_edit_global_delivery_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает запрос на изменение глобальной даты поставки."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    
    current_date = date.today()
    reply_markup = create_calendar_keyboard(current_date.year, current_date.month, prefix="EDIT_CAL_")
    await query.edit_message_text("Выберите новую глобальную дату поставки для всех позиций:", reply_markup=reply_markup)
    return GLOBAL_DELIVERY_DATE_SELECTION # Переход в состояние календаря

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Отменяет и завершает разговор."""
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(
            "Создание заявки отменено.",
            reply_markup=ReplyKeyboardRemove()
        )
        chat_id = update.callback_query.message.chat.id
    else:
        await update.message.reply_text(
            "Создание заявки отменено.",
            reply_markup=ReplyKeyboardRemove()
        )
        chat_id = update.effective_chat.id
    
    if chat_id in user_state:
        del user_state[chat_id]
        logger.info(f"Chat {chat_id}: User state cleared due to cancellation.")
    
    return ConversationHandler.END

async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает неизвестные команды/сообщения."""
    if update.message:
        await update.message.reply_text("Извините, я не понял эту команду или сообщение. Пожалуйста, используйте кнопки.")

# Main function to run the bot
async def main() -> None:
    application = ApplicationBuilder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, start_conversation)], # Start with "Создать заявку"
        states={
            PROJECT: [CallbackQueryHandler(project_handler)],
            OBJECT: [CallbackQueryHandler(object_handler)],
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, name_handler)],
            UNIT: [CallbackQueryHandler(unit_handler)],
            QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, quantity_handler)],
            MODULE: [CallbackQueryHandler(module_handler)],
            POSITION_DELIVERY_DATE: [CallbackQueryHandler(process_position_calendar_callback)],
            ATTACHMENT_CHOICE: [CallbackQueryHandler(attachment_choice_handler)],
            FILE_INPUT: [MessageHandler(filters.Document.ALL, handle_file_input)],
            LINK_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_link_input)],
            CONFIRM_ADD_MORE: [CallbackQueryHandler(confirm_add_more_handler)],
            EDIT_MENU: [
                CallbackQueryHandler(select_position_to_edit_handler, pattern="^(edit_position|delete_position)$"),
                CallbackQueryHandler(handle_edit_project, pattern="^edit_project$"),
                CallbackQueryHandler(handle_edit_object, pattern="^edit_object$"),
                CallbackQueryHandler(handle_edit_global_delivery_date, pattern="^edit_global_delivery_date$"),
                CallbackQueryHandler(final_confirm_handler, pattern="^confirm_final$"), # Изменили, чтобы вел на выбор специалиста
            ],
            SELECT_POSITION: [CallbackQueryHandler(edit_field_selection_handler, pattern="^select_pos_")],
            EDIT_FIELD_SELECTION: [
                CallbackQueryHandler(process_field_to_edit_selection, pattern="^edit_"),
            ],
            EDIT_FIELD_INPUT: [
                MessageHandler(filters.TEXT | filters.Document.ALL & ~filters.COMMAND, edit_field_input_handler),
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],

            EDITING_UNIT: [
                CallbackQueryHandler(process_edited_unit_selection, pattern="^edit_unit_"),
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],
            EDITING_MODULE: [
                CallbackQueryHandler(process_edited_module_selection, pattern="^edit_module_"),
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],

            GLOBAL_DELIVERY_DATE_SELECTION: [
                CallbackQueryHandler(process_global_calendar_callback, pattern="^(CAL_|EDIT_CAL_)\\d+_\\d+"), # Уточнили паттерн
                CallbackQueryHandler(process_global_calendar_callback, pattern="^(CAL_|EDIT_CAL_)\d+_?\d*_\d*"), # Более общий паттерн для навигации
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],

            FINAL_CONFIRMATION: [ # Это состояние теперь переходит к выбору специалиста
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$"),
                CallbackQueryHandler(final_confirm_handler) # Этот handler теперь ведет на SPECIALIST_SELECTION
            ],
            # === НАЧАЛО ИЗМЕНЕНИЙ ===
            SPECIALIST_SELECTION: [
                CallbackQueryHandler(specialist_selection_handler, pattern="^specialist_"),
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],
            # === КОНЕЦ ИЗМЕНЕНИЙ ===
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CallbackQueryHandler(cancel, pattern="^cancel_dialog$"),
            MessageHandler(filters.COMMAND | filters.TEXT, unknown)
        ],
    )

    application.add_handler(conv_handler)

    application.add_handler(CommandHandler("start", initial_message_handler))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, initial_message_handler))


    await application.run_polling()

if __name__ == "__main__":
    try:
        # Проверяем, запущен ли уже цикл событий
        loop = asyncio.get_running_loop()
        print("Event loop is already running, using existing loop.")
        loop.create_task(main())
    except RuntimeError:
        # Если цикл не запущен, запускаем новый
        print("No event loop is running, starting a new one.")
        asyncio.run(main())