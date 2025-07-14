import os
import logging
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
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")
TEMPLATE_PATH = "template.xlsx" # Убедитесь, что template.xlsx существует в той же директории

# Состояния для ConversationHandler
# Обновлено количество состояний до 19
PROJECT, OBJECT, NAME, UNIT, QUANTITY, MODULE, POSITION_DELIVERY_DATE, \
ATTACHMENT_CHOICE, FILE_INPUT, LINK_INPUT, \
CONFIRM_ADD_MORE, \
EDIT_MENU, SELECT_POSITION, EDIT_FIELD_SELECTION, EDIT_FIELD_INPUT, \
FINAL_CONFIRMATION, GLOBAL_DELIVERY_DATE_INPUT, EDITING_UNIT, EDITING_MODULE = range(19)

# Словарь для хранения данных заявки
user_data_store = {}
temp_data_store = {}

# Временное хранилище для файла при редактировании
temp_file_storage = {}

# Список для хранения номеров модулей
modules = [f"{i+1}" for i in range(18)]

# Списки для кнопок
project_buttons = [
    [InlineKeyboardButton("Квартира", callback_data="Квартира")],
    [InlineKeyboardButton("Офис", callback_data="Офис")],
    [InlineKeyboardButton("Другое", callback_data="Другое")]
]
project_keyboard = InlineKeyboardMarkup(project_buttons)

confirm_add_more_keyboard = InlineKeyboardMarkup([
    [InlineKeyboardButton("Добавить еще позицию", callback_data="add_more_position")],
    [InlineKeyboardButton("Завершить и отправить", callback_data="finish_and_send")],
    [InlineKeyboardButton("Отменить заявку", callback_data="cancel_dialog")]
])

attachment_choice_keyboard = InlineKeyboardMarkup([
    [InlineKeyboardButton("Прикрепить файл", callback_data="attach_file")],
    [InlineKeyboardButton("Прикрепить ссылку", callback_data="attach_link")],
    [InlineKeyboardButton("Пропустить", callback_data="skip_attachment")]
])

edit_menu_keyboard = InlineKeyboardMarkup([
    [InlineKeyboardButton("Редактировать позицию", callback_data="edit_position")],
    [InlineKeyboardButton("Добавить позицию", callback_data="add_position_from_edit")],
    [InlineKeyboardButton("Удалить позицию", callback_data="delete_position")],
    [InlineKeyboardButton("Задать общую дату доставки", callback_data="set_global_delivery_date")],
    [InlineKeyboardButton("Продолжить и отправить", callback_data="continue_and_send")],
    [InlineKeyboardButton("Отменить заявку", callback_data="cancel_dialog")]
])

final_confirm_keyboard = InlineKeyboardMarkup([
    [InlineKeyboardButton("Подтвердить и отправить", callback_data="confirm_final_send")],
    [InlineKeyboardButton("Отменить", callback_data="cancel_dialog")]
])


async def start_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отправляет начальное сообщение и предлагает выбрать проект."""
    chat_id = update.effective_chat.id
    user_data_store[chat_id] = {'positions': []}
    await update.message.reply_text(
        "Привет! Я бот для создания заявок на снабжение.\n\n"
        "Для начала, выберите тип проекта:", reply_markup=project_keyboard
    )
    return PROJECT

async def initial_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает команды /start и произвольный текст для инициализации диалога."""
    chat_id = update.effective_chat.id
    if chat_id not in user_data_store or not user_data_store[chat_id].get('positions'):
        user_data_store[chat_id] = {'positions': []}
        await update.message.reply_text(
            "Привет! Я бот для создания заявок на снабжение.\n\n"
            "Для начала, выберите тип проекта:", reply_markup=project_keyboard
        )
        return PROJECT
    else:
        await update.message.reply_text("Вы уже начали заявку. Вы можете продолжить или /cancel.")
        return ConversationHandler.WAITING


async def project_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор проекта."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    project = query.data
    user_data_store[chat_id]['project'] = project
    await query.edit_message_text(f"Вы выбрали проект: {project}\nТеперь введите название объекта:")
    return OBJECT

async def object_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод названия объекта."""
    chat_id = update.effective_chat.id
    object_name = update.message.text
    user_data_store[chat_id]['object_name'] = object_name
    await update.message.reply_text("Отлично! Теперь введите наименование позиции:")
    return NAME

async def name_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод наименования позиции."""
    chat_id = update.effective_chat.id
    position_name = update.message.text
    temp_data_store[chat_id] = {'name': position_name} # Временно храним данные для текущей позиции

    unit_keyboard = ReplyKeyboardMarkup([
        ["м", "м.кв.", "м.куб."],
        ["шт.", "компл.", "усл.ед."],
        ["Отмена"]
    ], one_time_keyboard=True, resize_keyboard=True)

    await update.message.reply_text("Введите единицу измерения:", reply_markup=unit_keyboard)
    return UNIT

async def unit_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод единицы измерения."""
    chat_id = update.effective_chat.id
    unit = update.message.text

    if unit == "Отмена":
        await update.message.reply_text("Действие отменено. Введите наименование позиции заново:", reply_markup=ReplyKeyboardRemove())
        return NAME

    temp_data_store[chat_id]['unit'] = unit
    await update.message.reply_text("Введите количество:", reply_markup=ReplyKeyboardRemove())
    return QUANTITY

async def quantity_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод количества."""
    chat_id = update.effective_chat.id
    quantity_text = update.message.text
    try:
        quantity = float(quantity_text.replace(',', '.')) # Заменяем запятую на точку для правильного преобразования
        temp_data_store[chat_id]['quantity'] = quantity
    except ValueError:
        await update.message.reply_text("Неверное количество. Пожалуйста, введите числовое значение:")
        return QUANTITY

    module_buttons = [[InlineKeyboardButton(m, callback_data=m)] for m in modules]
    module_buttons.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
    module_keyboard = InlineKeyboardMarkup(module_buttons)

    await update.message.reply_text("Выберите номер модуля (1-18):", reply_markup=module_keyboard)
    return MODULE

async def module_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор номера модуля."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    module = query.data
    temp_data_store[chat_id]['module'] = module

    calendar_markup = create_calendar()
    await query.edit_message_text("Выберите дату доставки для этой позиции:", reply_markup=calendar_markup)
    return POSITION_DELIVERY_DATE

async def process_calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает коллбэки календаря для выбора даты позиции."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    data = query.data

    if data.startswith('CAL_PREV_MONTH'):
        parts = data.split('_')
        month = int(parts[3])
        year = int(parts[4])
        new_date = datetime(year, month, 1) - timedelta(days=1)
        calendar_markup = create_calendar(new_date.year, new_date.month)
        await query.edit_message_reply_markup(reply_markup=calendar_markup)
        return POSITION_DELIVERY_DATE
    elif data.startswith('CAL_NEXT_MONTH'):
        parts = data.split('_')
        month = int(parts[3])
        year = int(parts[4])
        new_date = datetime(year, month, 1) + timedelta(days=31)
        calendar_markup = create_calendar(new_date.year, new_date.month)
        await query.edit_message_reply_markup(reply_markup=calendar_markup)
        return POSITION_DELIVERY_DATE
    elif data.startswith('CAL_DATE'):
        date_str = data.split('_')[2]
        selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        temp_data_store[chat_id]['delivery_date'] = selected_date.strftime('%d.%m.%Y')
        await query.edit_message_text(f"Дата доставки для позиции: {selected_date.strftime('%d.%m.%Y')}\n\nХотите прикрепить файл или ссылку к этой позиции?", reply_markup=attachment_choice_keyboard)
        return ATTACHMENT_CHOICE
    else:
        # Это должен быть CAL_IGNORE, ничего не делаем
        return POSITION_DELIVERY_DATE


async def attachment_choice_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор прикрепления файла/ссылки или пропуска."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    choice = query.data

    if choice == "attach_file":
        temp_data_store[chat_id]['attachment_type'] = 'file'
        await query.edit_message_text("Отправьте файл:")
        return FILE_INPUT
    elif choice == "attach_link":
        temp_data_store[chat_id]['attachment_type'] = 'link'
        await query.edit_message_text("Отправьте ссылку:")
        return LINK_INPUT
    elif choice == "skip_attachment":
        temp_data_store[chat_id]['attachment_type'] = 'none'
        temp_data_store[chat_id]['attachment_content'] = ''
        await add_current_position_to_store(chat_id, context)
        return CONFIRM_ADD_MORE

async def file_input_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод файла."""
    chat_id = update.effective_chat.id
    if update.message.document:
        document = update.message.document
        file_id = document.file_id
        file_name = document.file_name
        temp_data_store[chat_id]['attachment_content'] = f"file_id:{file_id}||file_name:{file_name}"
        await add_current_position_to_store(chat_id, context)
        return CONFIRM_ADD_MORE
    else:
        await update.message.reply_text("Пожалуйста, отправьте файл.")
        return FILE_INPUT

async def link_input_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод ссылки."""
    chat_id = update.effective_chat.id
    link = update.message.text
    if link and (link.startswith('http://') or link.startswith('https://')):
        temp_data_store[chat_id]['attachment_content'] = link
        await add_current_position_to_store(chat_id, context)
        return CONFIRM_ADD_MORE
    else:
        await update.message.reply_text("Пожалуйста, введите корректную ссылку, начинающуюся с http:// или https://")
        return LINK_INPUT

async def add_current_position_to_store(chat_id: int, context: ContextTypes.DEFAULT_TYPE):
    """Добавляет текущую позицию из temp_data_store в user_data_store."""
    current_position = temp_data_store.get(chat_id, {}).copy() # Копируем, чтобы избежать мутаций
    if 'positions' not in user_data_store[chat_id]:
        user_data_store[chat_id]['positions'] = []
    user_data_store[chat_id]['positions'].append(current_position)
    await context.bot.send_message(
        chat_id=chat_id,
        text="Позиция добавлена. Что дальше?",
        reply_markup=confirm_add_more_keyboard
    )

async def confirm_add_more_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор добавить еще позицию или завершить."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    choice = query.data

    if choice == "add_more_position":
        await query.edit_message_text("Отлично! Введите наименование следующей позиции:")
        return NAME
    elif choice == "finish_and_send":
        await query.edit_message_text("Почти готово! Проверьте вашу заявку:", reply_markup=ReplyKeyboardRemove())
        await send_summary(chat_id, context)
        await query.message.reply_text("Все верно?", reply_markup=final_confirm_keyboard)
        return FINAL_CONFIRMATION
    # Отмена обрабатывается в cancel


async def send_summary(chat_id: int, context: ContextTypes.DEFAULT_TYPE):
    """Отправляет краткую сводку текущей заявки."""
    user_data = user_data_store.get(chat_id, {})
    summary = "Ваша текущая заявка:\n\n"
    summary += f"Проект: {user_data.get('project', 'Не указан')}\n"
    summary += f"Объект: {user_data.get('object_name', 'Не указан')}\n\n"
    
    global_delivery_date = user_data.get('global_delivery_date')
    if global_delivery_date:
        summary += f"Общая дата доставки: {global_delivery_date}\n\n"

    positions = user_data.get('positions', [])
    if positions:
        for i, pos in enumerate(positions):
            summary += f"Позиция {i + 1}:\n"
            summary += f"  Наименование: {pos.get('name', '')}\n"
            summary += f"  Ед.изм.: {pos.get('unit', '')}\n"
            summary += f"  Количество: {pos.get('quantity', '')}\n"
            summary += f"  Модуль: {pos.get('module', '')}\n"
            summary += f"  Дата доставки: {pos.get('delivery_date', 'Не указана')}\n"
            attachment = pos.get('attachment_content', '')
            if attachment:
                if pos.get('attachment_type') == 'file':
                    # Здесь отображаем имя файла, если оно есть, иначе просто указываем "Файл"
                    file_info = attachment.split('||')
                    file_name = next((info.split(':')[1] for info in file_info if info.startswith('file_name:')), 'Файл')
                    summary += f"  Вложение: {file_name}\n"
                elif pos.get('attachment_type') == 'link':
                    summary += f"  Вложение: {attachment}\n"
            summary += "\n"
    else:
        summary += "Позиции пока не добавлены.\n"

    if len(summary) > 4096: # Telegram message length limit
        for x in range(0, len(summary), 4096):
            await context.bot.send_message(chat_id=chat_id, text=summary[x:x+4096])
    else:
        await context.bot.send_message(chat_id=chat_id, text=summary)

    await context.bot.send_message(chat_id=chat_id, text="Что вы хотите сделать?", reply_markup=edit_menu_keyboard)
    return EDIT_MENU

async def edit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор действия в меню редактирования."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    choice = query.data

    if choice == "edit_position":
        positions = user_data_store[chat_id].get('positions', [])
        if not positions:
            await query.edit_message_text("Нет позиций для редактирования. Добавьте позиции сначала.", reply_markup=edit_menu_keyboard)
            return EDIT_MENU

        buttons = []
        for i, pos in enumerate(positions):
            buttons.append([InlineKeyboardButton(f"Позиция {i + 1}: {pos.get('name', 'Без названия')}", callback_data=f"select_pos_{i}")])
        buttons.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
        
        select_position_keyboard = InlineKeyboardMarkup(buttons)
        await query.edit_message_text("Выберите позицию для редактирования:", reply_markup=select_position_keyboard)
        return SELECT_POSITION

    elif choice == "add_position_from_edit":
        await query.edit_message_text("Отлично! Введите наименование новой позиции:")
        return NAME # Возвращаемся к началу процесса добавления позиции

    elif choice == "delete_position":
        positions = user_data_store[chat_id].get('positions', [])
        if not positions:
            await query.edit_message_text("Нет позиций для удаления.", reply_markup=edit_menu_keyboard)
            return EDIT_MENU
        
        buttons = []
        for i, pos in enumerate(positions):
            buttons.append([InlineKeyboardButton(f"Удалить Позицию {i + 1}: {pos.get('name', 'Без названия')}", callback_data=f"delete_pos_{i}")])
        buttons.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
        
        delete_position_keyboard = InlineKeyboardMarkup(buttons)
        await query.edit_message_text("Выберите позицию для удаления:", reply_markup=delete_position_keyboard)
        return SELECT_POSITION # Используем то же состояние, но с другим коллбэком

    elif choice == "set_global_delivery_date":
        calendar_markup = create_calendar()
        await query.edit_message_text("Выберите общую дату доставки:", reply_markup=calendar_markup)
        return GLOBAL_DELIVERY_DATE_SELECTION

    elif choice == "continue_and_send":
        await query.edit_message_text("Вы закончили редактирование. Отправить заявку?", reply_markup=final_confirm_keyboard)
        return FINAL_CONFIRMATION
    
    # Отмена обрабатывается в cancel
    return EDIT_MENU

async def select_position_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор позиции для редактирования или удаления."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    data = query.data

    if data.startswith("select_pos_"):
        index = int(data.split('_')[2])
        user_data_store[chat_id]['current_edit_position_index'] = index
        position = user_data_store[chat_id]['positions'][index]

        buttons = []
        buttons.append([InlineKeyboardButton(f"Наименование: {position.get('name', '')}", callback_data="edit_field_name")])
        buttons.append([InlineKeyboardButton(f"Ед.изм.: {position.get('unit', '')}", callback_data="edit_field_unit")])
        buttons.append([InlineKeyboardButton(f"Количество: {position.get('quantity', '')}", callback_data="edit_field_quantity")])
        buttons.append([InlineKeyboardButton(f"Модуль: {position.get('module', '')}", callback_data="edit_field_module")])
        buttons.append([InlineKeyboardButton(f"Дата доставки: {position.get('delivery_date', 'Не указана')}", callback_data="edit_field_delivery_date")])
        
        attachment_info = position.get('attachment_content', '')
        attachment_type = position.get('attachment_type', 'none')
        if attachment_type == 'file':
            file_info = attachment_info.split('||')
            file_name = next((info.split(':')[1] for info in file_info if info.startswith('file_name:')), 'Файл')
            buttons.append([InlineKeyboardButton(f"Вложение (файл): {file_name}", callback_data="edit_field_attachment")])
        elif attachment_type == 'link':
            buttons.append([InlineKeyboardButton(f"Вложение (ссылка): {attachment_info}", callback_data="edit_field_attachment")])
        else:
            buttons.append([InlineKeyboardButton("Добавить вложение", callback_data="edit_field_attachment")])

        buttons.append([InlineKeyboardButton("Назад в меню редактирования", callback_data="back_to_edit_menu")])

        edit_field_keyboard = InlineKeyboardMarkup(buttons)
        await query.edit_message_text("Выберите поле для редактирования:", reply_markup=edit_field_keyboard)
        return EDIT_FIELD_SELECTION

    elif data.startswith("delete_pos_"):
        index = int(data.split('_')[2])
        deleted_position = user_data_store[chat_id]['positions'].pop(index)
        await query.edit_message_text(f"Позиция '{deleted_position.get('name', 'Без названия')}' удалена.")
        await send_summary(chat_id, context) # Обновляем сводку после удаления
        return EDIT_MENU
    
    return SELECT_POSITION


async def edit_field_selection_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор поля для редактирования."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    field = query.data

    if field == "back_to_edit_menu":
        await send_summary(chat_id, context)
        return EDIT_MENU
    
    user_data_store[chat_id]['current_edit_field'] = field

    if field == "edit_field_unit":
        unit_keyboard = ReplyKeyboardMarkup([
            ["м", "м.кв.", "м.куб."],
            ["шт.", "компл.", "усл.ед."],
            ["Отмена"]
        ], one_time_keyboard=True, resize_keyboard=True)
        await query.edit_message_text("Введите новую единицу измерения:", reply_markup=unit_keyboard)
        return EDITING_UNIT
    
    elif field == "edit_field_module":
        module_buttons = [[InlineKeyboardButton(m, callback_data=f"edit_module_{m}")] for m in modules]
        module_buttons.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])
        module_keyboard = InlineKeyboardMarkup(module_buttons)
        await query.edit_message_text("Выберите новый номер модуля (1-18):", reply_markup=module_keyboard)
        return EDITING_MODULE

    elif field == "edit_field_delivery_date":
        calendar_markup = create_calendar()
        await query.edit_message_text("Выберите новую дату доставки:", reply_markup=calendar_markup)
        return EDIT_FIELD_INPUT # Используем то же состояние для календаря, что и для обычного ввода
    
    elif field == "edit_field_attachment":
        await query.edit_message_text("Выберите тип нового вложения:", reply_markup=attachment_choice_keyboard)
        user_data_store[chat_id]['attachment_edit_mode'] = True # Флаг для обработки вложений в режиме редактирования
        return EDIT_FIELD_INPUT # Временное состояние для обработки выбора прикрепления
        
    else: # Для name, quantity
        await query.edit_message_text(f"Введите новое значение для {field.replace('edit_field_', '')}:")
        return EDIT_FIELD_INPUT


async def process_edited_unit_selection(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор единицы измерения при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    unit = query.data.replace('edit_unit_', '')
    
    index = user_data_store[chat_id]['current_edit_position_index']
    user_data_store[chat_id]['positions'][index]['unit'] = unit

    await query.edit_message_text(f"Единица измерения обновлена на: {unit}")
    await send_summary(chat_id, context) # Показываем обновленную сводку
    return EDIT_MENU

async def process_edited_module_selection(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает выбор модуля при редактировании."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    module = query.data.replace('edit_module_', '')
    
    index = user_data_store[chat_id]['current_edit_position_index']
    user_data_store[chat_id]['positions'][index]['module'] = module

    await query.edit_message_text(f"Модуль обновлен на: {module}")
    await send_summary(chat_id, context) # Показываем обновленную сводку
    return EDIT_MENU

async def edit_field_input_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает ввод нового значения для поля при редактировании."""
    chat_id = update.effective_chat.id
    index = user_data_store[chat_id]['current_edit_position_index']
    field_to_edit = user_data_store[chat_id]['current_edit_field']

    if 'attachment_edit_mode' in user_data_store[chat_id] and user_data_store[chat_id]['attachment_edit_mode']:
        choice = update.callback_query.data if update.callback_query else None
        
        if choice == "attach_file":
            await update.callback_query.edit_message_text("Отправьте новый файл:")
            user_data_store[chat_id]['temp_attachment_type'] = 'file'
            return EDIT_FIELD_INPUT # Остаемся в этом состоянии для получения файла
        elif choice == "attach_link":
            await update.callback_query.edit_message_text("Отправьте новую ссылку:")
            user_data_store[chat_id]['temp_attachment_type'] = 'link'
            return EDIT_FIELD_INPUT # Остаемся в этом состоянии для получения ссылки
        elif choice == "skip_attachment":
            user_data_store[chat_id]['positions'][index]['attachment_type'] = 'none'
            user_data_store[chat_id]['positions'][index]['attachment_content'] = ''
            del user_data_store[chat_id]['attachment_edit_mode']
            if 'temp_attachment_type' in user_data_store[chat_id]:
                del user_data_store[chat_id]['temp_attachment_type']
            await update.callback_query.edit_message_text("Вложение удалено.")
            await send_summary(chat_id, context)
            return EDIT_MENU
        
        # Если это не выбор прикрепления, а уже идет ввод файла/ссылки
        attachment_type_in_edit = user_data_store[chat_id].get('temp_attachment_type')
        if attachment_type_in_edit == 'file':
            if update.message and update.message.document:
                document = update.message.document
                file_id = document.file_id
                file_name = document.file_name
                user_data_store[chat_id]['positions'][index]['attachment_type'] = 'file'
                user_data_store[chat_id]['positions'][index]['attachment_content'] = f"file_id:{file_id}||file_name:{file_name}"
                await update.message.reply_text(f"Вложение (файл: {file_name}) обновлено.")
                del user_data_store[chat_id]['attachment_edit_mode']
                del user_data_store[chat_id]['temp_attachment_type']
                await send_summary(chat_id, context)
                return EDIT_MENU
            else:
                await update.message.reply_text("Пожалуйста, отправьте файл.")
                return EDIT_FIELD_INPUT
        elif attachment_type_in_edit == 'link':
            if update.message and update.message.text and (update.message.text.startswith('http://') or update.message.text.startswith('https://')):
                link = update.message.text
                user_data_store[chat_id]['positions'][index]['attachment_type'] = 'link'
                user_data_store[chat_id]['positions'][index]['attachment_content'] = link
                await update.message.reply_text(f"Вложение (ссылка: {link}) обновлено.")
                del user_data_store[chat_id]['attachment_edit_mode']
                del user_data_store[chat_id]['temp_attachment_type']
                await send_summary(chat_id, context)
                return EDIT_MENU
            else:
                await update.message.reply_text("Пожалуйста, введите корректную ссылку, начинающуюся с http:// или https://")
                return EDIT_FIELD_INPUT
        else: # Если был выбран edit_field_attachment, но не было выбора типа или ввода
            if update.callback_query:
                # Уже обработано выше для выбора attach_file/link/skip
                return EDIT_FIELD_INPUT
            elif update.message:
                await update.message.reply_text("Пожалуйста, выберите тип вложения (файл/ссылка) или пропустите.")
                return EDIT_FIELD_INPUT

    # Обработка выбора даты из календаря при редактировании
    if field_to_edit == "edit_field_delivery_date":
        if update.callback_query:
            query = update.callback_query
            data = query.data
            
            if data.startswith('CAL_PREV_MONTH'):
                parts = data.split('_')
                month = int(parts[3])
                year = int(parts[4])
                new_date = datetime(year, month, 1) - timedelta(days=1)
                calendar_markup = create_calendar(new_date.year, new_date.month)
                await query.edit_message_reply_markup(reply_markup=calendar_markup)
                return EDIT_FIELD_INPUT
            elif data.startswith('CAL_NEXT_MONTH'):
                parts = data.split('_')
                month = int(parts[3])
                year = int(parts[4])
                new_date = datetime(year, month, 1) + timedelta(days=31)
                calendar_markup = create_calendar(new_date.year, new_date.month)
                await query.edit_message_reply_markup(reply_markup=calendar_markup)
                return EDIT_FIELD_INPUT
            elif data.startswith('CAL_DATE'):
                date_str = data.split('_')[2]
                selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
                user_data_store[chat_id]['positions'][index]['delivery_date'] = selected_date.strftime('%d.%m.%Y')
                await query.edit_message_text(f"Дата доставки обновлена на: {selected_date.strftime('%d.%m.%Y')}")
                await send_summary(chat_id, context)
                return EDIT_MENU
            else:
                return EDIT_FIELD_INPUT
        else: # Если не было коллбэка календаря
            await update.message.reply_text("Пожалуйста, выберите дату из календаря.")
            return EDIT_FIELD_INPUT

    # Обработка текстового ввода для других полей
    if update.message and update.message.text:
        new_value = update.message.text
        if field_to_edit == "edit_field_quantity":
            try:
                new_value = float(new_value.replace(',', '.'))
            except ValueError:
                await update.message.reply_text("Неверное количество. Пожалуйста, введите числовое значение:")
                return EDIT_FIELD_INPUT
        
        # Обновляем соответствующее поле
        user_data_store[chat_id]['positions'][index][field_to_edit.replace('edit_field_', '')] = new_value
        await update.message.reply_text(f"Поле '{field_to_edit.replace('edit_field_', '')}' обновлено на: {new_value}")
        await send_summary(chat_id, context)
        return EDIT_MENU
    else:
        # Если это не текстовое сообщение и не коллбэк (например, пустой ввод)
        await update.message.reply_text("Неверный ввод. Пожалуйста, попробуйте еще раз.")
        return EDIT_FIELD_INPUT


async def process_global_calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает коллбэки календаря для выбора общей даты доставки."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id

    data = query.data

    if data.startswith('CAL_PREV_MONTH'):
        parts = data.split('_')
        month = int(parts[3])
        year = int(parts[4])
        new_date = datetime(year, month, 1) - timedelta(days=1)
        calendar_markup = create_calendar(new_date.year, new_date.month)
        await query.edit_message_reply_markup(reply_markup=calendar_markup)
        return GLOBAL_DELIVERY_DATE_SELECTION
    elif data.startswith('CAL_NEXT_MONTH'):
        parts = data.split('_')
        month = int(parts[3])
        year = int(parts[4])
        new_date = datetime(year, month, 1) + timedelta(days=31)
        calendar_markup = create_calendar(new_date.year, new_date.month)
        await query.edit_message_reply_markup(reply_markup=calendar_markup)
        return GLOBAL_DELIVERY_DATE_SELECTION
    elif data.startswith('CAL_DATE'):
        date_str = data.split('_')[2]
        selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        user_data_store[chat_id]['global_delivery_date'] = selected_date.strftime('%d.%m.%Y')
        await query.edit_message_text(f"Общая дата доставки установлена: {selected_date.strftime('%d.%m.%Y')}")
        await send_summary(chat_id, context) # Показываем обновленную сводку
        return EDIT_MENU
    else:
        # Это должен быть CAL_IGNORE, ничего не делаем
        return GLOBAL_DELIVERY_DATE_SELECTION


async def final_confirm_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обрабатывает окончательное подтверждение перед отправкой."""
    query = update.callback_query
    await query.answer()
    chat_id = query.message.chat.id
    choice = query.data

    if choice == "confirm_final_send":
        await query.edit_message_text("Ваша заявка отправляется...")
        await send_email_with_excel(chat_id, context)
        # Очистка данных после отправки
        if chat_id in user_data_store:
            del user_data_store[chat_id]
        if chat_id in temp_data_store:
            del temp_data_store[chat_id]
        if chat_id in temp_file_storage:
            del temp_file_storage[chat_id]
        return ConversationHandler.END
    # Отмена обрабатывается в cancel
    return FINAL_CONFIRMATION

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Отменяет текущий диалог."""
    chat_id = update.effective_chat.id
    await context.bot.send_message(
        chat_id=chat_id,
        text="Заявка отменена.",
        reply_markup=ReplyKeyboardRemove()
    )
    # Очистка данных
    if chat_id in user_data_store:
        del user_data_store[chat_id]
    if chat_id in temp_data_store:
        del temp_data_store[chat_id]
    if chat_id in temp_file_storage:
        del temp_file_storage[chat_id]
    return ConversationHandler.END

async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отвечает на неизвестные команды/сообщения."""
    await context.bot.send_message(chat_id=update.effective_chat.id, text="Извините, я не понял вашу команду или сообщение. Пожалуйста, используйте кнопки или начните заново с /start.")


# === Вспомогательные функции ===

def create_calendar(year: int = None, month: int = None) -> InlineKeyboardMarkup:
    """Создает клавиатуру календаря."""
    now = datetime.now()
    if year is None:
        year = now.year
    if month is None:
        month = now.month

    cal = calendar.Calendar()
    month_days = cal.monthdayscalendar(year, month)

    keyboard = []
    # Заголовок с месяцем и годом
    header = [
        InlineKeyboardButton("◀️", callback_data=f"CAL_PREV_MONTH_{month}_{year}"),
        InlineKeyboardButton(f"{datetime(year, month, 1).strftime('%B %Y')}", callback_data="CAL_IGNORE"),
        InlineKeyboardButton("▶️", callback_data=f"CAL_NEXT_MONTH_{month}_{year}")
    ]
    keyboard.append(header)

    # Дни недели
    week_days = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    keyboard.append([InlineKeyboardButton(day, callback_data="CAL_IGNORE") for day in week_days])

    # Дни месяца
    for week in month_days:
        row = []
        for day in week:
            if day == 0:
                row.append(InlineKeyboardButton(" ", callback_data="CAL_IGNORE"))
            else:
                date_str = f"{year}-{month:02d}-{day:02d}"
                row.append(InlineKeyboardButton(str(day), callback_data=f"CAL_DATE_{date_str}"))
        keyboard.append(row)
    
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel_dialog")])

    return InlineKeyboardMarkup(keyboard)


async def fill_excel(chat_id: int, context: ContextTypes.DEFAULT_TYPE, output_filename: str) -> None:
    """Заполняет шаблон Excel данными заявки."""
    user_data = user_data_store.get(chat_id, {})
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    today = date.today().strftime('%d.%m.%Y')
    project = user_data.get('project', '')
    object_name = user_data.get('object_name', '')
    user_full_name = f"{update.effective_user.first_name} {update.effective_user.last_name or ''}".strip() if update.effective_user else "N/A"
    telegram_id_or_username = f"@{update.effective_user.username}" if update.effective_user.username else str(update.effective_user.id)

    ws['H2'] = today
    ws['H3'] = project
    ws['H4'] = object_name
    ws['H5'] = user_full_name
    ws['H6'] = telegram_id_or_username

    global_delivery_date = user_data.get('global_delivery_date', '')
    if global_delivery_date:
        ws['H7'] = global_delivery_date # Общая дата доставки в H7

    start_row = 10 # Начальная строка для позиций
    
    for i, pos in enumerate(user_data.get('positions', [])):
        row = start_row + i
        ws[f'B{row}'] = i + 1 # Номер по порядку
        ws[f'C{row}'] = pos.get('name', '')
        ws[f'D{row}'] = pos.get('unit', '')
        ws[f'E{row}'] = pos.get('quantity', '')
        ws[f'F{row}'] = pos.get('module', '')
        
        # Дата доставки позиции, если есть, иначе общая дата
        position_delivery_date = pos.get('delivery_date')
        if position_delivery_date:
            ws[f'G{row}'] = position_delivery_date
        elif global_delivery_date:
            ws[f'G{row}'] = global_delivery_date

        attachment_type = pos.get('attachment_type', 'none')
        attachment_content = pos.get('attachment_content', '')

        if attachment_type == 'file':
            file_info = attachment_content.split('||')
            file_id = next((info.split(':')[1] for info in file_info if info.startswith('file_id:')), None)
            file_name = next((info.split(':')[1] for info in file_info if info.startswith('file_name:')), 'файл')
            if file_id:
                file_path = f"downloads/{file_name}"
                temp_file_storage[chat_id] = temp_file_storage.get(chat_id, []) + [file_path] # Сохраняем путь для последующего удаления
                try:
                    file = await context.bot.get_file(file_id)
                    await file.download_to_drive(file_path)
                    ws[f'H{row}'] = f"См. прикрепленный файл: {file_name}"
                except Exception as e:
                    logger.error(f"Ошибка при загрузке файла {file_name}: {e}")
                    ws[f'H{row}'] = f"Ошибка загрузки файла: {file_name}"
            else:
                ws[f'H{row}'] = "Неверные данные файла"
        elif attachment_type == 'link':
            ws[f'H{row}'] = attachment_content
        else:
            ws[f'H{row}'] = '' # Пусто, если нет вложения

    wb.save(output_filename)


async def send_email_with_excel(chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отправляет электронное письмо с заполненным файлом Excel и вложениями."""
    output_filename = f"Заявка_{chat_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # Создаем директорию для загрузок, если ее нет
    os.makedirs('downloads', exist_ok=True)

    try:
        await fill_excel(chat_id, context, output_filename)

        msg = EmailMessage()
        msg['Subject'] = f"Новая заявка на снабжение от Telegram ID: {chat_id}"
        msg['From'] = EMAIL_LOGIN
        msg['To'] = EMAIL_RECEIVER
        msg.set_content("Во вложении файл заявки и дополнительные вложения.")

        with open(output_filename, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=output_filename)

        # Прикрепляем файлы, которые были скачаны
        for file_path in temp_file_storage.get(chat_id, []):
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    file_name = os.path.basename(file_path)
                    msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=file_name)

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL_LOGIN, EMAIL_PASSWORD)
            smtp.send_message(msg)
        
        await context.bot.send_message(chat_id=chat_id, text="Заявка успешно отправлена на почту!")

    except Exception as e:
        logger.error(f"Ошибка при отправке письма: {e}")
        await context.bot.send_message(chat_id=chat_id, text=f"Произошла ошибка при отправке заявки: {e}")
    finally:
        # Очистка временных файлов
        if os.path.exists(output_filename):
            os.remove(output_filename)
        for file_path in temp_file_storage.get(chat_id, []):
            if os.path.exists(file_path):
                os.remove(file_path)
        if chat_id in temp_file_storage:
            del temp_file_storage[chat_id]
        
        # Удаление папки downloads, если она пуста
        if os.path.exists('downloads') and not os.listdir('downloads'):
            os.rmdir('downloads')


def main() -> None:
    """Запускает бота."""
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", initial_message_handler)],
        states={
            PROJECT: [CallbackQueryHandler(project_handler, pattern="^(Квартира|Офис|Другое)$")],
            OBJECT: [MessageHandler(filters.TEXT & ~filters.COMMAND, object_handler)],
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, name_handler)],
            UNIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, unit_handler)],
            QUANTITY: [MessageHandler(filters.TEXT & ~filters.COMMAND, quantity_handler)],
            MODULE: [CallbackQueryHandler(module_handler, pattern="^(?:[1-9]|1[0-8])$")],
            POSITION_DELIVERY_DATE: [CallbackQueryHandler(process_calendar_callback, pattern="^(CAL_|EDIT_CAL_)")],
            ATTACHMENT_CHOICE: [CallbackQueryHandler(attachment_choice_handler, pattern="^(attach_file|attach_link|skip_attachment)$")],
            FILE_INPUT: [MessageHandler(filters.Document.ALL & ~filters.COMMAND, file_input_handler)],
            LINK_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, link_input_handler)],
            CONFIRM_ADD_MORE: [CallbackQueryHandler(confirm_add_more_handler, pattern="^(add_more_position|finish_and_send)$")],

            EDIT_MENU: [CallbackQueryHandler(edit_menu_handler, pattern="^(edit_position|add_position_from_edit|delete_position|set_global_delivery_date|continue_and_send)$")],
            SELECT_POSITION: [CallbackQueryHandler(select_position_handler, pattern="^(select_pos_|delete_pos_).*")],
            EDIT_FIELD_SELECTION: [CallbackQueryHandler(edit_field_selection_handler, pattern="^edit_field_|^back_to_edit_menu$")],
            EDIT_FIELD_INPUT: [
                CallbackQueryHandler(edit_field_input_handler, pattern="^(CAL_|EDIT_CAL_)|^(attach_file|attach_link|skip_attachment)$"), # Для календаря и выбора вложения
                MessageHandler(filters.TEXT | filters.Document.ALL & ~filters.COMMAND, edit_field_input_handler),
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
                CallbackQueryHandler(process_global_calendar_callback, pattern="^(CAL_|EDIT_CAL_)\w*"),
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$")
            ],

            FINAL_CONFIRMATION: [
                CallbackQueryHandler(cancel, pattern="^cancel_dialog$"),
                CallbackQueryHandler(final_confirm_handler)
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CallbackQueryHandler(cancel, pattern="^cancel_dialog$"),
            MessageHandler(filters.COMMAND | filters.TEXT, unknown)
        ],
    )

    app.add_handler(conv_handler)

    app.add_handler(CommandHandler("start", initial_message_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, initial_message_handler))


    app.run_polling()

if __name__ == "__main__":
    main()