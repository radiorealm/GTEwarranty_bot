import telebot
from telebot import types
from fpdf import FPDF
import os
import tempfile
import smtplib
from email.message import EmailMessage

bot = telebot.TeleBot("7955924179:AAFiaz3iF4nSfx7PNaW9jNPnasidu1zYZ1k")

user_states = {}

questions = [
    ("sender_name", "Введите ФИО отправителя:"),
    ("company_name", "Введите название компании, отправляющей заявление:"),
    ("project_name", "Введите название проекта:"),
    ("engine_number", "Введите номер двигателя:"),
    ("project_number", "Введите номер проекта:"),
    ("unit_number", "Введите номер агрегата:"),
    ("moto_hours", "Введите наработку мотто-часов:"),
    ("start_count", "Введите количество стартов:"),
    ("problem_description", "Опишите проблему:")
]

SPARE_PARTS_STAGE = 'spare_parts_stage'

SMTP_SERVER = "smtp-mail.outlook.com"  # Замените на ваш SMTP сервер
SMTP_PORT = 587  # Обычно 587 для TLS
SMTP_USER = "warranty@gte.su"  # Ваш email
SMTP_PASSWORD = "7d2758Pz7"  # Ваш пароль или app password
# ============================

@bot.message_handler(commands=['start'])
def start(message):
    user_states[message.from_user.id] = {"step": 0, "data": {}}
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Начать заполнение")
    markup.add(btn1)
    bot.send_message(message.from_user.id, "Здравствуйте! Нажмите 'Начать заполнение' для подачи заявления.", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text == "Начать заполнение")
def begin_form(message):
    user_states[message.from_user.id] = {"step": 0, "data": {}}
    ask_next_question(message)

def ask_next_question(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    if state is None:
        return
    step = state["step"]
    if step < len(questions):
        key, question = questions[step]
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        skip_btn = types.KeyboardButton("Пропустить")
        markup.add(skip_btn)
        bot.send_message(user_id, question, reply_markup=markup)
    else:
        state[SPARE_PARTS_STAGE] = {"stage": 0, "current": {}, "parts": []}
        bot.send_message(user_id, "Введите данные о запасных частях. Если не требуется — нажмите 'Пропустить'.")
        ask_spare_part(message)

def ask_spare_part(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    sp = state.get(SPARE_PARTS_STAGE)
    if not sp:
        return
    stage = sp["stage"]
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    skip_btn = types.KeyboardButton("Пропустить")
    markup.add(skip_btn)
    if stage == 0:
        bot.send_message(user_id, "Каталожный номер:", reply_markup=markup)
    elif stage == 1:
        bot.send_message(user_id, "Наименование:", reply_markup=markup)
    elif stage == 2:
        bot.send_message(user_id, "Количество:", reply_markup=markup)
    elif stage == 3:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        yes_btn = types.KeyboardButton("Добавить ещё")
        no_btn = types.KeyboardButton("Закончить ввод")
        markup.add(yes_btn, no_btn)
        part = sp["current"]
        bot.send_message(user_id, f"Добавить ещё одну позицию?\nТекущая: {part['catalog']} | {part['name']} | {part['qty']}", reply_markup=markup)

@bot.message_handler(
    func=lambda message: (
        user_states.get(message.from_user.id, {}).get("step") is not None
        and user_states.get(message.from_user.id, {}).get("step") < len(questions)
        and SPARE_PARTS_STAGE not in user_states.get(message.from_user.id, {})
        and not user_states.get(message.from_user.id, {}).get("attaching_files")
    )
)
def handle_answers(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    if state is None:
        return
    step = state["step"]
    key, _ = questions[step]
    answer = message.text.strip()
    if answer.lower() == "пропустить" or answer == "-":
        answer = ""
    state["data"][key] = answer
    state["step"] += 1
    ask_next_question(message)

@bot.message_handler(func=lambda message: SPARE_PARTS_STAGE in user_states.get(message.from_user.id, {}))
def handle_spare_parts(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    sp = state.get(SPARE_PARTS_STAGE)
    if not sp:
        return
    text = message.text.strip()
    if text.lower() == "пропустить" or text == "-":
        state["data"]["spare_parts"] = []
        state.pop(SPARE_PARTS_STAGE, None)
        ask_next_question_after_spare_parts(message)
        return
    if sp["stage"] == 0:
        sp["current"] = {"catalog": text}
        sp["stage"] = 1
        ask_spare_part(message)
    elif sp["stage"] == 1:
        sp["current"]["name"] = text
        sp["stage"] = 2
        ask_spare_part(message)
    elif sp["stage"] == 2:
        sp["current"]["qty"] = text
        sp["stage"] = 3
        ask_spare_part(message)
    elif sp["stage"] == 3:
        if text.lower() == "добавить ещё":
            sp["parts"].append(sp["current"])
            sp["current"] = {}
            sp["stage"] = 0
            ask_spare_part(message)
        elif text.lower() == "закончить ввод":
            sp["parts"].append(sp["current"])
            state["data"]["spare_parts"] = sp["parts"]
            state.pop(SPARE_PARTS_STAGE, None)
            ask_next_question_after_spare_parts(message)
        else:
            ask_spare_part(message)

def ask_next_question_after_spare_parts(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    finish_btn = types.KeyboardButton("Завершить прикрепление")
    skip_btn = types.KeyboardButton("Пропустить")
    markup.add(finish_btn, skip_btn)
    bot.send_message(user_id, "Прикрепите необходимые файлы (фото или документы). Когда закончите — нажмите 'Завершить прикрепление' или 'Пропустить'.", reply_markup=markup)
    state["attaching_files"] = True
    state["files"] = []

@bot.message_handler(content_types=['document', 'photo'])
def handle_files(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    if not state or not state.get("attaching_files"):
        return
    if message.content_type == 'document':
        file_id = message.document.file_id
        file_name = message.document.file_name
        state["files"].append({"type": "document", "file_id": file_id, "file_name": file_name})
        bot.reply_to(message, f"Документ '{file_name}' добавлен.")
    elif message.content_type == 'photo':
        file_id = message.photo[-1].file_id
        # Скачиваем фото во временный файл
        file_info = bot.get_file(file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        temp_dir = tempfile.gettempdir()
        photo_path = os.path.join(temp_dir, f"{file_id}.jpg")
        with open(photo_path, 'wb') as new_file:
            new_file.write(downloaded_file)
        state["files"].append({"type": "photo", "file_id": file_id, "photo_path": photo_path})
        bot.reply_to(message, "Фото добавлено.")

@bot.message_handler(func=lambda message: user_states.get(message.from_user.id, {}).get("attaching_files"))
def handle_finish_attach(message):
    user_id = message.from_user.id
    state = user_states.get(user_id)
    if not state or not state.get("attaching_files"):
        return
    if message.text.strip().lower() in ["завершить прикрепление", "пропустить"]:
        state["data"]["attached_files"] = state.get("files", [])
        state.pop("attaching_files", None)
        state.pop("files", None)

        data = state["data"]
        file_path, temp_photos = generate_pdf(data)

        with open(file_path, "rb") as pdf_file:
            bot.send_document(user_id, pdf_file)

        # Формируем тему письма
        from datetime import datetime
        today = datetime.today().strftime('%d.%m.%Y')
        subject = f"Гарантийный случай {data.get('company_name', '-')} {today}"
        body = f"Поступило новое гарантийное заявление от {data.get('company_name', '-')}, проект: {data.get('project_name', '-')}."
        send_pdf_to_email(file_path, subject, body, "warranty@gte.su")

        bot.send_message(user_id, "Спасибо! Ваше заявление сформировано и отправлено в виде PDF.")
        user_states.pop(user_id, None)
        os.remove(file_path)
        # Удаляем временные фото
        for p in temp_photos:
            try:
                os.remove(p)
            except Exception:
                pass
        # Автоматический рестарт: предлагаем начать заново
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton("Начать заполнение")
        markup.add(btn1)
        bot.send_message(user_id, "Хотите подать новое заявление? Нажмите 'Начать заполнение'.", reply_markup=markup)

def generate_pdf(data):
    from datetime import datetime
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('Arial', '', 'arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', 'arialbd.ttf', uni=True)
    pdf.set_font('Arial', '', 11)

    today = datetime.today().strftime('%d.%m.%Y')

    # Заголовок
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, txt=f"Заявление о гарантийном случае от {today}", ln=1, align="C")

    pdf.set_font('Arial', '', 11)
    pdf.ln(3)

    # Таблица параметров в 4 колонки
    def table_row(label1, val1, label2, val2):
        pdf.cell(55, 8, label1, border=1)
        pdf.cell(45, 8, val1 or "-", border=1)
        pdf.cell(55, 8, label2, border=1)
        pdf.cell(0, 8, val2 or "-", border=1, ln=1)

    table_row("Название проекта", data.get("project_name", "-"),
              "Номер двигателя", data.get("engine_number", "-"))
    table_row("Номер проекта", data.get("project_number", "-"),
              "Номер агрегата", data.get("unit_number", "-"))
    table_row("Наработка мотто-часов", data.get("moto_hours", "-"),
              "Количество стартов", data.get("start_count", "-"))

    pdf.ln(5)

    # Описание проблемы
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 8, "Описание проблемы:", ln=1, border=0)
    pdf.multi_cell(0, 8, data.get("problem_description", "-"), border=1)

    pdf.ln(5)

    # Прилагаемые материалы
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 8, "Прилагаемые материалы:", ln=1, border=0)
    files = data.get("attached_files", [])
    temp_photos = []
    if files:
        photo_count = 1
        for f in files:
            if f['type'] == 'photo' and f.get('photo_path') and os.path.exists(f['photo_path']):
                try:
                    pdf.cell(0, 8, f"Фото {photo_count}", ln=1, border=1)
                    # Вставляем фото, ширина 100мм, сохраняем пропорции
                    pdf.image(f['photo_path'], w=100)
                    temp_photos.append(f['photo_path'])
                    photo_count += 1
                except Exception:
                    pdf.cell(0, 8, f"Фото {photo_count} (ошибка вставки)", ln=1, border=1)
                    photo_count += 1
            elif f['type'] == 'document':
                label = f.get("file_name") or f["file_id"]
                pdf.cell(0, 8, f"Документ: {label}", ln=1, border=1)
    else:
        pdf.cell(0, 8, "—", ln=1, border=1)

    pdf.ln(5)

    # Запасные части
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 8, "Запасные части, необходимые для устранения проблемы", ln=1, border=0)
    parts = data.get("spare_parts", [])
    if parts:
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(60, 8, "Каталожный номер", border=1)
        pdf.cell(80, 8, "Название", border=1)
        pdf.cell(30, 8, "Количество", border=1, ln=1)
        pdf.set_font('Arial', '', 11)
        for part in parts:
            pdf.cell(60, 8, part['catalog'], border=1)
            pdf.cell(80, 8, part['name'], border=1)
            pdf.cell(30, 8, part['qty'], border=1, ln=1)
    else:
        pdf.cell(0, 8, "—", ln=1, border=1)

    pdf.ln(10)

    # Подписи
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 8, "Подписи", ln=1, border=0)
    pdf.ln(3)

    company = data.get("company_name", "________________")
    sender = data.get("sender_name", "_____________________")

    pdf.cell(95, 8, company, ln=0)
    pdf.cell(0, 8, "ООО «ГринТех Энерджи»", ln=1)

    pdf.cell(95, 8, f"{sender}/______________/", ln=0)
    pdf.cell(0, 8, "_____________________/______________/", ln=1)

    pdf.cell(95, 8, f"Дата: {today}", ln=0)
    pdf.cell(0, 8, "", ln=1)

    file_path = f"zayavlenie_{data.get('engine_number', 'no_engine')}.pdf"
    pdf.output(file_path)
    return file_path, temp_photos

def send_pdf_to_email(pdf_path, subject, body, to_email):
    from_email = SMTP_USER
    if not all([SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD]):
        print('SMTP credentials are not set.')
        return
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    msg.set_content(body)
    with open(pdf_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
    except Exception as e:
        print(f'Error sending email: {e}')

if __name__ == '__main__':
    bot.polling(none_stop=True)
