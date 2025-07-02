import telebot
from telebot import types
from fpdf import FPDF
import os

bot = telebot.TeleBot("7955924179:AAFiaz3iF4nSfx7PNaW9jNPnasidu1zYZ1k")

user_states = {}

questions = [
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
        bot.send_message(user_id, "Название:", reply_markup=markup)
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
        state["files"].append({"type": "photo", "file_id": file_id})
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
        file_path = generate_pdf(data)

        with open(file_path, "rb") as pdf_file:
            bot.send_document(user_id, pdf_file)

        bot.send_message(user_id, "Спасибо! Ваше заявление сформировано и отправлено в виде PDF.")
        user_states.pop(user_id, None)
        os.remove(file_path)

# Функция генерации PDF
def generate_pdf(data):
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('Arial', '', 'arial.ttf', uni=True)
    pdf.set_font('Arial', '', 12)

    def write_line(label, value=""):
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(60, 10, txt=label, ln=0)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, txt=value if value else "-")

    pdf.cell(0, 10, txt="Заявление о гарантийном случае", ln=1, align="C")

    write_line("Название компании:", data.get("company_name", "-"))
    write_line("Название проекта:", data.get("project_name", "-"))
    write_line("Номер двигателя:", data.get("engine_number", "-"))
    write_line("Номер проекта:", data.get("project_number", "-"))
    write_line("Номер агрегата:", data.get("unit_number", "-"))
    write_line("Наработка мотто-часов:", data.get("moto_hours", "-"))
    write_line("Количество стартов:", data.get("start_count", "-"))
    write_line("Описание проблемы:", data.get("problem_description", "-"))

    pdf.ln(5)
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(0, 10, txt="Запасные части", ln=1)
    pdf.set_font("Arial", size=12)

    parts = data.get("spare_parts", [])
    if parts:
        for part in parts:
            pdf.cell(0, 10, txt=f"{part['catalog']} | {part['name']} | {part['qty']}", ln=1)
    else:
        pdf.cell(0, 10, txt="—", ln=1)

    pdf.ln(10)
    pdf.cell(0, 10, txt="Подпись: _______________________", ln=1)
    pdf.cell(0, 10, txt="Дата: __________________________", ln=1)

    file_path = f"/tmp/zayavlenie_{data.get('engine_number', 'no_engine')}.pdf"
    pdf.output(file_path)
    return file_path

if __name__ == '__main__':
    bot.polling(none_stop=True)
