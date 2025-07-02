from fpdf import FPDF

pdf = FPDF()
pdf.add_page()
pdf.add_font('Arial', '', 'arial.ttf', uni=True)  # Подключаем ваш шрифт
pdf.set_font('Arial', '', 14)
pdf.cell(0, 10, 'Пример текста на русском', ln=True)
pdf.output('example.pdf')