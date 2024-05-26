from datetime import datetime

from docxtpl import DocxTemplate

doc = DocxTemplate("./assets/in/3.docx")

now = datetime.now()

all_guests = [
        {"name": "Иванов Иван Иванович", "relation": "друг жениха",},
        {"name": "Петров Петр Николавевич", "relation": "напарник жениха в доте",},
        {"name": "Сергеевко Лариса Витальевна", "relation": "мастер по ноготочкам невесты",},
        {"name": "Маврения Диана Кирилловна", "relation": "подруга невесты",},
        {"name": "Петровна Василила Американовна", "relation": "мать невесты",},
    ]

context = {
    "guest": "",
    "whose_wedding": "Некрика Кирилла Андреева и Мариновой Ларисы Олеговны",
    "wedding_date": "12 июня 2024 года",
    "wedding_time": "12:00",
    "wedding_place": "ресторан \"У Андрея\" на Светланской 25",
    "guests": []
}

for guest in all_guests:
    context["guest"] = guest["name"]
    
    for _guest in all_guests:
        if _guest["name"] == guest["name"]:
            continue
        else:
            context["guests"].append(_guest)
            
    doc.render(context)
    doc.save(f"./assets/out/3/{guest["name"]}.docx")

    context["guests"] = []
