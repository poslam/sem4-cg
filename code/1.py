import random
from datetime import datetime

from docxtpl import DocxTemplate

doc = DocxTemplate("./assets/in/1.docx")

now = datetime.now()

products = [
    "молоко",
    "колбаса",
    "сыр",
    "свиная вырезка",
    "хлеб",
    "масло",
    "яйца",
    "сахар",
    "соль",
    "перец",
    "макароны",
    "рис",
    "гречка",
    "пшено",
    "говядина",
    "курица",
    "креветки",
    "крабовые палочки",
    "консервы",
]

context = {
    "company": "ДВФУ",
    "seller": "FEFU",
    "address": "г. Владивосток, п. Аякс",
    "ORGN": random.randint(100000000, 999999999),
    "check_number": random.randint(100000000, 999999999),
    "day": str(now.day).zfill(2),
    "month": str(now.month).zfill(2),
    "year": now.year,
    "general_sum": random.randint(10000, 100000),
    "products": [
        {
            "title": random.choice(products),
            "code": random.randint(100000, 999999),
            "unit": random.choice(["шт.", "кг", "л", "мл", "г"]),
            "amount": random.randint(1, 10),
            "price": random.randint(1, 1000),
            "sum": random.randint(1, 1000),
        }
        for _ in range(15)
    ],
}

doc.render(context)

doc.save("./assets/out/1.docx")
