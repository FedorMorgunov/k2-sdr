#!/usr/bin/env python3
"""Создаёт пример Excel-файла contacts.example.xlsx с нужными колонками,
чтобы было видно ожидаемый формат входных данных."""

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Контакты"

ws.append(["Имя", "Компания", "Почта", "Тема письма"])
ws.append(["Иван", "ООО Ромашка", "ivan@romashka.ru", "Предложение о сотрудничестве"])
ws.append(["Мария", "АО Василёк", "m.petrova@vasilek.ru", "Сотрудничество с K2 SDR"])
ws.append(["Сергей", "ИП Сидоров", "sergey@sidorov.ru", "Короткий вопрос по вашему направлению"])

wb.save("contacts.example.xlsx")
print("Создан файл contacts.example.xlsx")
