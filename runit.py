#!/usr/bin/env python3
"""
Скрипт для конвертации экспортированных данных из vaultwarden (JSON) в Excel формат.
"""

import json
import openpyxl
from openpyxl import Workbook


def load_json_data(filepath):
    """Загружает данные из JSON файла."""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_collection_id_to_name_map(collections):
    """Создает словарь для сопоставления ID коллекции с её названием."""
    return {coll['id']: coll['name'] for coll in collections}


def extract_items_data(data):
    """Извлекает данные из записей и сопоставляет их с коллекциями."""
    collection_map = create_collection_id_to_name_map(data.get('collections', []))
    items = data.get('items', [])

    result = []
    for item in items:
        name = item.get('name', '')
        note = item.get('notes', '')
        login_info = item.get('login', {})
        username = login_info.get('username', '') if login_info else ''
        password = login_info.get('password', '') if login_info else ''

        # Получаем список коллекций для этой записи
        collection_ids = item.get('collectionIds', [])

        if collection_ids:
            # Если у записи есть коллекции, создаем строку для каждой коллекции
            for coll_id in collection_ids:
                collection_name = collection_map.get(coll_id, '')
                result.append({
                    'collection_name': collection_name,
                    'name': name,
                    'note': note,
                    'username': username,
                    'password': password
                })
        else:
            # Если у записи нет коллекций, добавляем одну строку без названия коллекции
            result.append({
                'collection_name': '',
                'name': name,
                'note': note,
                'username': username,
                'password': password
            })

    return result


def write_to_excel(items_data, output_filepath):
    """Записывает данные в Excel файл."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Vaultwarden Export"

    # Заголовки столбцов
    headers = ['Наименование коллекции (collection_name)',
               'Наименование записи (name)',
               'Комментарий к записи (note)',
               'username (login)',
               'Пароль(password)']

    ws.append(headers)

    # Данные
    for item in items_data:
        ws.append([
            item['collection_name'],
            item['name'],
            item['note'],
            item['username'],
            item['password']
        ])

    # Автоподбор ширины столбцов
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_filepath)
    print(f"Файл успешно создан: {output_filepath}")


def main():
    input_file = 'example.json'
    output_file = 'vaultwarden_export.xlsx'

    # Загрузка данных
    print(f"Чтение файла {input_file}...")
    data = load_json_data(input_file)

    # Извлечение данных
    print("Обработка данных...")
    items_data = extract_items_data(data)

    # Запись в Excel
    print(f"Создание Excel файла {output_file}...")
    write_to_excel(items_data, output_file)

    print(f"Готово! Обработано записей: {len(items_data)}")


if __name__ == '__main__':
    main()