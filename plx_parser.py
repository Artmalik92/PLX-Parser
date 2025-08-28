import xml.etree.ElementTree as ET
from collections import defaultdict

def parse_plx(file_path):
    """
    Парсит .plx файл и возвращает структурированные данные.
    """
    # Указываем кодировку utf-16, так как файл в ней сохранен
    tree = ET.parse(file_path, parser=ET.XMLParser(encoding='utf-16'))
    root = tree.getroot()

    # Для простоты будем использовать пространства имен из файла
    # Определяем пространства имен, которые увидели в корневом теге
    namespaces = {
        'diffgr': 'urn:schemas-microsoft-com:xml-diffgram-v1',
        'msdata': 'urn:schemas-microsoft-com:xml-msdata',
        'ns': 'http://tempuri.org/dsMMISDB.xsd'  # Основное пространство имен для данных
    }

    # Вся полезная информация находится в тегах внутри <ns:dsMMISDB>
    data_dict = defaultdict(list)

    # Найдем все теги внутри dsMMISDB
    ds_mmisdb = root.find('.//ns:dsMMISDB', namespaces)
    
    if ds_mmisdb is not None:
        # Проходим по всем дочерним элементам (ООП, ПланыКомпетенции и т.д.)
        for child in ds_mmisdb:
            # Берем локальное имя тега (без пространства имен)
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            # Для каждого тега собираем атрибуты
            record = {}
            for attr_name, attr_value in child.attrib.items():
                # Также очищаем имена атрибутов от пространств имен
                clean_attr_name = attr_name.split('}')[-1] if '}' in attr_name else attr_name
                record[clean_attr_name] = attr_value
            
            # Добавляем запись в словарь по имени тега
            data_dict[tag_name].append(record)

    return data_dict

# Пример использования
if __name__ == "__main__":
    file_path = "21.03.03_Геодезия_2025_оч_.plx"  # Замените на путь к вашему файлу
    data = parse_plx(file_path)
    
    # Посмотрим, какие таблицы мы получили
    print("Найдены таблицы:", list(data.keys()))
    
    # Выведем, например, все учебные планы (ООП)
    for oop_record in data['ООП']:
        print(f"Шифр: {oop_record.get('Шифр', 'N/A')}, Название: {oop_record.get('Название', 'N/A')}")
    
    # Выведем компетенции
    for comp_record in data['ПланыКомпетенции']:
        print(f"Компетенция: {comp_record.get('ШифрКомпетенции', 'N/A')}, Описание: {comp_record.get('Наименование', 'N/A')}")

        