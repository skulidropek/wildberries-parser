import pandas as pd
import requests

def get_prices_by_article(article):
    """
    Получает три цены товара по артикулу Wildberries: текущая цена, цена со скидкой, базовая цена.
    
    :param article: Артикул товара (str или int)
    :return: Словарь с ценами или сообщение об ошибке
    """
    url = f"https://card.wb.ru/cards/v2/detail"
    params = {
        "appType": 1,
        "curr": "rub",
        "dest": 123586309,
        "spp": 30,
        "ab_testing": "false",
        "nm": article
    }

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        
        # Проверяем наличие данных о продукте
        products = data.get("data", {}).get("products", [])
        if not products:
            return f"Причина: Товар с артикулом {article} не найден."
        
        # Берем первый продукт из массива (если их несколько)
        product = products[0]
        sizes = product.get("sizes", [])
        
        # Проверяем цены для каждого размера
        for size in sizes:
            price_info = size.get("price", {})
            if price_info:
                discount_price = price_info.get("product")
                sale_price = price_info.get("total")
                basic_price = price_info.get("basic")
                if discount_price and sale_price and basic_price:
                    return {
                        "Цена со скидкой": f"{discount_price / 100:.2f} RUB",
                        "Итоговая цена (акция)": f"{sale_price / 100:.2f} RUB",
                        "Базовая цена": f"{basic_price / 100:.2f} RUB"
                    }
        
        return f"Причина: Для артикула {article} отсутствуют доступные размеры с ценами."
    
    except requests.RequestException as e:
        return f"Причина: Ошибка при запросе: {str(e)}"

def process_excel(input_file, output_file):
    """
    Обрабатывает Excel-файл: считывает артикули из первой колонки, получает цены и записывает их в новый файл.
    
    :param input_file: Путь к входному Excel-файлу
    :param output_file: Путь к выходному Excel-файлу
    """
    try:
        # Читаем Excel-файл
        df = pd.read_excel(input_file)
        
        # Убедимся, что колонка с артикулами существует
        if df.empty or df.columns[0] is None:
            print("Файл пустой или не содержит колонку с артикулами.")
            return
        
        # Обработка цен
        prices_list = []
        for article in df.iloc[:, 0]:
            if pd.isna(article):
                prices_list.append("Причина: Пустой артикул")
                print("Пустой артикул")
                continue
            
            # Получение цен
            result = get_prices_by_article(str(int(article)))
            if isinstance(result, dict):
                # Форматируем результат в строку
                prices_str = "; ".join([f"{key}: {value}" for key, value in result.items()])
                print(f"Артикул {article}: {prices_str}")
                prices_list.append(prices_str)
            else:
                print(f"Артикул {article}: {result}")
                prices_list.append(result)
        
        # Добавляем результаты в новый столбец
        df["Цены"] = prices_list
        
        # Сохраняем в новый Excel
        df.to_excel(output_file, index=False)
        print(f"Результаты сохранены в файл {output_file}")
    
    except Exception as e:
        print(f"Ошибка обработки файла: {str(e)}")

# Пример использования
if __name__ == "__main__":
    input_file = "input.xlsx"  # Входной файл
    output_file = "output.xlsx"  # Выходной файл
    
    process_excel(input_file, output_file)
