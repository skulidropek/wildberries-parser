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
            return None, None, None, f"Товар с артикулом {article} не найден."
        
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
                    return (
                        f"{basic_price / 100:.2f} RUB",  # Базовая цена
                        f"{discount_price / 100:.2f} RUB",  # Цена со скидкой
                        f"{sale_price / 100:.2f} RUB",  # Итоговая цена
                        None  # Ошибок нет
                    )
        
        return None, None, None, f"Для артикула {article} отсутствуют доступные размеры с ценами."
    
    except requests.RequestException as e:
        return None, None, None, f"Ошибка при запросе: {str(e)}"

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
        
        # Создание новых колонок для результатов
        df["Базовая цена"] = ""
        df["Цена со скидкой"] = ""
        df["Итоговая цена"] = ""
        df["Ошибка"] = ""

        # Обработка цен
        for idx, article in enumerate(df.iloc[:, 0]):
            if pd.isna(article):
                df.loc[idx, "Ошибка"] = "Пустой артикул"
                print(f"Пустой артикул на строке {idx + 1}")
                continue
            
            # Получение цен
            basic_price, discount_price, total_price, error = get_prices_by_article(str(int(article)))
            df.loc[idx, "Базовая цена"] = basic_price
            df.loc[idx, "Цена со скидкой"] = discount_price
            df.loc[idx, "Итоговая цена"] = total_price
            df.loc[idx, "Ошибка"] = error

            if error:
                print(f"Артикул {article}: {error}")
            else:
                print(f"Артикул {article}: Базовая цена: {basic_price}, Цена со скидкой: {discount_price}, Итоговая цена: {total_price}")
        
        # Сохраняем результаты в новый Excel
        df.to_excel(output_file, index=False)
        print(f"Результаты сохранены в файл {output_file}")
    
    except Exception as e:
        print(f"Ошибка обработки файла: {str(e)}")

# Пример использования
if __name__ == "__main__":
    input_file = "input.xlsx"  # Входной файл
    output_file = "output.xlsx"  # Выходной файл
    
    process_excel(input_file, output_file)
