import pandas as pd
import requests

def get_prices_by_article(article):
    """
    Fetches prices for a product by its Wildberries article number: basic price, product price (discounted), 
    total price, logistics cost, and return value.
    
    :param article: Article number of the product (str or int)
    :return: Tuple containing prices or an error message
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
        
        # Check if product data exists
        products = data.get("data", {}).get("products", [])
        if not products:
            return None, None, None, None, None, f"Product with article {article} not found."
        
        # Take the first product in the list (if multiple exist)
        product = products[0]
        sizes = product.get("sizes", [])
        
        # Check prices for each size
        for size in sizes:
            price_info = size.get("price", {})
            if price_info:
                basic_price = price_info.get("basic")      # basic price
                discount_price = price_info.get("product") # product (discounted) price
                total_price = price_info.get("total")      # total price including logistics
                logistics = price_info.get("logistics")    # logistics cost
                return_value = price_info.get("return")    # return cost
                
                if basic_price and discount_price and total_price:
                    return (
                        f"{basic_price / 100:.2f} RUB",    # basic price
                        f"{discount_price / 100:.2f} RUB", # product (discounted) price
                        f"{total_price / 100:.2f} RUB",    # total price
                        f"{logistics / 100:.2f} RUB",      # logistics cost
                        f"{return_value / 100:.2f} RUB",   # return cost
                        None                                # No error
                    )
        
        return None, None, None, None, None, f"No available sizes with prices for article {article}."
    
    except requests.RequestException as e:
        return None, None, None, None, None, f"Request error: {str(e)}"

def process_excel(input_file, output_file):
    """
    Processes an Excel file: reads article numbers from the first column, fetches prices, and writes results to a new file.
    
    :param input_file: Path to the input Excel file
    :param output_file: Path to the output Excel file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        # Ensure the column with articles exists
        if df.empty or df.columns[0] is None:
            print("The file is empty or does not contain an article column.")
            return
        
        # Create new columns for results based on JSON field names
        df["basic"] = ""
        df["product"] = ""
        df["total"] = ""
        df["logistics"] = ""
        df["return"] = ""
        df["error"] = ""

        total_items = len(df)  # Total number of articles
        # Process prices
        for idx, article in enumerate(df.iloc[:, 0]):
            current_item = idx + 1  # Current item
            if pd.isna(article):
                df.loc[idx, "error"] = "Empty article"
                print(f"[{current_item}/{total_items}] Empty article")
                continue
            
            # Fetch prices
            basic_price, discount_price, total_price, logistics, return_value, error = get_prices_by_article(str(int(article)))
            df.loc[idx, "basic"] = basic_price
            df.loc[idx, "product"] = discount_price
            df.loc[idx, "total"] = total_price
            df.loc[idx, "logistics"] = logistics
            df.loc[idx, "return"] = return_value
            df.loc[idx, "error"] = error

            if error:
                print(f"[{current_item}/{total_items}] Article {article}: {error}")
            else:
                print(f"[{current_item}/{total_items}] Article {article}: Basic: {basic_price}, Product (Discounted): {discount_price}, Total: {total_price}, Logistics: {logistics}, Return: {return_value}")
        
        # Save results to a new Excel file
        df.to_excel(output_file, index=False)
        print(f"Results saved to file {output_file}")
    
    except Exception as e:
        print(f"Error processing file: {str(e)}")

# Example usage
if __name__ == "__main__":
    input_file = "input.xlsx"  # Input file
    output_file = "output.xlsx"  # Output file
    
    process_excel(input_file, output_file)
