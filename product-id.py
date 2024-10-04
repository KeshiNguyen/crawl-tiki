import requests
import openpyxl
import pandas as pd
from collections import defaultdict
import re


def remove_illegal_char(text):
    if isinstance(text, str):
        return ''.join(c for c in text if c.isprintable())
    return text
def read_file_excel(filepath):
    df = pd.read_excel(filepath, header=None, skiprows=1)  # Bỏ qua dòng đầu tiên
    category_dict = defaultdict(list)

    # Nhóm các ID theo danh mục gốc
    for index, row in df.iterrows():
        category_name = row[0]  # Tên danh mục gốc
        if row.dropna().iloc[-1]:
            category_id = int(row.dropna().iloc[-1])  # ID cuối cùng là số
            category_dict[category_name].append(category_id)  # Thêm ID vào danh sách của danh mục gốc

    return category_dict

def get_products_limit_category(category_id): #max 40 product per sub-category_id
    try:
        print(f'Get products for category {category_id}')
        url = f'https://tiki.vn/api/v2/products?limit=40&category={category_id}&page=1'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0',
        }
        response = requests.get(url, headers=headers)
        return response.json().get('data', [])
    except Exception as e:
        print(f"Error occurred while getting products for category {category_id}: {str(e)}")
        return []

def save_products_to_excel(products, output_file):
    wb = openpyxl.Workbook()
    worksheet = wb.active
    worksheet.title = "Products"
    worksheet.append(['STT', 'id', 'Tên sản phẩm', "category_path", 'Giá', 'Image', 'availability', 'URL'])
    for index, product in enumerate(products, start=1):
        row = [
            index,
            product.get('id'),
            remove_illegal_char(product.get('name')),
            product.get('primary_category_path'),
            product.get('price'),
            product.get('thumbnail_url'),
            product.get('availability'),
            product.get('url_path')
        ]
        worksheet.append(row)
        print(f"Saved product information {product.get('id')} to {output_file} successfully")
    wb.save(output_file)
    print(f"All Products saved to '{output_file}' successfully.")

def main():
    input_file = 'FullCategory.xlsx'
    category_dict = read_file_excel(input_file)

    for category_name, category_ids in category_dict.items():
        all_products = []
        for category_id in category_ids:
            products = get_products_limit_category(category_id)
            all_products.extend(products)

        # Tạo tên file từ danh mục
        clean_category_name = re.sub(r'[^\w\s]', '', category_name)  # Loại bỏ ký tự đặc biệt
        output_file = f"{clean_category_name.strip().replace(' ', '_')}.xlsx"
        save_products_to_excel(all_products, output_file)

if __name__ == "__main__":
    main()

