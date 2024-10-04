import openpyxl
import requests

# URL category_level_1: https://api.tiki.vn/raiden/v2/menu-config?platform=desktop
# URL get category_child: https://tiki.vn/api/v2/categories?include=children&parent_id={parent_id}
# URL get products of category: https://tiki.vn/api/v2/products?limit={limit}&category={category}&page={page}  tối đa lấy ra là 40 sản phẩm mỗi page
################################ => "paging": {
    #     "total": 2000,
    #     "total_text": "10000+",
    #     "per_page": 10,
    #     "current_page": 1,
    #     "last_page": 200,
    #     "from": 1,
    #     "to": 10
    # },

def get_categories_level_1():
    url = 'https://api.tiki.vn/raiden/v2/menu-config?platform=desktop'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0',
    }
    response = requests.get(url, headers=headers)
    metadata = response.json().get('menu_block', {}).get('items', [])
    level_1_categories = []
    for data in metadata:
        name = data.get('text')
        category_id = data.get('link').split('/')[-1].lstrip('c')
        level_1_categories.append((name, category_id))
    print('Get category level 1 successfully')
    return level_1_categories

def get_categories(parent_id):
    print('Start get categories:::', parent_id)
    url = f'https://tiki.vn/api/v2/categories?include=children&parent_id={parent_id}'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0',
    }
    response = requests.get(url, headers=headers)
    print(f"Get {parent_id} successfully")
    return response.json().get('data', [])

def build_category_tree(worksheet, parent_category, parent_path=[]):
    name, category_id = parent_category
    current_path = parent_path + [(name, category_id)]
    children = get_categories(category_id)

    if children:
        for child in children:
            child_name = child.get('name')
            child_id = child.get('id')
            build_category_tree(worksheet, (child_name, child_id), current_path)
    else:
        row = []
        for name, category_id in current_path:
            row.extend([name, category_id])
        while len(row) < 6:
            row.extend(['', ''])
        worksheet.append(row)


def main():
    wb = openpyxl.Workbook()
    worksheet = wb.active
    worksheet.title = "Category Hierarchy"
    worksheet.append(['category_level_1_name', 'id_1',
                      'category_level_2_name', 'id_2',
                      'category_level_3_name', 'id_3',
                      'category_level_4_name', 'id_4',
                      'category_level_5_name', 'id_5'])
    
    level_1_categories = get_categories_level_1()
    for level_1_name, level_1_id in level_1_categories:
        build_category_tree(worksheet, (level_1_name, level_1_id))
    print('---------GET CATEGORY SUCCESSFULLY-------------')
    wb.save('FullCategory.xlsx')

if __name__ == '__main__':
    main()
