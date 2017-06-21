# -*- encoding:  utf-8 -*-

"""
读取高校xls文件，转换成json数据
"""

import xlrd
import json
import re

def convert(filename="./university.xls"):
    book = xlrd.open_workbook(filename)
    sheets = book.sheets()
    result = {
        'province': [],
        'city': [],
        'university': []
    }
    for sheet in sheets:
        province_id = None
        city_id = None
        cities = []
        for row in range(0, sheet.nrows):
            if row < 3:
                continue
            for col in range(0, sheet.ncols):
                cell = sheet.cell(row, col)
                if col == 0 and cell.ctype == 1:
                    # 省份格式如: 河北省（121所）
                    if re.search(ur"^\W+（\d+所）$", cell.value):
                        if province_id is None:
                            province_id = 1
                        else:
                            province_id += 1
                        province_name = re.sub(ur"（\d+所）", u'', cell.value)
                        province_obj = {
                            'id': province_id,
                            'name': province_name
                        }
                        if province_obj not in result['province']:
                            result['province'].append(province_obj)

                if col == 4 and cell.ctype == 1:
                    if cell.value not in cities:
                        if len(cities) == 0:
                            city_id = 1
                        else:
                            city_id += 1
                        cities.append(cell.value)
                        city_obj = {
                            'id': city_id,
                            'name': cell.value,
                            'pid': province_id
                        }
                        result['city'].append(city_obj)

                    result['university'].append({
                        'id': sheet.cell(row, 2).value,
                        'name': sheet.cell(row, 1).value,
                        'level': sheet.cell(row, 5).value,
                        'type': sheet.cell(row, 6).value or u'公办',
                        'cid': city_id
                    })

    return result

def save_json(json_data={}):
    if json_data:
        with open('./data.json', 'w') as f:
            f.write(json.dumps(json_data))

def main():
    json_data = convert()
    save_json(json_data)

if __name__ == '__main__':
    main()
