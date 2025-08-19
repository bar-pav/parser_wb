import asyncio
import json
import datetime
from urllib.parse import urlencode, quote


import aiohttp
from openpyxl import Workbook
from playwright.async_api import async_playwright, expect


# TODO: при запросе подкатегорий может оказаться, что в ответе нет поля 'data'. нужно повторить запрос на другой адрес в переменной second_preset_url.
# TODO: в результирующем файле excel будут повторяться подкатегории с вложенностью 99. Добавить проверку на наличие категории в результирующем списке.

site = "https://www.wildberries.ru/"
search_url = "https://search.wb.ru/exactmatch/sng/common/v18/search?ab_testing=false&appType=1&curr=rub&dest=-59202&filters=ffsubject&lang=ru&%s&resultset=filters&spp=30&suppressSpellcheck=false"
preset_url = "https://catalog.wb.ru/catalog/promo/bucket_28/v8/filters?ab_testing=false&appType=1&curr=rub&dest=-59252&lang=ru&%s&spp=30"
second_preset_url = "https://search.wb.ru/exactmatch/sng/common/v18/search?ab_testing=false&appType=1&autoselectFilters=false&curr=rub&dest=-59202&inheritFilters=false&lang=ru&%s&resultset=filters&spp=30&suppressSpellcheck=false"
categories_url = "https://static-basket-01.wbbasket.ru/vol0/data/main-menu-by-ru-v3.json"
result_ = {}
current_root = None
queue_for_task = []
tasks = []



def save(data):
    wb =Workbook()
    header = ['ID', 'Name', 'Level']
    for key, values in data.items():
        sheet = wb.create_sheet(values[0]['name'])
        wb.active = sheet
        sheet.append(header)
        for v in values:
            sheet.append([v['id'], v['name'], v['level']])
    wb.remove(wb['Sheet'])
    try:
        wb.save('result.xlsx')
    except:
        now = datetime.datetime.now()
        wb.save(f"result{now.strftime('%H%M%d%m%y.xlsx')}")


async def fetch(url, session, results, parent_category):

    async with session.get(url) as response:
        try:
            categories_json = json.loads(await response.text())
            results[parent_category['root_id']] = results.get(parent_category['root_id'], [])
            if categories_json['data']['filters'][0]['name'] == 'Категория':
                for item in categories_json['data']['filters'][0]['items']:
                    results[parent_category['root_id']].append({'id': item['id'], 'name': item['name'], 'level': 99})
            else:
                for item in categories_json['data']['filters']:
                    if item['name'] == 'Категория':
                        for i in item['items']:
                            results[parent_category['root_id']].append(
                                {'id': i['id'], 'name': i['name'], 'level': 99})
        except Exception as e:
            print(f'Error in ({url}):', e)


async def get_subcategories(queue_for_task=None):
    '''Создает сессию и запускает асинхронно запросы'''
    connector = aiohttp.TCPConnector(limit=50)
    results = {}
    async with aiohttp.ClientSession(connector=connector) as session:
        for i in queue_for_task:
            try:
                param_search_query = i.get('searchQuery')
                if param_search_query:
                    param = urlencode({'query': param_search_query}, quote_via=quote)
                    url = search_url % param
                else:
                    param_preset = i.get('query')
                    if param_preset and param_preset.startswith('preset'):
                        param = urlencode({'preset': param_preset.split('=')[1]}, quote_via=quote)
                        url = preset_url % param
                    else: continue
                task = asyncio.create_task(fetch(url, session, results, i))
                tasks.append(task)  
            except Exception as e:
                print('Error:', i, e)
        await asyncio.gather(*tasks)
    return results


def cat(categories, level, result):
    """Проходит по категориям и присваивает уровни вложенности.
        Также передаем элементы в очередь для запроса подкатегорий
    """
    level = level + 1
    global current_root
    for category in categories:
        if level == 1:
            current_root = category
        result[current_root['id']] = result.get(current_root['id'], [])
        result[current_root['id']].append({'id': category['id'], 'name': category['name'], 'level': level})

        if category.get('childs'):
            cat(category['childs'], level, result)
        else:
            category['root_id'] = current_root['id']
            queue_for_task.append(category)


async def get_all_categories_from_wb(url):
        async with aiohttp.ClientSession() as session:
            async with session.get(categories_url) as response:
                all_categories = await response.json()
        return all_categories



async def main():
    all_categories = await get_all_categories_from_wb(site)
    cat(all_categories, 0, result_)
    subcategories = await get_subcategories(queue_for_task)
    for key, val in subcategories.items():
        for i in val:
            result_[key].append(i)
    save(result_)
asyncio.run(main())