import asyncio
import json
import datetime
from urllib.parse import urlencode, quote


import aiohttp
from openpyxl import Workbook
from playwright.async_api import async_playwright, expect


site = "https://www.wildberries.ru/"
search_url = "https://search.wb.ru/exactmatch/sng/common/v18/search?ab_testing=false&appType=1&curr=rub&dest=-59202&filters=ffsubject&lang=ru&%s&resultset=filters&spp=30&suppressSpellcheck=false"
preset_url = "https://catalog.wb.ru/catalog/promo/bucket_28/v8/filters?ab_testing=false&appType=1&curr=rub&dest=-59252&lang=ru&%s&spp=30"
result_ = {}
current_root = None
queue_for_task = []


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
        wb.save(f'result{now.strftime('%H%M%d%m%y.xlsx')}')


async def fetch(url, session, results, parent_category):
    try:
        async with session.get(url) as response:
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


tasks = []

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



def add_to_task(subcategory):
    queue_for_task.append(subcategory)


def cat(categories, level, result):
    """Проходит по категориям и присваивает уровни влоденности.
        Также передаем элеиенты в очередь для запроса подкатегорий
    """
    level = level + 1
    global current_root
    # print(result)
    for category in categories:
        print('::', category['id'])
        if level == 1:
            current_root = category
        result[current_root['id']] = result.get(current_root['id'], [])
        result[current_root['id']].append({'id': category['id'], 'name': category['name'], 'level': level})

        if category.get('childs'):
            cat(category['childs'], level, result)
        else:
            category['root_id'] = current_root['id']
            add_to_task(category)



async def get_all_categories_from_wb(url):
    """Данная функция при загрузке сайте получает json файл со списком категорий."""
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=False,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--start-maximized',
            ],
        )
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36',
            java_script_enabled=True
        )
        page = await context.new_page()
        await page.goto(url)
        async with page.expect_response('**/main-menu-by-ru-v3.json') as res:
            response_text =  await res.value
            all_categories = await response_text.json()
        await page.wait_for_timeout(3000)
        # local_categories зависят от местоположения. На данный момент не используются в дальнейшем
        local_categories = await page.locator('ul.menu-burger__main-list',).all_inner_texts()
        return all_categories, local_categories



async def main():
    (all_categories, local_categories) = await get_all_categories_from_wb(site)
    main_categories = [category for category in all_categories
                           # if category['name'] in local_categories[0].split()
                           # # and not category.get('dynamic', False)
                           #   and category['name'] != 'Wibes'
                             ]
    cat(main_categories, 0, result_)
    subcategories = await get_subcategories(queue_for_task)
    for key, val in subcategories.items():
        for i in val:
            result_[key].append(i)
    save(result_)


asyncio.run(main())