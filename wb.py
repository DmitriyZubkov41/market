from time import time, sleep
from urllib.request import urlopen
from io import BytesIO

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import asyncio
from aiohttp import ClientSession
#from aiohttp.exceptions import ConnectionTimeoutError 
from openpyxl import Workbook
from openpyxl.drawing.image import Image as oImage
from openpyxl.styles import Alignment
from PIL import Image as pImage


def settings_browser():
    # Настраиваем браузер и возвращаем его
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-cache")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 YaBrowser/25.6.0.0 Safari/537.36")
    
    # Создаем объект браузера (в данном случае Chrome)
    browser = webdriver.Chrome(options=options)
    
    return browser


def open_page(browser, url, locator, lst_data, count_page):
    """
    Открываем страницу по адресу url браузером browser.
    Убеждаемся, что страница полностью загрузилась по локатору locator.
    Полностью её прокручиваем, чтобы все элементы загрузились.
    Возвращаем список данных.
    """
    count_page +=1
    count_attempt = 1
    while True:
        try:
            browser.set_page_load_timeout(70)
            if count_attempt == 1:
                print("\n==============Открываю страницу номер ", count_page, " ===================")
            browser.get(url)
            # locator - флаг загруженной страницы и одновременно элемент к которому прокручиваем страницу
            WebDriverWait(browser, 60).until(EC.presence_of_element_located(locator))
            locator_act_page = (By.CSS_SELECTOR, 'span[class="pagination-item pagination__item active"]')
            WebDriverWait(browser, 40).until(EC.presence_of_element_located(locator_act_page))
            cur_page = browser.find_element(*locator_act_page).text
            print(f'Страницу № {cur_page} открыли')
            if count_page > 59:
                html = browser.page_source
                f = open(f'page{count_page}.html', 'w', encoding='utf-8')
                f.write(html)
                f.close()
            break
        except Exception as e:
            print("_______________Исключение:__________")
            print(e)
            print("________________________________________")
            try:
                browser.find_element(By.XPATH, '//h1[text()="Ничего не нашлось по запросу "]')
                # p class="not-found-search__text"
                print("Больше поиск ничего не выдаёт, заканчиваем считывание страниц")
                print("\nВсе страницы прочитали\nВсего их: ", count_page)
                print("Всего элементов-книг:", len(lst_data))
                return lst_data
            except:
                print("Страница не загрузилась или локатор не обнаружен.")
            count_attempt +=1
            print("Пауза 10 секунд")
            sleep(10)
            try:
                if count_attempt > 5:
                    html = browser.page_source
                    f = open(f'page{count_page}.html', 'w', encoding='utf-8')
                    f.write(html)
                    f.close()
            except Exception as e:
                print("_________Исключение при попытке сохранить страницу_____")
                print(e)
                print("====================================\n")
            print("Открываем страницу, попытка номер ", count_attempt)
        
    '''
    прокручиваем страницу к locator,
    если количество книг после прокручивания и после предыдущего прокручивания одинаковые,
    то заканчивает скроллинг.
    '''
    count_scroll = 0  # счетчик прокручиваний
    count_prev = 0    # счетчик элементов - книг с предыдущего прокручивания
    while True:
        bottom_element = browser.find_element(*locator)
        # Используем ActionChains для перемещения курсора к элементу
        actions = ActionChains(browser)
        actions.move_to_element(bottom_element).perform()
        count_scroll +=1
        print(f"\nПрокручивание номер {count_scroll}")
        print("Пауза 1 секунда")
        sleep(1)
        
        elements_books = browser.find_elements(By.CSS_SELECTOR, 'article[class="product-card j-card-item j-analitics-item"]')
        count_all = len(elements_books)
        print(f"Всего: {count_all} книг\nДобавлено: {count_all-count_prev} книг")
        
        if count_all == count_prev:
            print("\nСтраница прокрутилась, запишем данные в список lst_data")
            elements_books = browser.find_elements(By.CSS_SELECTOR, 'article[class="product-card j-card-item j-analitics-item"]')
            for book in elements_books:
                element = book.find_element(By.CSS_SELECTOR, 'span[class="product-card__name"]').text
                if element == '':
                    title = book.find_element(By.CSS_SELECTOR, 'a[class^="product-card__link"]').get_attribute("aria-label")
                elif element[0] == '/':
                    title = element[2:]
                else:
                    title = element
                #print("Название книги:", title)
                                
                url_book = book.find_element(By.CSS_SELECTOR, 'a[class="product-card__link j-card-link j-open-full-product-card"]').get_attribute('href')
                #print(url_book)
                try:
                    price = int(book.find_element(By.CSS_SELECTOR, 'ins[class^="price__lower-price wallet-price"]').text[:-2].replace(' ', ''))
                #print("price=", price)
                except:
                    price = 0
                # URL картинки
                url_img = book.find_element(By.CSS_SELECTOR, 'img').get_attribute("src")
                                
                lst_data.append(
                    {
                     'title': title,
                     'url'  : url_book,
                     'price': price,
                     'img_url': url_img
                    }
                )
            print("Сформировали список со страницы")

            try:
                print("Проверяем есть ли еще страница, которую надо посетить")
                locator_next_page = (By.CSS_SELECTOR, 'a[class="pagination-next pagination__next j-next-page"]')
                WebDriverWait(browser, 50).until(EC.presence_of_element_located(locator_next_page))
                next_url = browser.find_element(*locator_next_page).get_attribute('href')
                print("Следующая страница есть, перейдем на неё")
                return open_page(browser, next_url, locator, lst_data, count_page)
                '''
                if count_page == 2:
                    return lst_data
                else:
                    return open_page(browser, next_url, locator, lst_data, count_page)
                '''
            except Exception as e:
                print(f"______________Ошибка или все страницы прочитали______________\n{e}")
                print("\nВсе страницы прочитали\nВсего их: ", count_page)
                print("Всего элементов-книг:", len(lst_data))
                return lst_data
            
        #Страница еще до конца не прокрутилась
        else:
            count_prev = count_all
       

async def write_webp(session, sheet, url, row):
    try:
        async with session.get(url) as response:
            image_in_bufer = BytesIO(await response.read())
            bufer_png = pImage.open(image_in_bufer).convert("RGB")
            img_stream = BytesIO()
            bufer_png.save(img_stream, format='PNG')
            img_openpyxl = oImage(img_stream)
            img_openpyxl.height = 220
            img_openpyxl.width = 190
            sheet.add_image(img_openpyxl, f"D{row}")
    except Exception as e:
        print(f"_______Исключение при получении картинки с {url}")
        print(e)
        print("____________________________________________________")
        sheet[f"D{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet[f"D{row}"] = f"(Картинку не получили\nurl: {url}"


async def write_to_xlsx(lst_books):
    """
    Запишем список lst_books в books_wb.xlsx:
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Программирование"
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 32
    sheet.column_dimensions['c'].width = 6
    sheet.column_dimensions['D'].width = 21 
    # Заголовок
    sheet.append(["Название книги", "Адрес страницы", "Цена", "Картинка"])
    # Заполняем таблицу ниже заголовка
    for i in range(len(lst_books)):
        row = i + 2
        # Выравнивание текста по верху слева
        sheet[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet[f"B{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        sheet[f"C{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        # Высота строки
        sheet.row_dimensions[row].height = 170
        # Записываем данные
        sheet[f"A{row}"] = lst_books[i]['title']
        sheet[f"B{row}"] = lst_books[i]['url']
        sheet[f"C{row}"] = lst_books[i]['price']

    # Асинхронно открываем картинки по url, конвертируем webp в png и записываем в столбец D
    async with ClientSession() as session:
        tasks = [write_webp(session, sheet, lst_books[i]['img_url'], i+2) for i in range(len(lst_books))]
        await asyncio.gather(*tasks)
    
    wb.save('books_wb.xlsx')
    print("Записали в файл")


def main():
    start_time = time()
    browser = settings_browser()
    
    url_start = 'https://www.wildberries.ru/catalog/0/search.aspx?search=книги%20программирование'
    locator = (By.CSS_SELECTOR, 'h2[class="search-tags__header section-header"]')
    lst_data = open_page(browser, url_start, locator, lst_data=[], count_page=0)
    
    #print(f"Запишем в файл books_wb.xlsx {len(lst_data)} книг")
    #asyncio.run(write_to_xlsx(lst_data))
    
    import database
    database.write_db(lst_data)

    browser.quit()

    end_time = time()
    hours = int(round(end_time - start_time, 0)) // 3600
    minutes = (int(round(end_time - start_time, 0)) % 3600) // 60
    seconds = int(round(end_time - start_time, 0)) - hours * 3600 - minutes * 60
    print(f'Время, потраченное на выполнение программы: {hours} часов, {minutes} минуты, {seconds} секунд')
    

main()
