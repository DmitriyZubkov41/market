from time import time, sleep
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

start_time = time()

# настройки браузера
options = webdriver.ChromeOptions()
#options.add_argument("--headless") # для запуска в фоне
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

options.add_argument("--disable-cache")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

#options.add_argument("--disable-extensions")
options.add_argument("--disable-plugins")

# Создаем объект браузера (в данном случае Chrome)
browser = webdriver.Chrome(options=options)

browser.implicitly_wait(10)
#browser.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

# Открываем страницу с url:
url = "https://www.ozon.ru/category/knigi-16500/?category_was_predicted=true&deny_category_prediction=true&from_global=true&text=программирование"

while True:
   try:
       browser.set_page_load_timeout(60)
       browser.get(url)
       # элемент загруженной страницы <div class="b25_4_4-a"></div> (Вас заинтересует)
       try:
           locator = (By.CSS_SELECTOR, 'div[class="b25_4_4-a"]')
           WebDriverWait(browser, 20).until(EC.presence_of_element_located(locator))
           break
       except Exception as e:
           print(e)
           break
   except:
       print(f"снова открываем {url}")
       continue

# прокручиваем страницу в самый низ
count_scroll = 0  # счетчик прокручиваний
count_prev = 0 # счетчик элементов с предыдущего прокручивания
books = []
while True:
    # Прокручиваем страницу в самый низ
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    count_scroll +=1
    # Нужно подождать чтобы после прокрутки все элементы загрузились.
    sleep(5)

    # Прочитаем html страницы и сохраним в файл
    html = browser.page_source
    f = open(f'page{count_scroll}.html', 'w', encoding='utf-8')
    f.write(html)
    f.close()


    # Запишем в список books данные: title, price, img
    elements_books = browser.find_elements(By.CSS_SELECTOR, 'div[class^="tile-root i3s_24"]')
    count_all = len(elements_books)
    print("count_prev=", count_prev)
    print("count_all=", count_all)
    print(f"Прокручивание номер {count_scroll}\nВсего: {count_all} книг\nДобавлено: {count_all-count_prev} книг")
    first_book = elements_books[0].find_element(By.CSS_SELECTOR, 'span[class="tsBody500Medium"]').text
    last_book = elements_books[-1].find_element(By.CSS_SELECTOR, 'span[class="tsBody500Medium"]').text
    print(f"Первая книга {first_book}")
    print(f"Последняя книга {last_book}")
    

    for book in elements_books[count_prev::]:
        title = book.find_element(By.CSS_SELECTOR, 'span[class="tsBody500Medium"]').text
        l_title = title.lower()
        #if 'информатика' in l_title or 'егэ' in l_title or 'нлп' in l_title:
        #   continue
        print("Название книги", title)
        price = book.find_element(By.CSS_SELECTOR, 'span[class="c35_3_11-a1 tsHeadline500Medium c35_3_11-b1 c35_3_11-a6"]').text
        url = book.find_element(By.TAG_NAME, 'a').get_attribute('href')
        #img = book.find_element(By.CSS_SELECTOR, 'div[class="si8_24"]').text
        #bufer = BytesIO(urlopen(url).read())
        
        books.append(
            {
             'title': title,
             'url'  : url,
             'price': price
            }
        )
    count_prev = count_all   
    
    # Элемент-флаг конца прокрутки div class="mi8_24"
    try:
        browser.find_element(By.CSS_SELECTOR, 'div[class="mi8_24"]')
        print(f"Всё, страницу прокрутили\nПрокручивали {count_scroll} раз")
        break
    except:
        if count_scroll == 3:
            print("Закончили считывание")
            break



# Запишем список в xlsx:
# ТАБЛИЦА books.xlsx
from openpyxl import Workbook
# Создаем новую рабочую книгу (файл)
wb = Workbook()
# Получаем активный лист (по умолчанию создается один лист)
sheet = wb.active
# Устанавливаем имя листа
sheet.title = "Программирование"
        
# Заголовок
#sheet.row_dimensions[1].height = 27  # высота ячейки в 27 пикселей
sheet.append(["Название книги", "Адрес страницы", "Цена"])
# Заполняем таблицу ниже заголовка
for i in range(len(books)):
    row = i + 1
    sheet[f"A{row}"] = books[i]['title']
    sheet[f"B{row}"] = books[i]['url']
    sheet[f"C{row}"] = books[i]['price']
    #sheet[f"D{row}"] = books[i]['image']
    
# Сохраняем файл
wb.save('books.xlsx')
print("Сохранили файл")

# Завершаем работу
browser.quit()
print(f'Прошло времени: {time() - start_time} секунд')
