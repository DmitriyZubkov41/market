from io import BytesIO
from urllib.request import urlopen

from sqlalchemy import text
from sqlalchemy import create_engine
import pandas as pd
from PIL import Image as pImage

 
# Пользователь dmitriy, пароль: password, ip: localhost, порт: 5432, база данных: marketplaces
engine = create_engine('postgresql://dmitriy:passw@localhost:5432/market', echo=False)

lst_words = ['нейро', 'ai', ' ml', 'аналит', 'робототехник']
lst_results = []
with engine.connect() as conn:
    for word in lst_words:
        qr = f"SELECT title, price, url, url_img FROM unik_books_program WHERE LOWER(title) LIKE '%{word}%'"
        result = conn.execute(text(qr))
        for row in result:
            #print(f"Название книги: {row.title}  цена: {row.price}")
            lst_results.append(
                {
                  'title': row.title,
                  'url'  : row.url,
                  'price': row.price,
                  'img_url': row.url_img
                }
            )
print(len(lst_results))


# Запишем список в файл используя pandas

# Create a Pandas dataframe from some data.
full_df = pd.DataFrame(lst_results)
half_df = full_df.iloc[:,:-1]
# Вытащим список с url картинок
lst_img = []
for book in lst_results:
    lst_img.append(book['img_url'])
    

writer = pd.ExcelWriter('pandas.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
half_df.to_excel(writer, sheet_name='Лист', index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Лист']
# Ширина столбцов
writer.sheets['Лист'].set_column(0, 0, 47)
writer.sheets['Лист'].set_column(1, 1, 50)
writer.sheets['Лист'].set_column(2, 2, 6)
writer.sheets['Лист'].set_column(3, 3, 16)

# Запишем картинки
size = (120, 150)
for i, url in enumerate(lst_img):
    row = i + 1
    #print("row=", row)
    # высота строки
    worksheet.set_row(row, 125)
    try:
        image_in_bufer = BytesIO(urlopen(url).read())
        bufer_png = pImage.open(image_in_bufer).convert("RGB")
        bufer_png = bufer_png.resize(size)
        img_stream = BytesIO()
        bufer_png.save(img_stream, format='PNG')
        worksheet.insert_image(f"D{row+1}", img_stream)
    except Exception as e:
        print("___________Исключение_______________")
        print(e)
        print("______________________________________")


# Close the Pandas Excel writer and output the Excel file.
writer.close()
print("Записали датафрейм в таблицу")
