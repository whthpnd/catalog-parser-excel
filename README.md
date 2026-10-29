# 📊 Парсер каталога → Excel

Собирает весь каталог интернет-магазина и выгружает его в Excel: проходит по всем страницам, открывает каждую карточку товара и забирает нужные поля.

Демо на [books.toscrape.com](https://books.toscrape.com) — учебный сайт, сделанный специально для тренировки парсинга.

**Результат:** 1000 товаров из 50 категорий без дублей и пропусков за ~9 минут → [`books_full.xlsx`](books_full.xlsx)

| Название | Цена, £ | В наличии, шт | Рейтинг | Категория |
|---|---|---|---|---|
| A Light in the Attic | 51.77 | 22 | 3 | Poetry |
| Tipping the Velvet | 53.74 | 20 | 1 | Historical Fiction |
| Soumission | 50.10 | 20 | 1 | Fiction |
| Sharp Objects | 47.82 | 20 | 4 | Mystery |
| Sapiens: A Brief History of Humankind | 54.23 | 20 | 5 | History |

## Что умеет
- Обход пагинации до последней страницы
- Заход в карточку каждого товара: название, цена, наличие, рейтинг, категория, ссылка
- Повторные попытки при сетевых ошибках, пауза между запросами, чтобы не нагружать сайт
- Excel с оформленной шапкой, автофильтрами, закреплённой первой строкой и кликабельными ссылками
- Битая карточка не роняет весь сбор: она пропускается и попадает в лог

## Запуск
```bash
python -m venv .venv
.venv\Scripts\activate          # Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt

python parser.py                 # весь каталог → books.xlsx
python parser.py --pages 2       # только первые 2 страницы
python parser.py --out my.xlsx   # другой файл
```

## Стек
Python 3.12 · requests · BeautifulSoup · openpyxl
