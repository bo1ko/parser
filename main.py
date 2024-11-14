from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import openpyxl  # Для роботи з Excel файлами
import json  # Для роботи з JSON-даними
import re  # Для роботи з регулярними виразами

# Налаштування Chrome WebDriver
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Запуск в режимі headless (без вікна браузера)
# chrome_options.add_argument("--disable-gpu")  # Вимкнення GPU (для кращої стабільності в headless режимі)


driver = webdriver.Chrome(options=chrome_options)


# Функція для отримання значень із файлу Excel
def get_queries_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # Беремо активний лист
    data = []
    for row in sheet.iter_rows(
        min_row=2, max_col=2, values_only=True
    ):  # Стартуємо з другого рядка
        data.append(list(row))  # Додаємо пару (UPC, Назва)
    return data, workbook, sheet


def clean_string(s):
    return re.sub(r"\s+|/", "", s.lower())


try:
    driver.get("https://exist.ua/api/v1/fulltext/search-v2/?query=53219&short=true")

    data, workbook, sheet = get_queries_from_excel("data.xlsx")  # Шлях до вашого файлу

    for i, (query, name) in enumerate(data):
        # Формуємо URL з кожним значенням з Excel
        url = f"https://exist.ua/api/v1/fulltext/search-v2/?query={query}&short=true"
        driver.get(url)  # Відкриваємо сторінку з новим запитом

        # Отримуємо відповідь сторінки
        pre_tag = driver.find_element(By.TAG_NAME, "pre")
        pre_text = pre_tag.text
        response = json.loads(pre_text)  # Парсимо JSON відповідь

        # Перевіряємо наявність продуктів у відповіді
        if "products" in response["result"]:
            for product in response["result"]["products"]:
                # Очищаємо назву з Excel і slug
                cleaned_name = clean_string(name)

                print(cleaned_name, product["slug"])
                print(cleaned_name in product["slug"])

                # Перевіряємо, чи є слово з назви в slug
                if cleaned_name in product["slug"]:
                    # Якщо є, записуємо description в 3-ю колонку
                    with open("data.txt", "a+") as file:
                        file.write(product["description"] + "\n")

                    print(
                        f"Знайдено відповідність для UPC {query}: {product['slug']} - {product['description']}"
                    )

        # Зберігаємо зміни після кожного запиту

finally:
    # Закриваємо браузер
    driver.quit()
