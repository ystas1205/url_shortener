import os
from flask import Flask, render_template, request
import openpyxl
import pandas as pd
import requests
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
# TOKEN="vk1.a.qLJEG_uOoL5G5ikN6hK6rc3vkIiVyWpMH8h7tbDHqlmbLHFywu2BqSyjIzcWzt3SPGnyxj6fMttTcagCEacGIOQ2C6ogzNNaVsQADgtV9FENxB6mc0lFnkBzqruwst2olc8mzKr4IGCg1Az3qqUCfKW7kCBvL7IUIHEvsDLEIaRiHuNgR8BrpbqAfZfFNcxL"
api_url = "https://api.vk.com/method/utils.getShortLink"
TOKEN = os.getenv('TOKEN')


def upload_file(file):
    """Загрузка xlsx файла"""
    try:
        file = request.files[file]
        df = openpyxl.load_workbook(file)
        sheet = df.active
        return sheet
    except Exception as e:
        print(f"Ошибка загрузки файла: {e}")
        return None


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/button_clicked', methods=['POST'])
def button_clicked():
    """Формирование и загрузка xlsx файл, в первом столбце которого остаются
     ссылки из входного файла,а во втором короткая ссылка, полученная через VK API"""
    sheet = upload_file('file')
    if sheet is None:
        return "Ошибка загрузки файла"

    original_urls = []
    shortened_urls = []
    data = {}
    for row in sheet.iter_rows(values_only=True):
        params = {"v": "5.131", "url": row[0], "access_token": TOKEN}
        response = requests.get(api_url, params=params)
        original_urls.append(response.json()['response']['url'])
        shortened_urls.append(response.json()['response']['short_url'])
    data.update({'url': original_urls, 'short_url': shortened_urls})
    df = pd.DataFrame(data)
    # Записать в файл XLSX
    df.to_excel('short_links.xlsx', index=False)
    return "Файл загружен"


if __name__ == '__main__':
    app.run(debug=True)
