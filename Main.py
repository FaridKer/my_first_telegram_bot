from aiogram import Bot, Dispatcher, executor, types                                                                    #установил модуль для работы ТГ бота
import requests                                                                                                         #установил модуль для работы с сайтами
from bs4 import BeautifulSoup                                                                                           #установил модуль для анализа страницы
from openpyxl import load_workbook                                                                                      #установил модуль для переноса данных в эксель

#Раздел, связанный с настройкой бота
API_TOKEN = '5908870637:AAGqVbjT9V6SZb3zWlC2NGiOA4E9pkdB-BE'                                                            #Задаю переменную для своего токена ТГ
bot = Bot(token=API_TOKEN)                                                                                              #Передаем боту токен, чтобы была инициализация
dp = Dispatcher(bot)                                                                                                    #Отслеживает обновления?
@dp.message_handler(commands=['start'])                                                                                 #Указываем, на какую команду пользователя запускать функцию. Я не знаю, что значит "@"
async def send_welcome(message: types.Message):                                                                         #исполняется функция с аргументом, который написал пользователь, 'start'
   await message.reply("Привет!\nЯ первый бот Фарида!\nОтправь мне любое название вакансии, а я тебе обязательно сброшу excel-таблицу с доступными вакансиями на hh.ru") #в статье написано, что обязательно надо писать await, так как программа работает асинхронно, но у меня нет понимания, что это значит
@dp.message_handler(content_types=["text"])                                                                             #вызывается событие в ответ на любой текст пользователя
async def send_file(message: types.Document):                                                                           #аргументом является сообщение пользователя, а ответом файл?
    # ЧАСТЬ ПЕРВАЯ - ФОРМУЛИРОВАНИЕ НУЖНОЙ АДРЕСНОЙ СТРОКИ И ЗАПИСЬ В ФАЙЛ
    URL_dict = {'а': '%D0%B0', 'б': '%D0%B1', 'в': '%D0%B2', 'г': '%D0%B3',
                'д': '%D0%B4', 'е': '%D0%B5', 'ё': '%D0%B5', 'ж': '%D0%B6',
                'з': '%D0%B7', 'и': '%D0%B8', 'й': '%D0%B9', 'к': '%D0%BA',
                'л': '%D0%BB', 'м': '%D0%BC', 'н': '%D0%BD', 'о': '%D0%BE',
                'п': '%D0%BF', 'р': '%D1%80', 'с': '%D1%81', 'т': '%D1%82',
                'у': '%D1%83', 'ф': '%D1%84', 'х': '%D1%85', 'ц': '%D1%86',
                'ч': '%D1%87', 'ш': '%D1%88', 'щ': '%D1%89', 'ъ': '%D1%8A',
                'ы': '%D1%8B', 'ь': '%D1%8C', 'э': '%D1%8D', 'ю': '%D1%8E',
                'я': '%D1%8F',
                ' ': '+',
                'А': '%D0%90', 'Б': '%D0%91', 'В': '%D0%92', 'Г': '%D0%93',
                'Д': '%D0%94', 'Е': '%D0%95', 'Ё': '%D0%95', 'Ж': '%D0%96',
                'З': '%D0%97', 'И': '%D0%98', 'Й': '%D0%99', 'К': '%D0%9A',
                'Л': '%D0%9B', 'М': '%D0%9C', 'Н': '%D0%9D', 'О': '%D0%9E',
                'П': '%D0%9F', 'Р': '%D0%A0', 'С': '%D0%A1', 'Т': '%D0%A2',
                'У': '%D0%A3', 'Ф': '%D0%A4', 'Х': '%D0%A5', 'Ц': '%D0%A6',
                'Ч': '%D0%A7', 'Ш': '%D0%A8', 'Щ': '%D0%A9', 'Ъ': '%D0%AA',
                'Ы': '%D0%AB', 'Ь': '%D0%AC', 'Э': '%D0%AD', 'Ю': '%D0%AE',
                'Я': '%D0%AF',
                '1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6',
                '7': '7', '8': '8', '9': '9', '0': '0',
                'a': 'a', 'b': 'b', 'c': 'c', 'd': 'd', 'e': 'e', 'f': 'f',
                'g': 'g', 'h': 'h', 'i': 'i', 'j': 'j', 'k': 'k', 'l': 'l',
                'm': 'm', 'n': 'n', 'o': 'o', 'p': 'p', 'q': 'q', 'r': 'r',
                's': 's', 't': 't', 'u': 'u', 'v': 'v', 'w': 'w', 'x': 'x',
                'y': 'y', 'z': 'z',
                'A': 'A', 'B': 'B', 'C': 'C', 'D': 'D', 'E': 'E', 'F': 'F',
                'G': 'G', 'H': 'H', 'I': 'I', 'J': 'J', 'K': 'K', 'L': 'L',
                'M': 'M', 'N': 'N', 'O': 'O', 'P': 'P', 'Q': 'Q', 'R': 'R',
                'S': 'S', 'T': 'T', 'U': 'U', 'V': 'V', 'W': 'W', 'X': 'X',
                'Y': 'Y', 'Z': 'Z',
                '!': '%21', '#': '%23', '$': '%24', '%': '%25', '&': '%26',
                "'": "%27", '(': '%28', ')': '%29', '*': '%2A', '+': '%2B',
                ',': '%2C', '/': '%2F', ':': '%3A', ';': '%3B', '=': '%3D',
                '?': '%3F', '@': '%40', '[': '%5B',
                ']': '%5D'}                                                                                             # гуглил готовые словари или модули, не нашел, решил вручную вбить
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                             'Chrome/108.0.0.0 Safari/537.36'}                                                          # без headers был ответ 404. Нагуглил решение,
                                                                                                                        # что нужен headers, но я толком не понимаю что это и зачем нужно, что-то типо первичной отправки инфы на сервер ХХ,
                                                                                                                        # кто отправил запрос, чтобы получить одобрение
    vacancy = ''.join([URL_dict[i] for i in
                       message.text])                                                                                   # разбил вводимое слово на список из символов, чтобы использовать их, как ключи словаря
    HH_URL = 'https://hh.ru/search/vacancy?text=' + vacancy                                                             # на основе вводимого слова и словаря сделал конструктор текста требуемой адресной строки
    my_request = requests.get(HH_URL, headers=headers)                                                                  # как я понял из гугла, таким образом мы получаем ответ от сайта
    with open('Vacancy_list.html', 'wb') as output_file:
        output_file.write(my_request.text.encode(
            'utf-8'))                                                                                                   # как я понял, таким образом мы сохраняем страницу в файл и в кодировке utf-8

    # ЧАСТЬ ВТОРАЯ - РАБОТА С ФАЙЛОМ И СОСТАВЛЕНИЕ СПИСКА ВАКАНСИЙ
    with open('Vacancy_list.html', 'r',
              encoding='utf-8') as vacancy_page:                                                                        # открываю файл в режиме чтения, задаю имя переменной
        soup = BeautifulSoup(vacancy_page,
                             features="lxml")                                                                           # как я понял, это модуль для анализа содержимого веб-страниц
        vacancy_tab = soup.find('div', {
            'class': 'vacancy-serp-content'})                                                                           # с помощью метода поиска нашел по тегу столбец на сайте, который отвечает за выпадающие вакансии
        vacancies = vacancy_tab.find_all('div', {
            'class': 'serp-item'})                                                                                      # объединил одинаковые элементы в одну группу, чтобы воспользоваться циклом for
        vacancy_list = []                                                                                               # создал пустой список для дальнейшего наполнения
        for vacancy in vacancies:
            vacancy_info = []
            vacancy_link = vacancy.find(
                'a')                                                                                                    # изучив веб-страницу, я заметил, что ссылка всегда находится в строке, которая начинается на 'a'
            if vacancy_link is not None:                                                                                # здесь и далее, чтобы не возникала ошибка того, что я пытаюсь манипулировать с объектом типо None, провожу проверку
                vacancy_link = vacancy.find('a').get(
                    'href')                                                                                             # если проверка успешна - присваиваю переменной требуемую строку
            else:
                vacancy_link = '-'                                                                                      # если проверка провалена - ставлю прочерк, так как строка пустая, но для будущего списка нужно соблюдать равное число элементов списка
            vacancy_name = vacancy.find('a')                                                                            # поиск названия вакансии
            if vacancy_name is not None:
                vacancy_name = vacancy.find('a').text
            else:
                vacancy_name = '-'
            vacancy_salary = vacancy.find('span', {'class': 'bloko-header-section-3'})                                  # поиск зарплаты по вакансии
            if vacancy_salary is not None:
                vacancy_salary = vacancy.find('span', {'class': 'bloko-header-section-3'}).text
                s1 = vacancy_salary.replace('\u202f',
                                            ' ')                                                                        # здесь и далее - замечал артефакты в строках, связанных с отступами. Решил их удалить через строковый метод replace()
                s2 = s1.replace('\xa0', ' ')
                vacancy_salary = s2
            else:
                vacancy_salary = '-'
            vacancy_company = vacancy.find('a', {
                'class': 'bloko-link bloko-link_kind-tertiary'})                                                        # поиск названия компании работодателя
            if vacancy_company is not None:
                vacancy_company = vacancy.find('a', {'class': 'bloko-link bloko-link_kind-tertiary'}).text
                s1 = vacancy_company.replace('\xa0', ' ')
                vacancy_company = s1
            else:
                vacancy_company = '-'
            vacancy_city = vacancy.find('div', {
                'data-qa': 'vacancy-serp__vacancy-address'})                                                            # поиск указанного города по вакансии
            if vacancy_city is not None:
                vacancy_city = vacancy.find('div', {'data-qa': 'vacancy-serp__vacancy-address'}).text
                s1 = vacancy_city.replace('\xa0', ' ')
                vacancy_city = s1
            else:
                vacancy_city = '-'
            vacancy_description = vacancy.find('div', {
                'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'})                                             # поиск строки с описанием вакансии
            if vacancy_description is not None:
                vacancy_description = vacancy.find('div',
                                                   {'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text
            else:
                vacancy_description = '-'
            vacancy_requirement = vacancy.find('div', {
                'data-qa': 'vacancy-serp__vacancy_snippet_requirement'})                                                # поиск строки с описанием требований по вакансии
            if vacancy_requirement is not None:
                vacancy_requirement = vacancy.find('div', {'data-qa': 'vacancy-serp__vacancy_snippet_requirement'}).text
            else:
                vacancy_requirement = '-'
            vacancy_info.append(vacancy_link)
            vacancy_info.append(vacancy_name)
            vacancy_info.append(vacancy_salary)
            vacancy_info.append(vacancy_company)
            vacancy_info.append(vacancy_city)
            vacancy_info.append(vacancy_description)
            vacancy_info.append(vacancy_requirement)                                                                    # добавляю полученные данные в список vacancy_info на каждом цикле
            vacancy_list.append(
                vacancy_info)                                                                                           # каждый полученный список vacancy_info Добавляю в итоговый список vacancy_list

        # ЧАСТЬ ТРЕТЬЯ - ЗАНЕСЕНИЕ ДАННЫХ ПО ВАКАНСИЯМ В EXCEL-ТАБЛИЦУ
        fn = 'vacancy_excel_list.xlsx'                                                                                  # подсмотрел на ютубе модуль openpyxl для работы с эксель таблицами
        wb = load_workbook(fn)
        ws = wb['data']
        ws.delete_cols(1, 100)
        ws.delete_rows(1, 100)
        ws['A1'] = '№'
        ws['B1'] = 'Ссылка на вакансию'
        ws['C1'] = 'Вакансия'
        ws['D1'] = 'Зарплата'
        ws['E1'] = 'Работодатель'
        ws['F1'] = 'Город'
        ws['G1'] = 'Описание Вакансии'
        ws['H1'] = 'Требование по Вакансии'
        i = 2
        for vacancy in vacancy_list:
            ws['A' + str(i)] = i - 1
            ws['B' + str(i)] = vacancy[0]
            ws['C' + str(i)] = vacancy[1]
            ws['D' + str(i)] = vacancy[2]
            ws['E' + str(i)] = vacancy[3]
            ws['F' + str(i)] = vacancy[4]
            ws['G' + str(i)] = vacancy[5]
            ws['H' + str(i)] = vacancy[6]
            i += 1
        wb.save(fn)
        wb.close()                                                                                                      # по окончанию работы программы получаем готовую Excel-таблицу
    await message.reply_document(open(r'C:\Users\doggf\OneDrive\Документы\GitHub\MyFirstRepository\Мой первый тг-бот\vacancy_excel_list.xlsx', 'rb')) # файл отправляется пользователю в ТГ

if __name__ == '__main__':                                                                                              # как я понял это нужно, чтобы бот запустил работу
   executor.start_polling(dp, skip_updates=True)
