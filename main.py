import time
import telebot
import csv
from telebot import types
import datetime
import aspose.words as aw
import math as m
from token_bot import bot_token
bot=telebot.TeleBot(bot_token)

def csv_to_doxc(file_rows,open_file_name,vl_name,vl_voltage,vl_uchastok,vl_podstation,name_town,type_looking,date,name_worker1,name_worker2,name_master,name_ingeener,message):
    '''Функция принимает на вход данные для заполнения шапки и подвала отчёта'''
    open_file_name=open_file_name[:-4]
    def make_page(lst_rows,name_file,open_file_name,vl_name,vl_voltage,vl_uchastok,vl_podstation,name_town,type_looking,date,name_worker1,name_worker2,name_master,name_ingeener,message):
        #Создаётся новый документ
        doc = aw.Document()


        builder = aw.DocumentBuilder(doc)


        # Создание страницы

        # Шапка страницы
        builder.write(
            "Филиал БЭС                                                                                                      Сельский РЭС\n\n")
        builder.write("                                                   ЛИСТОК ОСМОТРА ВЛ 0,4-10 кВ\n")
        builder.write(f"ВЛ {vl_voltage}, кВ № {vl_name}, участок: {vl_uchastok}\n")
        builder.write(f"От подстанции {vl_podstation}, н.п. {name_town}\n")
        builder.write(f"Вид осмотра: {type_looking}, Дата осмотра:  {date[:11]}\n\n")

        # Таблица страницы
        table = builder.start_table()
        # Insert cell.
        builder.insert_cell()
        # Table wide formatting must be applied after at least one row is present in the table.
        table.left_indent = 1.0
        # Set height and define the height rule for the header row.
        builder.row_format.height = 10.0
        builder.row_format.height_rule = aw.HeightRule.AT_LEAST
        # Set alignment and font settings.
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.size = 8.0
        builder.font.name = "Arial"
        builder.font.bold = True

        builder.cell_format.width = 35.0
        builder.write("Номера опор,\nпролетов")

        # We don't need to specify this cell's width because it's inherited from the previous cell.
        builder.font.size = 10
        builder.insert_cell()
        builder.cell_format.width = 250.0
        builder.write("Выявленные дефекты")

        builder.insert_cell()
        builder.cell_format.width = 120.0
        builder.write("Состояние трассы")
        builder.end_row()
        for i in lst_rows:
            builder.cell_format.width = 35.0
            builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER

            # Reset height and define a different height rule for table body.
            builder.row_format.height = 10.0
            builder.row_format.height_rule = aw.HeightRule.AUTO
            builder.insert_cell()

            # Reset font formatting.
            builder.font.size = 10
            builder.font.bold = False

            builder.write(i[0])

            builder.insert_cell()
            builder.cell_format.width = 250.0
            builder.write(i[1])

            builder.insert_cell()
            builder.cell_format.width = 120.0
            builder.write(i[2])
            builder.end_row()


        # End table.
        builder.end_table()
        # Концовка страницы
        builder.write("Осмотр произвели:\n")
        builder.write(f"ФИО: {name_worker1}    Подпись: _____________\n")
        builder.write(f"ФИО: {name_worker2}    Подпись: _____________\n")
        builder.write(f"Листок осмотра принял:  {date[:11]} , {name_master} , Подпись: ___________\n")
        builder.write(f"Результаты осмотра проанализированы:   {date[:11]}\n")
        builder.write(f"Главный инженер РЭС: {name_ingeener}, Подпись: ________________\n")

        # Save the document.
        doc_name=(f'{open_file_name}, часть {name_file}.docx')
        doc.save(f'{open_file_name}, часть {name_file}.docx')
        f = open(doc_name.encode(encoding='UTF-8',errors='strict'), "rb")
        bot.send_document(message.chat.id, f)
    lst_rows=[]
    count=0
    name_file=1
    for i in file_rows:
        if count+m.ceil(len(i[1])/55)+1<39:

            fin_wrt=''
            for j in i[1]:
                if len(fin_wrt)%55==0 and len(fin_wrt)>=55:
                    fin_wrt+=(j)
                else:
                    fin_wrt+=j
            lst_rows.append([i[0],fin_wrt,i[2]])
            count += m.ceil(len(i[1])/55)+1
        else:
            make_page(lst_rows,name_file,open_file_name,vl_name,vl_voltage,vl_uchastok,vl_podstation,name_town,type_looking,date,name_worker1,name_worker2,name_master,name_ingeener,message)
            name_file+=1
            fin_wrt = ''
            for j in i[1]:
                if len(fin_wrt) % 55 == 0 and len(fin_wrt) >= 55:
                    fin_wrt += (j)
                else:
                    fin_wrt += j
            lst_rows=[[i[0],fin_wrt,i[2]]]
            count=m.ceil(len(i[1])/55)+1
    if len(lst_rows)>0:
        make_page(lst_rows,name_file, open_file_name,vl_name,vl_voltage,vl_uchastok,vl_podstation,name_town,type_looking,date,name_worker1,name_worker2,name_master,name_ingeener,message)

user={'642828322': {'name': 'Кемза А.Г.',
                    'write':[],
                    'step': 0,
                    'looking_day': 0,
                    'vl_name': '000',
                    'open_file_name': [],
                    'vl_voltage': 'none',
                    'vl_uchastok': 'none',
                    'vl_podstation': 'none',
                    'name_town': 'none',
                    'type_looking': 'none',
                    'name_worker1': 'none',
                    'name_worker2': 'none',
                    'name_master': 'none',
                    'name_ingeener': 'none',
                    'doc_name': 'none'}
      }





def start1(message):
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('Начать осмотр')
        markup.row(btn1)
        bot.send_message(message.from_user.id, 'Давай осмотрим ВЛ ;)', reply_markup=markup)

@bot.message_handler(commands=['start', "Главное меню Бота"])
def start(message):
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('Начать осмотр')
        btn3 = types.KeyboardButton('Главное меню бота')
        markup.row(btn1, btn3)
        bot.send_message(message.from_user.id, 'Давай осмотрим ВЛ ;)', reply_markup=markup)
do_look=False
step=0

@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    try:
        if message.text == 'Закончить осмотр':
            print('ok')
            with open(
                    (f'{user[str(message.from_user.id)]["open_file_name"]}').encode(encoding='UTF-8', errors='strict'),
                    encoding='utf-8') as csv_file:
                lst_defection = list(csv.reader(csv_file, delimiter=','))
                csv_to_doxc(lst_defection,
                            user[str(message.from_user.id)]["open_file_name"],
                            user[str(message.from_user.id)]["vl_name"],
                            user[str(message.from_user.id)]["vl_voltage"],
                            user[str(message.from_user.id)]["vl_uchastok"],
                            user[str(message.from_user.id)]["vl_podstation"],
                            user[str(message.from_user.id)]["name_town"],
                            user[str(message.from_user.id)]["type_looking"],
                            user[str(message.from_user.id)]["looking_day"],
                            user[str(message.from_user.id)]["name_worker1"],
                            user[str(message.from_user.id)]["name_worker2"],
                            user[str(message.from_user.id)]["name_master"],
                            user[str(message.from_user.id)]["name_ingeener"],
                            message)
            print('ok func')
            # file_rows,open_file_name,vl_name,vl_voltage,vl_uchastok,vl_podstation,name_town,type_looking,date,name_worker1,name_worker2,name_master,name_ingeener
            # print(f'{user[str(message.from_user.id)]["open_file_name"]}', "rb")
            # f = open(f'{user[str(message.from_user.id)]["open_file_name"]}', "rb")
            # bot.send_document(message.chat.id, f)
            user[str(message.from_user.id)] = {'name': 'Кемза А.Г.',
                                               'write': [],
                                               'step': 0,
                                               'looking_day': 0,
                                               'vl_name': '000',
                                               'open_file_name': [],
                                               'vl_voltage': 'none',
                                               'vl_uchastok': 'none',
                                               'vl_podstation': 'none',
                                               'name_town': 'none',
                                               'type_looking': 'none',
                                               'name_worker1': 'none',
                                               'name_worker2': 'none',
                                               'name_master': 'none',
                                               'name_ingeener': 'none',
                                               'doc_name': 'none'}
            bot.send_message(message.from_user.id, 'Осмотр закончен')
            start1(message)


        elif message.text == 'Начать осмотр' and user[str(message.from_user.id)][
            'step'] == 0:
            user[str(message.from_user.id)]['step'] = 1
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Введи номер ВЛ (Просто цифры)', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1 and message.text.isdigit():
            user[str(message.from_user.id)]['looking_day'] = str(
                datetime.datetime.fromtimestamp(int(message.date)).strftime("%d-%m-%Y %H-%M-%S"))
            user[str(message.from_user.id)]['vl_name'] = message.text
            with open((
                      f'{user[str(message.from_user.id)]["name"]} Осмотр {message.text} {user[str(message.from_user.id)]["looking_day"]}.csv').encode(
                    encoding='UTF-8', errors='strict'), 'w', encoding='utf-8', newline='') as csv_file:
                writer = csv.writer(csv_file)
            user[str(message.from_user.id)]['step'] = 1.1
            user[str(message.from_user.id)][
                'open_file_name'] = f'{user[str(message.from_user.id)]["name"]} Осмотр {message.text} {user[str(message.from_user.id)]["looking_day"]}.csv'
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Введи напряжение линии 0,4 или 10',
                             reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.1:
            user[str(message.from_user.id)]['step'] = 1.2
            user[str(message.from_user.id)]['vl_voltage'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши осматриваемый участок', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.2:
            user[str(message.from_user.id)]['step'] = 1.3
            user[str(message.from_user.id)]['vl_uchastok'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши название подстанции', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.3:
            user[str(message.from_user.id)]['step'] = 1.4
            user[str(message.from_user.id)]['vl_podstation'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши населённый пункт', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.4:
            user[str(message.from_user.id)]['step'] = 1.5
            user[str(message.from_user.id)]['name_town'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши ФИО первого работника осматривающего ВЛ',
                             reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.5:
            user[str(message.from_user.id)]['step'] = 1.6
            user[str(message.from_user.id)]['name_worker1'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши ФИО второго работника осматривающего ВЛ',
                             reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.6:
            user[str(message.from_user.id)]['step'] = 1.7
            user[str(message.from_user.id)]['name_worker2'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши ФИО мастера', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.7:
            user[str(message.from_user.id)]['step'] = 1.8
            user[str(message.from_user.id)]['name_master'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши Вид осмотра', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.8:
            user[str(message.from_user.id)]['step'] = 1.9
            user[str(message.from_user.id)]['type_looking'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши ФИО Главного инженера', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 1.9:
            user[str(message.from_user.id)]['step'] = 2
            user[str(message.from_user.id)]['name_ingeener'] = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши номер опоры или пролёт', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 2:
            user[str(message.from_user.id)]['write'] += [message.text]
            user[str(message.from_user.id)]['step'] = 3
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id, 'Напиши выявленный дефект', reply_markup=markup)

        elif user[str(message.from_user.id)]['step'] == 3:
            user[str(message.from_user.id)]['write'] += [message.text] + [' ']
            user[str(message.from_user.id)]['step'] = 2
            with open(
                    (f'{user[str(message.from_user.id)]["open_file_name"]}').encode(encoding='UTF-8', errors='strict'),
                    'a', encoding='utf-8', newline='') as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(user[str(message.from_user.id)]['write'])
            bot.send_message(message.from_user.id,
                             f'Дефект записан.{",".join(user[str(message.from_user.id)]["write"])}')
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            (user[str(message.from_user.id)]['write']) = []
            btn1 = types.KeyboardButton('Закончить осмотр')
            markup.row(btn1)
            bot.send_message(message.from_user.id,
                             'Продолжаем осмотр. Введи номер опоры или пролёт опор через "-". (Пример:"1-2")',
                             reply_markup=markup)
    except:
        user[str(message.from_user.id)] = {'name': 'Кемза А.Г.',
                                           'write': [],
                                           'step': 0,
                                           'looking_day': 0,
                                           'vl_name': '000',
                                           'open_file_name': [],
                                           'vl_voltage': 'none',
                                           'vl_uchastok': 'none',
                                           'vl_podstation': 'none',
                                           'name_town': 'none',
                                           'type_looking': 'none',
                                           'name_worker1': 'none',
                                           'name_worker2': 'none',
                                           'name_master': 'none',
                                           'name_ingeener': 'none',
                                           'doc_name': 'none'}
        bot.send_message(message.from_user.id, 'Ошибка. Осмотр закончен')
        start1(message)


if __name__ == '__main__':
    while True:
        try:
            bot.polling(none_stop=True)
        except Exception as e:
            print(e)
            time.sleep(15)







