import os
import json, xlwt, csv
from collections import defaultdict


def export_excel(chatbot_id, chat_objects):
    # Setting up the xls file
    file_name = './data_xls/chatbot_id_' + chatbot_id + '_data.xls'
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Page 1')
    worksheet.write(0, 0, 'chat_id')
    worksheet.write(0, 1, 'role')
    worksheet.write(0, 2, 'message')

    col_width_0 = col_width_1 = 0

    try:
        # Iterating through the chat objects
        row = 1

        for chat_item in chat_objects:
            for history_item in chat_item['history']:

                worksheet.write(row, 0, chat_item['chat_id'])
                if len(chat_item['chat_id']) > col_width_0:
                    col_width_0 = len(chat_item['chat_id'])

                worksheet.write(row, 1, history_item['role'].replace('system', 'airose'))
                if len(history_item['role']) > col_width_1:
                    col_width_1 = len(history_item['role'])

                worksheet.write(row, 2, history_item['content'])

                row += 2

            if chat_item != chat_objects[-1]:
                row += 1
                worksheet.write(row, 0, '***************NEW-CHAT***************')
                row += 3

    except Exception as e:
        print('An error occurred when exporting to xls :', e)
    
    worksheet.col(0).width = 256 * col_width_0
    worksheet.col(1).width = 256 * col_width_1
    worksheet.col(2).width = 65000

    workbook.save(file_name)


def export_csv(chatbot_id, chat_objects):
    file_name = './data_csv/chatbot_id_' + chatbot_id + '_data.csv'
    mode = 'w'

    if not os.path.exists(file_name):
        mode = 'a'

    with open(file_name, mode, newline='') as file:
        writer = csv.writer(file, delimiter='|')
        writer.writerow(['chat_id', 'role', 'message'])

        try:
            # Iterating through the chat objects
            for chat_item in chat_objects:
                for history_item in chat_item['history']:
                    writer.writerow([chat_item['chat_id'], history_item['role'].replace('system', 'airose'), history_item['content']])
                    writer.writerow('')

                if chat_item != chat_objects[-1]:
                    writer.writerow('')
                    writer.writerow(['***************NEW-CHAT***************'])
                    writer.writerow('')
                    writer.writerow('')

        except Exception as e:
            print('An error occurred when exporting to csv :', e)


json_file = 'chatbot_sample.json'
file = open(json_file)
data = json.load(file)

# Dictionary (key) => chatbot_id  -  (value) => chat_item(id, chat_id, chatbot_id, history)
chatbot_id_groups = defaultdict(list)

# Grouping the chat_items by chatbot_ids
for chat_obj in data:
    chatbot_id_groups[chat_obj['chatbot_id']].append(chat_obj)

# Exporting the data to xsl and csv
for chatbot_id in chatbot_id_groups:
    export_excel(chatbot_id, chatbot_id_groups[chatbot_id])
    export_csv(chatbot_id, chatbot_id_groups[chatbot_id])
    