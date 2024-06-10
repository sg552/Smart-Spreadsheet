from openai import OpenAI
from table_helper import TableHelper
import time
from pathlib import Path

client = OpenAI(
    api_key = 'sk-33KKyUO0tKNxch5RHxa5ELH50BHduZFmgyns6nrR8Fii3l9X',
    base_url = "https://api.moonshot.cn/v1",
)

history = [
    #{"role": "system", "content": "你好kimi, 下面我传给你多个csv 文件，每个csv 文件都代表了一个表格。接下来会让你分析"}
    {"role": "system", "content": "Hello kimi, I will send you some CSV content , each CSV represents a table, and you do some analysis work for me."}
]

# this method from ai API
def chat(query, history):
    history.append({
        "role": "user",
        "content": query
    })

    #print(f"== history: {history}")

    completion = client.chat.completions.create(
        model="moonshot-v1-8k",
        messages=history,
        temperature=0.2,
    )
    result = completion.choices[0].message.content
    history.append({
        "role": "assistant",
        "content": result
    })
    return result

def please_input_question():
    print("== Question: ")
    return input()

if __name__ == "__main__":
    print("=== hello, please give me the excel file ( e.g. /workspace/Smart-Spreadsheet/tests/example_2.xlsx ) , then I will analysis it. input 'bye' to exit ")
    user_input = input()

    table_helper = TableHelper()
    sheet = table_helper.open_file(user_input)

    print(f"=== ok, parsing excel file {user_input} ")
    table_helper.get_tables(sheet)
    print(f"=== parse done, got {len(table_helper.tables)} tables")

    print(f"=== converting to csv and upload to AI... ")
    csv_file_names = []
    for table in table_helper.tables:
        content = table_helper.get_content_from_sheet(sheet, table)
        csv_file_name = f"./csv_files/{table.name}-{time.time()}.csv"
        table_helper.save_as_csv(csv_file_name, content)
        csv_file_names.append(csv_file_name)

    #csv_file_names = ['./csv_files/BALANCE SHEET-1717975552.0232182.csv', './csv_files/INCOME STATEMENT-1717975552.0370402.csv']

    for csv_file in csv_file_names:
        file_object = client.files.create(file=Path(csv_file), purpose="file-extract")
        file_content = client.files.content(file_id=file_object.id).text
        history.append({
            'role': 'system',
            'content': file_content
            })

    print(f"== uploading to AI... ")

    print(chat("hello ~, can you analysis these csv files?", history))
    while(True):
        question = please_input_question()
        if question == 'bye':
            break

        print(f"== Answer: {chat(question,history)}")



