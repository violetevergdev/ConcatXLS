import threading
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showerror, showinfo
import os
import pandas as pd
import sqlite3

def main():
    def start():
        in_path = askdirectory(title='Выберите директориюс файлами XLS')
        if in_path:
            with sqlite3.connect('tmp.db') as conn:
                c = conn.cursor()
                col_names = ['Файл','Время', 'Сообщение', 'ФИО', 'СНИЛС', 'Процесс', 'Контекст', 'Уровень', 'Системное сообщение', 'Исключение', 'Куда уехал']
                sql_col = ', '.join([f"'{col}'" for col in col_names])
                c.execute(f"""CREATE TABLE IF NOT EXISTS xlsx_base ({sql_col})""")

                main_title['text'] += ' - в процессе'

                for el in os.listdir(in_path):
                    if el.endswith('.xlsx') or el.endswith('.xls'):
                        try:
                            file_path = os.path.join(in_path, el)

                            df = pd.read_excel(file_path, na_filter=False)
                            data = []
                            for _, row in df.iterrows():
                                if row[0] == '':
                                    continue

                                if "Переезд на новое место жительства в пределах субъекта РФ" in str(row[7]):
                                    try:
                                        ra = row[9]
                                    except Exception:
                                        ra = ''
                                    data.append((el, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], ra))

                            c.executemany(
                                f'INSERT INTO xlsx_base VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', data)

                        except Exception as e:
                            print(e)
                            showerror('Ошибка', str(el) + str(e))
                            return
                conn.commit()

                query = c.execute(f"SELECT * FROM xlsx_base")
                results = pd.DataFrame(query, columns=[col[0] for col in c.description])
                out = os.path.join(in_path, 'Обработанный список.xlsx')
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    results.to_excel(writer, index=False, sheet_name='Sheet')

                c.execute(f"DROP TABLE IF EXISTS xlsx_base")
                main_title['text'] = 'ConcatXLS'
                showinfo('Готово', 'Обработка завершена!')
        else:
            showerror('Ошибка', 'Вы должны выбрать директорию')

    def threaded_start():
        threading.Thread(target=start).start()

    root = tk.Tk()
    root.geometry('250x100')
    root.title('ConcatXLS')
    root.resizable(False, False)
    root.attributes('-topmost', True)
    root['bg'] = '#FF9966'

    main_title = tk.Label(root, text='ConcatXLS', bg='#FF9966', fg='#333', font=('Helvetica', 16))
    main_title.place(relx=0.01, rely=0.01)

    btn = tk.Button(root, text='Объединить', font=('Helvetica', 16), command=threaded_start)
    btn.place(relx=0.245, rely=0.4)


    def exit():
        if os.path.isfile('tmp.db'):
            os.remove('tmp.db')
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", exit)
    root.mainloop()

if __name__ == '__main__':
    main()