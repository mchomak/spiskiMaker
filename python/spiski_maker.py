import tkinter as tk
from tkinter import *
from spiski import Ecxel
from tkinter import messagebox
import tkinter.filedialog as fd
import os

exel_name=''
directory=''
window=Tk()
window_g='400x500'
start=0
choose_file=0
redact_lists=''
all_lists=0
errors=0

# обновление текстов
def text_ubdate(all_lists,errors):
    global directory,exel_name
    directory,exel_name=directory_obrbotka(start_name)
    derectiva['text'] = 'Директория: ' + '\n'+ str(directory)

    start_name_title['text']='Название файла: ' + '\n' + str(exel_name) + '.xlsx'

    save_name=exel_name+' отредактированный'+ '.xlsx'
    end_name_title['text']='Название сохраненного файла: ' + '\n' + str(save_name)

    lists_title['text']='Количество отредактированных листов: ' + str(all_lists)

    error_title['text']='Ошибки: ' + str(errors)

# кнопка выбора
def choose_derectore():
    global start_name,choose_file,startnametitle,title1
    if choose_file==0:
        start_name=fd.askopenfilename(title='открыть файл',initialdir='/')
        start_name=start_name[:-5]
        choose_file=1
        text_ubdate(all_lists,errors)
        title1['text'] = 'Файл выбран'
        file_name.delete(0, tk.END)
        file_name.insert(0, exel_name)

# начало обработки файла
def Start():
    global start,error_title,start_name,redact_lists,choose_file,save_name,otchet_name
    if choose_file==0:
        start_name = file_name.get()

    redact_lists=lists.get()

    if start_name!=' ' and redact_lists!=' ':
        try:
            directory, exel_name = directory_obrbotka(start_name)
            save_name, all_lists, errors,otchet_name=Ecxel.main('list',start_name,redact_lists,directory,exel_name)
            start = 1
        except FileNotFoundError:
            messagebox.showerror(title='error',message='такого файла нету в дериктории')
            start = 0

        if start==1:
            start=0
            choose_file = 0
            title1['text'] = 'Введите название файла:'
            text_ubdate(all_lists,errors)
        else:
            pass

def directory_obrbotka(start_name):
    directory=''
    start_name=start_name.split('/')
    name=start_name[-1]
    start_name.pop(-1)
    for i in start_name:
        directory=directory+i+'/'
    return directory,name

def open_otchet():
    os.startfile(otchet_name)
    print(str(directory)+'otchet_of_work.txt')

def open_file():
    print(save_name)
    os.startfile(save_name)

window.title('spiski_maker')
window.geometry(window_g)
window.minsize(400,550)

# астраиваем все виджеты
canvas=Canvas(window,height=600,width=300)
canvas.pack()

frame=Frame(window)
frame.place(relwidth=1,relheight=1)

title1=Label(frame,text='Введите название файла:',font=40,bg='white',anchor='w',relief=tk.RAISED,width=25)

choose_button=Button(frame,text='обзор:',bg='white',command=choose_derectore,relief=tk.RAISED,font=40)

file_name=Entry(frame,bg='white',width=25,font=40,text='Введите количество листов:')

start_name=file_name.get()

title2=Label(frame,text='Введите количество листов:',font=40,bg='white',anchor='w',relief=tk.RAISED,width=25)

lists=Entry(frame,bg='white',width=25,font=40)

start_button=Button(frame,text='запустить',bg='white',command=Start,relief=tk.RAISED,font=40)


start_name_title = Label(frame, text='Название файла: ' + '\n' + str(exel_name) + '.xlsx',
                         font=40, bg='white', anchor='w', width=60)

save_name=exel_name+' отредактированный'+ '.xlsx'
end_name_title = Label(frame, text='Название сохраненного файла: ' + '\n' + str(save_name),
                       font=40, bg='white', anchor='w', width=60)

lists_title = Label(frame, text='Количество отредактированных листов:' + str(all_lists),
                    font=40, bg='white', anchor='w', width=60)

error_title = Label(frame, text='Ошибки: ' + str(errors),
                    font=40, bg='white', anchor='w', width=60)

info_title=Label(frame, text='Информация о работе:',
                    font=40, bg='white', anchor='w', width=60)

derectiva=Label(frame, text='Директория:'+'\n'+str(directory),
                    font=40, bg='white', anchor='w', width=60)

otchet_button=Button(frame,text='Полный отчет о работе',
                     bg='white',command=open_otchet,relief=tk.RAISED,font=40)

open_file_button=Button(frame,text='Открыть файл',
                     bg='white',command=open_file,relief=tk.RAISED,font=40)


# отрисовываем все виджеты
title1.grid(stick='w',row=0,column=0,pady=5)
file_name.grid(stick='w',row=1,column=0)
title2.grid(stick='w',row=2,column=0,pady=5)
lists.grid(stick='w',row=3,column=0)
start_button.grid(stick='w',row=4,column=0,pady=10)
info_title.grid(stick='w',row=5,column=0)
derectiva.grid(stick='w',row=6,column=0)
start_name_title.grid(row=7,column=0)
end_name_title.grid(row=8,column=0)
lists_title.grid(row=9,column=0)
error_title.grid(row=10,column=0)
choose_button.grid(row=0,column=0)
otchet_button.grid(stick='w',row=11,column=0)
open_file_button.grid(stick='w',row=12,column=0)
window.mainloop()
