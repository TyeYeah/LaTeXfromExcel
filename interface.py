# !/user/bin/env Python3
# -*- coding:utf-8 -*-
from writeTable import *
from readTable import *
import tkinter as tk
from tkinter import filedialog, dialog, messagebox
import os

file_path = ''


def openfile():
    global file_path
    file_path = filedialog.askopenfilename()
    print(file_path)

def outputCSV():
    global file_path
    file_path = filedialog.askopenfilename()
    print(file_path)
    if file_path == '':
        return ''
    path, suffix = os.path.splitext(file_path)
    filesuffix = suffix.strip().lower()
    if filesuffix.lower() == '.csv':
        try:
            writeCSV(file_path[0:-4] + '.csv', readCSV(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening csv')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.csv')
            print('read csv successfully')
    elif filesuffix.lower() == '.xls':
        try:
            writeCSV(file_path[0:-4] + '.csv', read03xls(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xls')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.csv')
            print('read csv successfully')
    elif filesuffix.lower() == '.xlsx':
        try:
            writeCSV(file_path[0:-5] + '.csv', read07xlsx(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xlsx')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-5] + '.csv')
            print('read csv successfully')
    else:
        messagebox.showerror('Something Wrong',
                             'Unsopported file suffix!\nOnly ".xls", ".xlsx", ".csv" are permitted')

def outputXLSX():
    global file_path
    file_path = filedialog.askopenfilename()
    print(file_path)
    if file_path == '':
        return ''
    path, suffix = os.path.splitext(file_path)
    filesuffix = suffix.strip().lower()
    if filesuffix.lower() == '.csv':
        try:
            write07xlsx(file_path[0:-4] + '.xlsx', readCSV(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening csv')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.xlsx')
            print('read csv successfully')
    elif filesuffix.lower() == '.xls':
        try:
            write07xlsx(file_path[0:-4] + '.xlsx', read03xls(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xls')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.xlsx')
            print('read csv successfully')
    elif filesuffix.lower() == '.xlsx':
        try:
            write07xlsx(file_path[0:-5] + '.xlsx', read07xlsx(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xlsx')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-5] + '.xlsx')
            print('read csv successfully')
    else:
        messagebox.showerror('Something Wrong',
                             'Unsopported file suffix!\nOnly ".xls", ".xlsx", ".csv" are permitted')

def outputXLS():
    global file_path
    file_path = filedialog.askopenfilename()
    print(file_path)
    if file_path == '':
        return ''
    path, suffix = os.path.splitext(file_path)
    filesuffix = suffix.strip().lower()
    if filesuffix.lower() == '.csv':
        try:
            write03xls(file_path[0:-4] + '.xls', readCSV(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening csv')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.xls')
            print('read csv successfully')
    elif filesuffix.lower() == '.xls':
        try:
            write03xls(file_path[0:-4] + '.xls', read03xls(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xls')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.xls')
            print('read csv successfully')
    elif filesuffix.lower() == '.xlsx':
        try:
            write03xls(file_path[0:-5] + '.xls', read07xlsx(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xlsx')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-5] + '.xls')
            print('read csv successfully')
    else:
        messagebox.showerror('Something Wrong',
                             'Unsopported file suffix!\nOnly ".xls", ".xlsx", ".csv" are permitted')

def outputTEX():
    global file_path
    file_path = filedialog.askopenfilename()
    print(file_path)
    if file_path == '':
        return ''
    path, suffix = os.path.splitext(file_path)
    filesuffix = suffix.strip().lower()
    if filesuffix.lower() == '.csv':
        try:
            writeLaTeX(file_path[0:-4] + '.tex', readCSV(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening csv')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.tex')
            print('read csv successfully')
    elif filesuffix.lower() == '.xls':
        try:
            writeLaTeX(file_path[0:-4] + '.tex', read03xls(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xls')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-4] + '.tex')
            print('read csv successfully')
    elif filesuffix.lower() == '.xlsx':
        try:
            writeLaTeX(file_path[0:-5] + '.tex', read07xlsx(file_path))
        except Exception as e:
            messagebox.showerror('Something Wrong', 'Error occurred when opening xlsx')
        else:
            messagebox.showinfo('Info', 'Successfully converted!\nFile saved in \n' + file_path[0:-5] + '.tex')
            print('read csv successfully')
    else:
        messagebox.showerror('Something Wrong',
                             'Unsopported file suffix!\nOnly ".xls", ".xlsx", ".csv" are permitted')


def mainwindow():
    def conversion():
        window.destroy()

        def backmain():
            conversionwin.destroy()
            mainwindow()

        def nodelwin():
            messagebox.showinfo('Tip', 'Press Back to Main to Quit')
            return ''

        global file_path
        file_path = ''
        conversionwin = tk.Tk()
        conversionwin.title('File Format Conversion')
        conversionwin.geometry('500x300')
        # conversionwin.protocol("WM_DELETE_WINDOW", nodelwin)
        banner = 'Source file should have suffix as csv, xls or xlsx'
        txt = tk.Label(conversionwin, text=banner, font=('Arial', 12), width=50, height=4)
        txt.pack()
        print('Enter conversion')

        savexlsbut = tk.Button(conversionwin, text='Convert to XLS', font=('Arial', 12), width=30, height=1,
                               command=outputXLS)
        savexlsbut.pack()
        savexlsxbut = tk.Button(conversionwin, text='Convert to XLSX', font=('Arial', 12), width=30, height=1,
                                command=outputXLSX)
        savexlsxbut.pack()
        savecsvbut = tk.Button(conversionwin, text='Convert to CSV', font=('Arial', 12), width=30, height=1,
                               command=outputCSV)
        savecsvbut.pack()
        savetexbut = tk.Button(conversionwin, text='Convert to TEX', font=('Arial', 12), width=30, height=1,
                               command=outputTEX)
        savetexbut.pack()
        backbut = tk.Button(conversionwin, text='Back to Main Menu', font=('Arial', 12), width=30, height=1,
                            command=backmain)
        backbut.pack()
        conversionwin.mainloop()
        return ''

    def outputint():
        window.destroy()

        def backmain():
            outputwin.destroy()
            mainwindow()

        def nodelwin():
            messagebox.showinfo('Tip', 'Press Back to Main to Quit')
            return ''

        def outputlatex():
            global file_path
            file_path = filedialog.askopenfilename()
            print(file_path)
            nonlocal text1
            path, suffix = os.path.splitext(file_path)
            filesuffix = suffix.strip().lower()
            if filesuffix.lower() == '.csv':
                try:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', writeLaTeX('./tmptex.tex', readCSV(file_path)))
                except Exception as e:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', 'Error occurred when opening csv')
                else:
                    print('read csv successfully')
            elif filesuffix.lower() == '.xls':
                try:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', writeLaTeX('./tmptex.tex', read03xls(file_path)))
                except Exception as e:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', 'Error occurred when opening xls')
                else:
                    print('read csv successfully')
            elif filesuffix.lower() == '.xlsx':
                try:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', writeLaTeX('./tmptex.tex', read07xlsx(file_path)))
                except Exception as e:
                    text1.delete('1.0', tk.END)
                    text1.insert('insert', 'Error occurred when opening xlsx')
                else:
                    print('read csv successfully')
            else:
                text1.delete('1.0', tk.END)
                text1.insert('insert',
                             'Unsopported file suffix\nOnly ".xls", ".xlsx", ".csv" are permitted\n' + file_path)
            os.remove('./tmptex.tex')

        global file_path
        file_path = ''
        outputwin = tk.Tk()
        outputwin.title('Output LaTeX Source')
        outputwin.geometry('500x300')
        # outputwin.protocol("WM_DELETE_WINDOW", nodelwin)
        print('Enter outputfunc')
        text1 = tk.Text(outputwin, width=50, height=10, bg='cyan', font=('Arial', 12))
        text1.pack()
        outbut = tk.Button(outputwin, text='Choose File (xls, xlsx, csv)', font=('Arial', 12), width=30, height=1,
                           command=outputlatex)
        outbut.pack()
        backbut = tk.Button(outputwin, text='Back to Main Menu', font=('Arial', 12), width=30, height=1,
                            command=backmain)
        backbut.pack()
        outputwin.mainloop()

    window = tk.Tk()
    window.title('LaTeXfromExcel')  # set title
    window.geometry('500x300')  # set size

    banner = 'Convert Excel Data to LaTeX Source'
    txt = tk.Label(window, text=banner, font=('Arial', 12), width=30, height=6)
    txt.pack()
    conbut = tk.Button(window, text='File Format Conversion', font=('Arial', 12), width=30, height=1,
                       command=conversion)
    conbut.pack()
    outbut = tk.Button(window, text='Output LaTeX Source', font=('Arial', 12), width=30, height=1, command=outputint)
    outbut.pack()

    window.mainloop()  # show window


if __name__ == '__main__':
    mainwindow()
