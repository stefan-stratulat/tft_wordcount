import openpyxl
from tkinter import *
from tkinter import messagebox

root = Tk()
root.title("TfT word count")
root.configure(bg="#add8e6")
root.geometry('500x150')

#quick instructions for the user
instruction1  = Label(root,text="Add the file name in the box")
instruction2 = Label(root,text="The file should have the xlsx extention")
instruction3 = Label(root,text="The file be in the same folder as the script")
instruction1.pack()
instruction2.pack()
instruction3.pack()

#input box for the filename
filename = Entry(root,width=40, borderwidth=5)
filename.pack()

#function that counts the words
def click():
    #the filename given by user plus the xlsx
    excel = filename.get()+'.xlsx'
#open the file and read column F from each sheet
    data = openpyxl.load_workbook(excel)

    txt_file = open("word_count.txt",'w+')
    txt_file.write('Tab name'+" "+ 'No. of words'+'\n')
    #loop to go trough each sheet
    for sheet in data.get_sheet_names():
        ws = data[sheet]
        colF = ws['F']
        word_count = 0 #counter for the number of words in sheet

        #go trough each sheet and row values
        for row in colF:
            words = str(row.value)
            if words is None:
                continue
            else:
            #make a list for the words in each row
                split_words = words.split()
            #count the words in each row
            counter = len(split_words)
            #add the number for words to the total words in sheet
            word_count = word_count+counter
        #add info in txt_file
        txt_file.write(sheet+" "+ str(word_count)+'\n')

#add button to get words
button = Button(root,text="Count words",command=click)
button.pack()

root.mainloop()
