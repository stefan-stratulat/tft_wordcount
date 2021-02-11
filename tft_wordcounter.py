import openpyxl

#input the file name (file must be in the same folder as the script)
instruction  = "Add the file name, insluding the extentions. For example: example.xlsx "
filename = input("Input file: ")

#open the file and read column G from each sheet
data = openpyxl.load_workbook(filename)

#loop to go trough each sheet
for sheet in data.get_sheet_names():
    print(sheet)
    ws = data[sheet]
    colF = ws['F']
    word_count = 0 #counter for the number of words in sheet

    #go trough each sheet and row values
    for row in colF:
        words = row.value
        if words is None:
            continue
        else:
        #make a list for the words in each row
            split_words = words.split()
        #count the words in each row
        counter = len(split_words)
        #add the number for words to the total words in sheet
        word_count = word_count+counter
    print(word_count)
#save the data based on sheet name
