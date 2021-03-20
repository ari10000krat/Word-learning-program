from tkinter import *
from tkinter import filedialog
from PIL import ImageTk, Image
import PIL
import os
import openpyxl
import random
from copy import deepcopy
from tkinter import messagebox


def learningWindow():
    def SaveWindow():
        def SaveToExel():
            nonlocal CurrentArr
            CurrentDir = os.getcwd()
            file = f'{CurrentDir}\\{nmEntry.get()}.xlsx'
            wb = openpyxl.Workbook()
            sheet = wb.active
            print(CurrentArr)
            for i in range(len(CurrentArr)):
                for j in range(2):
                    sheet.cell(row=i+1,column=j+1).value=CurrentArr[i][j]
            wb.save(f'{nmEntry.get()}.xlsx')

        #nonlocal CurrentArr
        swWindow = Tk()
        swWindow['bg'] = lightBlue
        swLabel = Label(swWindow,
                        text='NAME:',
                        font=myFont,
                        bg=lightBlue)
        swLabel.pack()
        nmEntry= Entry(swWindow,
                       font=myFont,
                       bg=lightBlue)
        nmEntry.pack()
        okButton = Button(swWindow,
                          text='SAVE',
                          bg=lightBlue,
                          font=myFont,
                          command=SaveToExel)
        okButton.pack()

    def StartPenaltyRound():
        nonlocal CurrentArr
        nonlocal wordIndex
        nonlocal rounds
        nonlocal FailArr
        nonlocal CountOfWords
        nonlocal CountOfTrue

        CurrentArr = FailArr
        CountOfWords = len(CurrentArr)
        CountOfTrue = 0 # НОВОЕ
        FailArr = []
        random.shuffle(CurrentArr)
        wordIndex = 0
        rounds += 1

    def StartNewRound():
        nonlocal rounds
        nonlocal Progress
        nonlocal CountOfWords
        nonlocal CurrentArr
        nonlocal FailArr
        nonlocal wordIndex
        nonlocal CountOfTrue
        rounds += 1
        Progress = 0
        CountOfWords = len(wordsArr)
        CurrentArr = deepcopy(wordsArr)
        random.shuffle(CurrentArr)
        FailArr = []
        wordIndex = 0
        CountOfTrue = 0

    def UbdateStatistics():
        nonlocal Progress
        nonlocal window1
        if CountOfWords > 1:
            Progress = int(CountOfTrue / CountOfWords * 100)
            statisticsLabel.config(
                text=f'Round: {rounds}\nWords left: {CountOfWords - CountOfTrue}\nProgress: {Progress}%')
        else:
            statisticsLabel.config(
                text=f'Осталось 1 слово')

    def UbdateWords():
        nonlocal window1
        wordLabel.config(text=CurrentArr[wordIndex][mainIndex])
        wordLabelHelp.config(text=CurrentArr[wordIndex][secondIndex])

    def IsComplieted():  # TODO
        nonlocal CountOfTrue
        nonlocal CurrentArr
        nonlocal wordIndex
        if wordIndex == len(CurrentArr) - 1:  # конец массива
            if CountOfTrue != len(CurrentArr) - 1 and len(CurrentArr) != 1:  # есть ошибка    and len(CurArr) != 1
                return -1
            else:  # ошибок нет
                return 1
        else:  # не конец списка
            return 0

    def helpClick():
        if wordLabelHelp['fg'] == lightBlue:
            wordLabelHelp['fg'] = 'black'
        else:
            wordLabelHelp['fg'] = lightBlue

    def yesClick():  # TODO
        nonlocal CountOfTrue
        nonlocal wordIndex
        nonlocal window1
        nonlocal FailArr
        nonlocal CountOfWords
        nonlocal CurrentArr
        nonlocal Progress
        nonlocal rounds
        global wordsArr
        m = IsComplieted()
        if m == 0:  # Есть еще слова
            CountOfTrue += 1
            wordIndex += 1
            UbdateStatistics()
            UbdateWords()
        elif m == -1:  # Есть ошибки
            StartPenaltyRound()
            UbdateStatistics()
            UbdateWords()
        elif m == 1:
            UbdateStatistics()
            messagebox.showinfo('ИНФА', 'Вы выучили слова')
            StartNewRound()
            UbdateWords()
            UbdateStatistics()

    def noClick():  # TODO
        nonlocal CountOfTrue
        nonlocal wordIndex
        nonlocal window1
        nonlocal FailArr
        nonlocal CurrentArr
        nonlocal CountOfWords
        m = IsComplieted()
        if m == 0:
            if CurrentArr[wordIndex] not in FailArr:
                FailArr.append(CurrentArr[wordIndex])
            wordIndex += 1
            UbdateStatistics()
            UbdateWords()
        elif m == -1:
            if CurrentArr[wordIndex] not in FailArr:
                FailArr.append(CurrentArr[wordIndex])
            StartPenaltyRound()
            UbdateStatistics()
            UbdateWords()

    # Настройка формы***************************************************************************************
    global wordsArr
    global mode
    rounds = -1
    Progress = 0
    CountOfWords = 0
    CurrentArr = []
    FailArr = []
    wordIndex = 0
    CountOfTrue = 0

    if mode == 'DE -> RU':
        mainIndex = 0
        secondIndex = 1
    elif mode == 'RU -> DE':
        mainIndex = 1
        secondIndex = 0

    StartNewRound()

    window1 = Tk()
    window1.geometry('800x300')
    window1.config(bg='#AECAFC')
    window1.resizable(False, False)

    # Настройка вывода слова
    LabelFrame = Frame(
        window1,
        bg=lightBlue
    )

    wordLabel = Label(LabelFrame,
                      text=CurrentArr[wordIndex][mainIndex],
                      bg=lightBlue,
                      font=("Times", "24", "bold"),
                      justify='center',
                      width=43)
    wordLabel.pack()

    wordLabelHelp = Label(LabelFrame,
                          text=CurrentArr[wordIndex][secondIndex],
                          fg=lightBlue,
                          bg=lightBlue,
                          font=("Times", "24", "bold"),
                          justify='center',
                          width=43)
    wordLabelHelp.pack()

    LabelFrame.place(x=5, y=5)
    buttonsFrame = Frame(window1,
                         bg='gray'
                         )

    buttonYes = Button(buttonsFrame,
                       text='YES',
                       bg='green2',
                       font=("Times", "24"),
                       justify='center',
                       width=10,
                       command=yesClick
                       )
    buttonYes.pack(side=LEFT)

    buttonNo = Button(buttonsFrame,
                      text='NO',
                      bg='red3',
                      font=("Times", "24"),
                      justify='center',
                      width=10,
                      command=noClick
                      )
    buttonNo.pack(side=LEFT)

    buttonHelp = Button(buttonsFrame,
                        text='HELP',
                        bg='orange2',
                        font=("Times", "24", "bold italic"),
                        justify='center',
                        width=10,
                        command=helpClick
                        )
    buttonHelp.pack(side=LEFT)

    buttonsFrame.place(x=5, y=140)

    buttonSave = Button(window1,
                        text='SAVE',
                        bg='cyan3',
                        font=("Times", "24", "bold italic"),
                        justify='center',
                        width=31,
                        command=SaveWindow)
    buttonSave.place(x=5, y=202)

    statisticsLabel = Label(window1,
                            text=f'Round: {rounds}\nWords left: {CountOfWords - CountOfTrue}\nProgress: {Progress}%',
                            font=myFont,
                            bg=lightBlue,
                            justify='left',
                            width=15)
    statisticsLabel.place(x=580, y=140)

    window1.mainloop()

    # *************************************************************************************************************


if __name__ == '__main__':
    # Переменные####################################################################
    myFont = ('Times', '20')

    #   root = Tk()
    mode = 'DE -> RU'
    lightBlue = '#AECAFC'
    wordsArrIsReady = False
    wordsArr = [['kochen', 'готовить'],
                ['bleiben', 'оставаться'],
                ['fliegen', 'лететь'],
                ['tragen','одеть']]

    learningWindow()
