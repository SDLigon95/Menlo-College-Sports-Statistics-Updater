#!/usr/bin/env python2.7
# -*- coding: utf8 -*-
from Tkinter import *
from tkFileDialog import *
import os
# encoding=utf8
import sys

reload(sys)
sys.setdefaultencoding('utf8')
win = Tk()
win.geometry('575x357') # Size 200, 200
choice = 0

class UI(Frame):

    oldDocName = []
    newDocDirectory = []

    def constructButtons(self,choice):
        self.openButton = Button(win, text="Open Word Document", command=self.openFile, background='blue',
                                 fg='white', width=45, height=4)
        self.openButton.pack()
        self.saveAsButton = Button(win, text="Save As", command=self.saveFileAs, background='red', fg='white',
                                    width=45,
                                    height=4)
        self.saveAsButton.pack()
        if (choice==1):
            self.convertButton = Button(win, text="Convert Baseball Stats", command=self.convertBaseball,
                                        background='green', fg='white', width=45, height=4)
        if (choice==2):
            self.convertButton = Button(win, text="Convert Women's Soccer Stats", command=self.convertWomanSoccer,
                                        background='green', fg='white', width=45, height=4)

        if (choice==3):
            self.convertButton = Button(win, text="Convert Volleyball Stats", command=self.convertVolleyball,
                                        background='green', fg='white', width=45, height=4)
        self.convertButton.pack()

    def destroyButtons(self):


        self.openButton.destroy()
        self.saveAsButton.destroy()
        self.convertButton.destroy()
        self.backButton.destroy()
        self.basketballOption = Button(win, text="Baseball", command=self.baseballMenu, height=3)
        self.basketballOption.pack(fill=X)
        self.womanSoccerOption = Button(win, text="Soccer", command=self.womanSoccerMenu, height=3)
        self.womanSoccerOption.pack(fill=X)
        self.volleyballOption = Button(win, text="Volleyball", command=self.volleyballMenu, height=3)
        self.volleyballOption.pack(fill=X)
        self.pack()

    #Create UI
    def __init__(self):
        #Open and Convert button
        Frame.__init__(self)


        self.basketballOption = Button(win, text="Baseball", command=self.baseballMenu, height=3)
        self.basketballOption.pack(fill=X)
        self.womanSoccerOption = Button(win, text="Women's Soccer", command=self.womanSoccerMenu, height=3)
        self.womanSoccerOption.pack(fill=X)
        self.volleyballOption = Button(win, text="Volleyball", command=self.volleyballMenu, height=3)
        self.volleyballOption.pack(fill=X)
        self.pack()
        
    def baseballMenu(self):
        # self.tk_menuBar.add_command(label="Open", command=self.openFile)
        # menubar.add_command(label="Quit!", command=root.quit)
        self.volleyballOption.destroy()
        self.basketballOption.destroy()
        self.womanSoccerOption.destroy()
        self.backButton = Button(win, text="Back", command=self.destroyButtons, background='grey', fg='white',
                                 width=4, height=4)
        self.backButton.pack(side=LEFT)
        choice = 1
        self.constructButtons(choice)

        # Window
        # self.output = tk.Text(self, width=100, height=15, background='black', fg='white')
        # self.output.pack(side=tk.LEFT)
        # scrollbar
        # self.scrollbar = tk.Scrollbar(self,orient="vertical", command=self.output.yview)
        # self.scrollbar.pack(side=tk.RIGHT, fill="y")
        # self.output['yscrollcommand'] = self.scrollbar.set

        # Background and Pack
        self.configure(background="white")
        self.pack()
    
    def womanSoccerMenu(self):
        # self.tk_menuBar.add_command(label="Open", command=self.openFile)
        # menubar.add_command(label="Quit!", command=root.quit)
        self.volleyballOption.destroy()
        self.basketballOption.destroy()
        self.womanSoccerOption.destroy()
        self.backButton = Button(win, text="Back", command=self.destroyButtons, background='grey', fg='white',
                                 width=4, height=4)
        self.backButton.pack(side=LEFT)
        choice = 2
        self.constructButtons(choice)

        # Window
        # self.output = tk.Text(self, width=100, height=15, background='black', fg='white')
        # self.output.pack(side=tk.LEFT)
        # scrollbar
        # self.scrollbar = tk.Scrollbar(self,orient="vertical", command=self.output.yview)
        # self.scrollbar.pack(side=tk.RIGHT, fill="y")
        # self.output['yscrollcommand'] = self.scrollbar.set

        # Background and Pack
        self.configure(background="white")
        self.pack()

    def volleyballMenu(self):
        # self.tk_menuBar.add_command(label="Open", command=self.openFile)
        # menubar.add_command(label="Quit!", command=root.quit)
        self.volleyballOption.destroy()
        self.basketballOption.destroy()
        self.womanSoccerOption.destroy()
        self.backButton = Button(win, text="Back", command=self.destroyButtons, background='grey', fg='white',
                                 width=4, height=4)
        self.backButton.pack(side=LEFT)
        choice = 3
        self.constructButtons(choice)

        # Window
        # self.output = tk.Text(self, width=100, height=15, background='black', fg='white')
        # self.output.pack(side=tk.LEFT)
        # scrollbar
        # self.scrollbar = tk.Scrollbar(self,orient="vertical", command=self.output.yview)
        # self.scrollbar.pack(side=tk.RIGHT, fill="y")
        # self.output['yscrollcommand'] = self.scrollbar.set

        # Background and Pack
        self.configure(background="white")
        self.pack()


    #Grab file name and location from window
    def openFile(self):
        global oldDocName
        oldDocName = askopenfilename(filetypes=(("Word files", "*.docx"),))

        print(oldDocName)
        # return fileName

    def saveFileAs(self):
        global newDocDirectory
        newDocDirectory = askdirectory()
        if newDocDirectory is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return
        print newDocDirectory

    #Grabbing input to display in log window
        # def write (self, txt):
        #     self.output.insert(END,str(txt))
        #     self.update_idletasks()

    #Grabs input from website and compares names found in original document.
    #Scores will get updated based on which names are found.
    def convertBaseball(self):
        # Reading docx file
        import sys
        from docx import Document
        document = Document(oldDocName)
        tables = document.tables
        section = []
        i = 0
        for table in tables:
            section.append(table)
        # document.save('MEN Game Notes Hitting_Updated.docx')
        #
        working = section[1]
        # self.write("Reading document...\n")
        content = []
        flag = 0
        lastNameCounter = 0
        #Import library for reading website
        from bs4 import BeautifulSoup
        import urllib2
        import pandas as pd
        # download html from link
        url = "http://www.dakstats.com/WebSync/Pages/Team/IndividualStats.aspx?association=10&sg=MBA&sea=NAIMBA_2017&team=9101"
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "html.parser")
        table = soup.find("table", {"class": "gridViewReportBuilderWide"})
        avg_scores = []
        names = []
        count = 1
        dic = {}
        row_counter = 0
        for row in table.find_all('tr')[1:]:
            col = row.find_all('td')
            row_counter = row_counter + 1
            if len(col) == 22:
                name = col[0].find(text=True)
                names.append(name)
                GP = col[1].find(text=True)
                #
                #average = field 1
                average = col[3].find(text=True)
                #OB = field 2
                OB = col[17].find(text=True)
                #SLG = field 3
                SLG = col[12].find(text=True)
                #
                avg_scores.append(average)
                R = col[5].find(text=True)
                #2B
                TwoB = col[7].find(text=True)
                #3B
                ThreeB = col[8].find(text=True)
                #HR
                HR = col[9].find(text=True)
                #BB
                BB = col[13].find(text=True)
                #K not found
                #SB
                SB = col[20].find(text=True)

                
                try:
                    for row in working.rows:
                        for cell in row.cells:
                            for index, paragraph in enumerate(cell.paragraphs):
                                content.append(paragraph.text)
                                switch = False
                                if flag == 1:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + average
                                    paragraph.text = average
                                    flag = flag + 1
                                    continue
                                if flag == 2:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + OB
                                    paragraph.text = OB
                                    flag = flag + 1
                                    continue
                                if flag == 3:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + SLG
                                    paragraph.text = SLG
                                    flag = flag + 1
                                    continue
                                if flag > 3:
                                    # print paragraph.text + '\n'
                                   
                                    
                                    if paragraph.text.__contains__('2B'):
                                        print '2B found'
                                        paragraph.text = "2B - " + TwoB

                                    if paragraph.text.__contains__('3B'):
                                        print '3B found'
                                        paragraph.text = "3B - " + ThreeB

                                    if paragraph.text.__contains__('HR'):
                                        print 'HR found'
                                        paragraph.text = "HR - " + HR

                                    if paragraph.text.__contains__('BB'):
                                        print 'BB found'
                                        paragraph.text = "BB - " + BB

                                    if paragraph.text.__contains__('K'):
                                        print 'K found'

                                    if paragraph.text.__contains__('SB'):
                                        print 'SB found'
                                        paragraph.text = "SB - " + SB
                                    
                                # EXPERIMENT FOR DUPLICATE NAMES
                                    # if paragraph.text == '3':
                                    #     print '3 FOUND YAY!'   

                                    # if paragraph.text != (name.split(', ')[0]):
                                    #     lastNameCounter = lastNameCounter + 1
                                    
                                    # if lastNameCounter >= 14:
                                    #     if paragraph.text != (name.split(', ')[0]):
                                    #         print "NONE!!"
                                    #     if paragraph.text == (name.split(', ')[0]):
                                    #         print "QQQQ!!"
                                    
                                    if paragraph.text == (name.split(', ')[0]):
                                        # print 'LAST NAME- ' + name.split(', ')[0] + ' FOUND!!!'
                                        # print lastNameCounter
                                        # lastNameCounter = 0
                                        flag = 0
                                        break
                                        
                                
                                if paragraph.text == name.split(' ')[1]:
                                    flag = flag + 1
                                    # self.write(name + '\n')
                                    print paragraph.text

                                    #scan has a dictionary of all the names found on the website
                                    #the program will check whether or not a first name has already been recorded
                                    #if it has, then the counter will go up from 1 to 2
                                    #if the counter is more than 1, then the program will have to match the specific name
                                    #from the document with a matching number
                                    scan = dic.keys()
                                    for index in scan:
                                        if index.__contains__(name.split(' ')[1]):
                                            switch = True
                                            count = count + 1
                                            dic[name] = count
                                            count = 1
                                    if switch == True:
                                        continue
                                    dic[name] = count
                                    # paragraph.text[index + 1] = '.555'
                                    # print paragraph.text[index + 1]
                except IndexError:
                    print "list index out of range"

        print "almost done"
        # print dic.items()

        print "Check on these players' stats."
        for name in dic.keys():
            if dic[name] > 1:
                print name + '\n'
        print " Information might be wrong due to multiple first names appearing."
            
        
        # for key, value in dic.items():
#             print key + " => " + value
        # print names
        # print row_counter
        document.add_table(1, 2, style=None)
        document.save(newDocDirectory+"/Updated File.docx")
        sys.exit()
        # self.write("Document is fully updated\n")
    
    def convertWomanSoccer(self):
        # Reading docx file
        import sys
        from docx import Document
        document = Document(oldDocName)
        tables = document.tables
        section = []
        i = 0
        for table in tables:
            section.append(table)
        # document.save('MEN Game Notes Hitting_Updated.docx')
        #
        working = section[1]
        # self.write("Reading document...\n")
        content = []
        flag = 0
        lastNameCounter = 0
        #Import library for reading website
        from bs4 import BeautifulSoup
        import urllib2
        import pandas as pd
        # download html from link
        url = "http://www.dakstats.com/WebSync/Pages/Team/IndividualStats.aspx?association=10&sg=WSO&sea=NAIWSO_2017&team=2409"
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "html.parser")
        table = soup.find("table", {"class": "gridViewReportBuilderWide"})
        avg_scores = []
        names = []
        count = 1
        dic = {}
        row_counter = 0
        for row in table.find_all('tr')[1:]:
            col = row.find_all('td')
            row_counter = row_counter + 1
            if len(col) == 16:
                name = col[1].find(text=True)
                names.append(name)
                #col[0]= Player Number
                GP = col[2].find(text=True)
                #average = field 1
                GS = col[3].find(text=True)
                #OB = field 2
                G = col[4].find(text=True)
                #SLG = field 3
                A = col[5].find(text=True)
                #

                pts = col[6].find(text=True)

                SH = col[7].find(text=True)

                SOG = col[8].find(text=True)

                YC = col[9].find(text=True)

                GWG = col[11].find(text=True)

                PKM = col[12].find(text=True)

                
                try:
                    for row in working.rows:
                        for cell in row.cells:
                            for index, paragraph in enumerate(cell.paragraphs):
                                content.append(paragraph.text)
                                switch = False
                                if flag == 1:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + G
                                    paragraph.text = G
                                    flag = flag + 1
                                    continue
                                if flag == 2:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + SOG
                                    paragraph.text = SOG
                                    flag = flag + 1
                                    continue
                                # if flag == 3:
                                #     # print "doc: " + paragraph.text
                                #     # print "website: " + SLG
                                #     paragraph.text = SLG
                                #     flag = flag + 1
                                #     continue
                                if flag > 2:
                                    if paragraph.text.__contains__('GP'):
                                        print 'GP found'
                                        paragraph.text = "GP - " + GP

                                    if paragraph.text.__contains__('GS'):
                                        print 'GS found'
                                        paragraph.text = "GS - " + GS

                                    if paragraph.text.__contains__('A'):
                                        print 'A found'
                                        paragraph.text = "HR - " + A

                                    if paragraph.text.__contains__('pts'):
                                        print 'pts found'
                                        paragraph.text = "pts - " + pts

                                    if paragraph.text.__contains__('SH'):
                                        print 'SH found'
                                        paragraph.text = "SH - " + SH

                                    if paragraph.text.__contains__('SOG'):
                                        print 'SOG found'
                                        paragraph.text = "SOG - " + SOG

                                    if paragraph.text.__contains__('YC'):
                                        print 'YC found'
                                        paragraph.text = "YC - " + YC

                                    if paragraph.text.__contains__('GWG'):
                                        print 'GWG found'
                                        paragraph.text = "GWG - " + GWG

                                    if paragraph.text.__contains__('PKM'):
                                        print 'PKM found'
                                        paragraph.text = "PKM - " + PKM
                                    
                                # EXPERIMENT FOR DUPLICATE NAMES
                                    # if paragraph.text == '3':
                                    #     print '3 FOUND YAY!'   

                                    # if paragraph.text != (name.split(', ')[0]):
                                    #     lastNameCounter = lastNameCounter + 1
                                    
                                    # if lastNameCounter >= 14:
                                    #     if paragraph.text != (name.split(', ')[0]):
                                    #         print "NONE!!"
                                    #     if paragraph.text == (name.split(', ')[0]):
                                    #         print "QQQQ!!"
                                    
                                    if paragraph.text == (name.split(', ')[0]):
                                        # print 'LAST NAME- ' + name.split(', ')[0] + ' FOUND!!!'
                                        # print lastNameCounter
                                        # lastNameCounter = 0
                                        flag = 0
                                        break
                                    flag = flag + 1

                                if paragraph.text == name.split(' ')[1]:
                                    flag = flag + 1
                                    # self.write(name + '\n')
                                    print paragraph.text

                                    #scan has a dictionary of all the names found on the website
                                    #the program will check whether or not a first name has already been recorded
                                    #if it has, then the counter will go up from 1 to 2
                                    #if the counter is more than 1, then the program will have to match the specific name
                                    #from the document with a matching number
                                    scan = dic.keys()
                                    for index in scan:
                                        if index.__contains__(name.split(' ')[1]):
                                            switch = True
                                            count = count + 1
                                            dic[name] = count
                                            count = 1
                                    if switch == True:
                                        continue
                                    dic[name] = count
                                    # paragraph.text[index + 1] = '.555'
                                    # print paragraph.text[index + 1]
                except IndexError:
                    print "list index out of range"
            #GOALIE STATS
            if len(col) == 13:
                name = col[1].find(text=True)
                names.append(name)
                # col[0]= Player Number
                GP = col[2].find(text=True)
                # average = field 1
                GS = col[3].find(text=True)
                # OB = field 2
                GA = col[5].find(text=True)
                # SLG = field 3
                GAAvg = col[6].find(text=True)
                #

                SV = col[7].find(text=True)

                Pct = col[8].find(text=True)

                W = col[9].find(text=True)

                L = col[10].find(text=True)

                T = col[11].find(text=True)

                SHO = col[12].find(text=True)

                try:
                    for row in working.rows:
                        for cell in row.cells:
                            for index, paragraph in enumerate(cell.paragraphs):
                                content.append(paragraph.text)
                                switch = False
                                if flag == 1:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + G
                                    paragraph.text = GAAvg
                                    flag = flag + 1
                                    continue
                                if flag == 2:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + SOG
                                    paragraph.text = Pct
                                    flag = flag + 1
                                    continue
                                if flag == 3:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + SLG
                                    paragraph.text = SHO
                                    flag = flag + 1
                                    continue
                                if flag > 3:
                                    if paragraph.text.__contains__('GA'):
                                        print 'GA found'
                                        paragraph.text = "GA - " + GA

                                    if paragraph.text.__contains__('SV'):
                                        print 'SV found'
                                        paragraph.text = "SV - " + SV

                                    if paragraph.text.__contains__('W'):
                                        print 'W found'
                                        paragraph.text = "HR - " + W

                                    if paragraph.text.__contains__('L'):
                                        print 'L found'
                                        paragraph.text = "L - " + L

                                    if paragraph.text.__contains__('SH'):
                                        print 'SH found'
                                        paragraph.text = "SH - " + SH

                                    if paragraph.text.__contains__('T'):
                                        print 'T found'
                                        paragraph.text = "T - " + T

                                    if paragraph.text.__contains__('YC'):
                                        print 'YC found'
                                        paragraph.text = "YC - " + YC


                                        # EXPERIMENT FOR DUPLICATE NAMES
                                    # if paragraph.text == '3':
                                    #     print '3 FOUND YAY!'

                                    # if paragraph.text != (name.split(', ')[0]):
                                    #     lastNameCounter = lastNameCounter + 1

                                    # if lastNameCounter >= 14:
                                    #     if paragraph.text != (name.split(', ')[0]):
                                    #         print "NONE!!"
                                    #     if paragraph.text == (name.split(', ')[0]):
                                    #         print "QQQQ!!"

                                    if paragraph.text == (name.split(', ')[0]):
                                        # print 'LAST NAME- ' + name.split(', ')[0] + ' FOUND!!!'
                                        # print lastNameCounter
                                        # lastNameCounter = 0
                                        flag = 0
                                        break
                                    flag = flag + 1

                                if paragraph.text == name.split(' ')[1]:
                                    flag = flag + 1
                                    # self.write(name + '\n')
                                    print paragraph.text

                                    # scan has a dictionary of all the names found on the website
                                    # the program will check whether or not a first name has already been recorded
                                    # if it has, then the counter will go up from 1 to 2
                                    # if the counter is more than 1, then the program will have to match the specific name
                                    # from the document with a matching number
                                    scan = dic.keys()
                                    for index in scan:
                                        if index.__contains__(name.split(' ')[1]):
                                            switch = True
                                            count = count + 1
                                            dic[name] = count
                                            count = 1
                                    if switch == True:
                                        continue
                                    dic[name] = count
                                    # paragraph.text[index + 1] = '.555'
                                    # print paragraph.text[index + 1]
                except IndexError:
                    print "list index out of range"

        print "almost done"
        print dic
        print names
        print row_counter
        document.add_table(1, 2, style=None)
        document.save(newDocDirectory+"/Updated File.docx")
        # self.write("Document is fully updated\n")


        # Reading docx file
        import sys
        from docx import Document
        document = Document(oldDocName)
        tables = document.tables
        section = []
        i = 0
        for table in tables:
            section.append(table)
        # document.save('MEN Game Notes Hitting_Updated.docx')
        #
        working = section[1]
        # self.write("Reading document...\n")
        content = []
        flag = 0
        lastNameCounter = 0
        # Import library for reading website
        from bs4 import BeautifulSoup
        import urllib2
        import pandas as pd
        # download html from link
        url = "http://www.dakstats.com/WebSync/Pages/Team/IndividualStats.aspx?association=10&sg=MBA&sea=NAIMBA_2017&team=9101"
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "html.parser")
        table = soup.find("table", {"class": "gridViewReportBuilderWide"})
        avg_scores = []
        names = []
        count = 1
        dic = {}
        row_counter = 0
        for row in table.find_all('tr')[1:]:
            col = row.find_all('td')
            row_counter = row_counter + 1
            if len(col) == 22:
                name = col[0].find(text=True)
                names.append(name)
                GP = col[1].find(text=True)
                #
                # average = field 1
                average = col[3].find(text=True)
                # OB = field 2
                OB = col[17].find(text=True)
                # SLG = field 3
                SLG = col[12].find(text=True)
                #
                avg_scores.append(average)
                R = col[5].find(text=True)
                # 2B
                TwoB = col[7].find(text=True)
                # 3B
                ThreeB = col[8].find(text=True)
                # HR
                HR = col[9].find(text=True)
                # BB
                BB = col[13].find(text=True)
                # K not found
                # SB
                SB = col[20].find(text=True)

                try:
                    for row in working.rows:
                        for cell in row.cells:
                            for index, paragraph in enumerate(cell.paragraphs):
                                content.append(paragraph.text)
                                switch = False
                                if flag == 1:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + average
                                    paragraph.text = average
                                    flag = flag + 1
                                    continue
                                if flag == 2:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + OB
                                    paragraph.text = OB
                                    flag = flag + 1
                                    continue
                                if flag == 3:
                                    # print "doc: " + paragraph.text
                                    # print "website: " + SLG
                                    paragraph.text = SLG
                                    flag = flag + 1
                                    continue
                                if flag > 3:
                                    # print paragraph.text + '\n'


                                    if paragraph.text.__contains__('2B'):
                                        print '2B found'
                                        paragraph.text = "2B - " + TwoB

                                    if paragraph.text.__contains__('3B'):
                                        print '3B found'
                                        paragraph.text = "3B - " + ThreeB

                                    if paragraph.text.__contains__('HR'):
                                        print 'HR found'
                                        paragraph.text = "HR - " + HR

                                    if paragraph.text.__contains__('BB'):
                                        print 'BB found'
                                        paragraph.text = "BB - " + BB

                                    if paragraph.text.__contains__('K'):
                                        print 'K found'

                                    if paragraph.text.__contains__('SB'):
                                        print 'SB found'
                                        paragraph.text = "SB - " + SB

                                        # EXPERIMENT FOR DUPLICATE NAMES
                                    # if paragraph.text == '3':
                                    #     print '3 FOUND YAY!'

                                    # if paragraph.text != (name.split(', ')[0]):
                                    #     lastNameCounter = lastNameCounter + 1

                                    # if lastNameCounter >= 14:
                                    #     if paragraph.text != (name.split(', ')[0]):
                                    #         print "NONE!!"
                                    #     if paragraph.text == (name.split(', ')[0]):
                                    #         print "QQQQ!!"

                                    if paragraph.text == (name.split(', ')[0]):
                                        # print 'LAST NAME- ' + name.split(', ')[0] + ' FOUND!!!'
                                        # print lastNameCounter
                                        # lastNameCounter = 0
                                        flag = 0
                                        break

                                if paragraph.text == name.split(' ')[1]:
                                    flag = flag + 1
                                    # self.write(name + '\n')
                                    print paragraph.text

                                    # scan has a dictionary of all the names found on the website
                                    # the program will check whether or not a first name has already been recorded
                                    # if it has, then the counter will go up from 1 to 2
                                    # if the counter is more than 1, then the program will have to match the specific name
                                    # from the document with a matching number
                                    scan = dic.keys()
                                    for index in scan:
                                        if index.__contains__(name.split(' ')[1]):
                                            switch = True
                                            count = count + 1
                                            dic[name] = count
                                            count = 1
                                    if switch == True:
                                        continue
                                    dic[name] = count
                                    # paragraph.text[index + 1] = '.555'
                                    # print paragraph.text[index + 1]
                except IndexError:
                    print "list index out of range"

        print "almost done"
        # print dic.items()

        print "Check on these players' stats."
        for name in dic.keys():
            if dic[name] > 1:
                print name + '\n'
        print " Information might be wrong due to multiple first names appearing."

        # for key, value in dic.items():
        #             print key + " => " + value
        # print names
        # print row_counter
        document.add_table(1, 2, style=None)
        document.save(newDocDirectory + "/Updated File.docx")
        sys.exit()
        # self.write("Document is fully updated\n")

    def convertVolleyball(self):
        # Reading docx file
        import sys
        from docx import Document
        document = Document(oldDocName)
        tables = document.tables
        section = []
        i = 0
        for table in tables:
            section.append(table)
        # document.save('MEN Game Notes Hitting_Updated.docx')
        #
        working = section[1]
        # self.write("Reading document...\n")
        content = []
        flag = 0
        lastNameCounter = 0
        # Import library for reading website
        from bs4 import BeautifulSoup
        import urllib2
        import pandas as pd
        # download html from link
        url = "http://www.dakstats.com/WebSync/Pages/Team/IndividualStats.aspx?association=10&sg=WVB&sea=NAIWVB_2017&team=2030"
        page = urllib2.urlopen(url)
        soup = BeautifulSoup(page, "html.parser")
        table = soup.find("table", {"class": "gridViewReportBuilderWide"})
        avg_scores = []
        names = []
        count = 1
        dic = {}
        row_counter = 0
        for row in table.find_all('tr')[1:]:
            col = row.find_all('td')
            row_counter = row_counter + 1
            if row_counter > 14:
                break
            if len(col) == 30:
                name = col[0].find('a', href=True)
                name = name.get('title')
                names.append(name.split(', ')[1])
                print name.split(', ')[1]

                GP = col[1].find(text=True)
                print "GP: " + GP
                #
                # average = field 1
                MP = col[2].find(text=True)
                print "MP: " + MP

                K = col[3].find(text=True)
                print "K: " + K

                A = col[8].find(text=True)
                print "A: " + A

                # avg_scores.append(average)

                SA = col[12].find(text=True)
                print "SA: " + SA

                SE = col[14].find(text=True)
                print "SE: " + SE

                DIG = col[20].find(text=True)
                print "DIG: " + DIG

                BS = col[22].find(text=True)
                print "BS: " + BS

                BA = col[23].find(text=True)
                print "BA: " + BA

                TOT = col[24].find(text=True)
                print "TOT: " + TOT

                try:
                    print "TESTING"
                    for row in working.rows:
                        for cell in row.cells:
                            for index, paragraph in enumerate(cell.paragraphs):
                                content.append(paragraph.text)
                                switch = False
                            if flag == 1:
                                print "BS FOUND: " + BS
                                paragraph.text = BS
                                flag = flag + 1
                                continue
                            if flag == 2:
                                print "BA FOUND: " + BA
                                paragraph.text = BA
                                flag = flag + 1
                                continue
                            if flag == 3:
                                print "TOT FOUND: " + TOT
                                paragraph.text = TOT
                                flag = flag + 1
                                continue
                            if flag > 3:
                                # print paragraph.text + '\n'

                                if paragraph.text.__contains__('GP-'):
                                    print 'GP found'
                                    paragraph.text = "GP- " + GP

                                if paragraph.text.__contains__('MP-'):
                                    print 'MP found'
                                    paragraph.text = "MP- " + MP

                                if paragraph.text.__contains__('K-'):
                                    print 'K found'
                                    paragraph.text = "K- " + K

                                if paragraph.text.__contains__('SA-'):
                                    print 'SA found'
                                    paragraph.text = "SA- " + SA

                                if paragraph.text.__contains__('SE-'):
                                    print 'SE found'
                                    paragraph.text = "SE- " + SE

                                if paragraph.text.__contains__('AST-'):
                                    print 'A found' + A
                                    paragraph.text = "AST- " + A
                                    print A

                                if paragraph.text.__contains__('DIG-'):
                                    print 'DIG found'
                                    paragraph.text = "DIG - " + DIG
                                    flag = 0
                                    print "ROW COUNTER: " + str(row_counter)
                                    document.save(newDocDirectory + "/Updated File.docx")
                                    break



                                    # EXPERIMENT FOR DUPLICATE NAMES
                                    # if paragraph.text == '3':
                                    #     print '3 FOUND YAY!'

                                    # if paragraph.text != (name.split(', ')[0]):
                                    #     lastNameCounter = lastNameCounter + 1

                                    # if lastNameCounter >= 14:
                                    #     if paragraph.text != (name.split(', ')[0]):
                                    #         print "NONE!!"
                                    #     if paragraph.text == (name.split(', ')[0]):
                                    #         print "QQQQ!!"

                                    # if paragraph.text == (name.split(', ')[0]):
                                    #     # print 'LAST NAME- ' + name.split(', ')[0] + ' FOUND!!!'
                                    #     # print lastNameCounter
                                    #     # lastNameCounter = 0

                            if name.split(', ')[1] == 'K.':
                                if paragraph.text == "Ka’imilani":
                                    flag = flag + 1
                                # scan has a dictionary of all the names found on the website
                                # the program will check whether or not a first name has already been recorded
                                # if it has, then the counter will go up from 1 to 2
                                # if the counter is more than 1, then the program will have to match the specific name
                                # from the document with a matching number
                                scan = dic.keys()
                                for index in scan:
                                    if index.__contains__(name.split(', ')[1]):
                                        switch = True
                                        count = count + 1
                                        dic[name] = count
                                        count = 1

                                if switch == True:
                                    continue
                                dic[name] = count

                            if name.split(', ')[0] == '10 SCOTT':
                                print "SCOTT PASSED"
                                if paragraph.text == "Jaden":
                                    flag = flag + 1

                                # scan has a dictionary of all the names found on the website
                                # the program will check whether or not a first name has already been recorded
                                # if it has, then the counter will go up from 1 to 2
                                # if the counter is more than 1, then the program will have to match the specific name
                                # from the document with a matching number
                                scan = dic.keys()
                                for index in scan:
                                    if index.__contains__(name.split(', ')[1]):
                                        switch = True
                                        count = count + 1
                                        dic[name] = count
                                        count = 1

                                if switch == True:
                                    continue
                                dic[name] = count
                                
                            if paragraph.text == name.split(', ')[1]:
                                print "NAMES: " + paragraph.text + " " + name.split(', ')[1]
                                flag = flag + 1
                                # self.write(name + '\n')
                                print paragraph.text
                                print "TEST2"

                                # scan has a dictionary of all the names found on the website
                                # the program will check whether or not a first name has already been recorded
                                # if it has, then the counter will go up from 1 to 2
                                # if the counter is more than 1, then the program will have to match the specific name
                                # from the document with a matching number
                                scan = dic.keys()
                                for index in scan:
                                    if index.__contains__(name.split(', ')[1]):
                                        switch = True
                                        count = count + 1
                                        dic[name] = count
                                        count = 1

                                if switch == True:
                                    continue
                                dic[name] = count
                                # paragraph.text[index + 1] = '.555'
                                # print paragraph.text[index + 1]
                except IndexError:
                    print "list index out of range"

        print "almost done"
        # print dic.items()

        print "Check on these players' stats."
        for name in dic.keys():
            if dic[name] > 1:
                print name + '\n'
        print " Information might be wrong due to multiple first names appearing."

        # for key, value in dic.items():
        #             print key + " => " + value
        # print names
        # print row_counter
        document.add_table(1, 2, style=None)
        document.save(newDocDirectory + "/Updated File.docx")
        sys.exit()
        # self.write("Document is fully updated\n")
if __name__ == '__main__':
    UI().mainloop()
