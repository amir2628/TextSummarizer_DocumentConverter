from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import *
import os
import spacy
from spacy.lang.en.stop_words import STOP_WORDS
from string import punctuation
from heapq import nlargest
from PIL import ImageTk, Image
import customtkinter
import pdfplumber
from tkinter import messagebox, Canvas
from tkinter.filedialog import asksaveasfile
import tkinter.filedialog as fudder
import tkinter.messagebox as mbox
import docx2txt
from io import BufferedReader
import tkinter.ttk as ttk
from tktooltip import ToolTip
import webbrowser

Word2TextFile = ""
text=""
summary_file=""
summary=""
languageis = ""
fullPdfPath = ""
wordDir = ""
txtAlreadyDir = ""
new = 1
url = "https://github.com/amir2628"

class MainApplication(Frame):
    def __init__(self, parent, *args, **kwargs):
        Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

    # def choose_file(): #Choosing the PDF file
    #     Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    #     global filename
    #     filename = askopenfilename(filetypes=[('pdf file', '*.pdf'),("word Document", '*.doc'),("word Document", '*.docx')] ) # show an "Open" dialog box and return the path to the selected file
    #     if (len(filename) == 0):
    #         quit()
    #     else:
    #         # Entry Field to show the chosen file path
    #         mbox.showinfo(message="File Path is : '" + str(filename)+ "'")

    def choose_already_txt(): #Choosing the PDF file
        Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        global txtAlreadyDir
        txtAlreadyDir = askopenfilename(filetypes=[('txt file', '*.txt'), ('all files', '*.*')]) # show an "Open" dialog box and return the path to the selected file
        if (len(txtAlreadyDir) == 0):
            quit()
        else:
            # Entry Field to show the chosen file path
            mbox.showinfo(message="File Path is : '" + str(txtAlreadyDir)+ "'")

    def close_app(): # closing the App
        quit()
        window.destroy()

    def wordtoText():
        global wordDir
        global Word2TextFile
        wordDir = askopenfilename(filetypes=[('Word Document', '*.doc'),('Word Document', '*.docx')]) # show an "Open" dialog box and return the path to the selected file
        if (len(wordDir) == 0):
            quit()
        else:
            mbox.showinfo(message="File Path is : '" + str(wordDir)+ "'")
        WordDoc = docx2txt.process(wordDir)
        Word2TextFile = asksaveasfile(mode = 'wt', title="Select Location to save the converted word document", filetypes=[('txt file', '*.txt')])
        if Word2TextFile is None:
            return
        else:
            with open(Word2TextFile.name, 'wt', encoding="utf-8") as text_file:
                print(WordDoc, file=text_file)
                # doc.save("Word2TxtOutput.txt")
                Word2TextFile.close()
        

    def pdfTotext(): #Convert PDF to Text

        Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        global filename
        filename = askopenfilename(filetypes=[('pdf file', '*.pdf')] ) # show an "Open" dialog box and return the path to the selected file
        if (len(filename) == 0):
            quit()
        else:
            # Entry Field to show the chosen file path
            mbox.showinfo(message="File Path is : '" + str(filename)+ "'")

        pdftext=""
        print(filename)
        with pdfplumber.open(filename) as pdf:
            pages = pdf.pages[0:]
            print(pages)
            for item in pages:
                pdftext += item.extract_text()
        saveDir=os.path.dirname(filename)
        with open(os.path.join(str(saveDir), "GUITextfrompdf.txt"), 'w', encoding="utf-8") as f:
            f.write(pdftext)

        PDF2TextFile = asksaveasfile(mode = 'wt', title="Select Location to save the converted word document", filetypes=[('txt file', '*.txt')])
        if PDF2TextFile is None:
            return
        else:
            with open(PDF2TextFile.name, 'wt', encoding="utf-8") as text_file:
                print(pdftext, file=text_file)
                # doc.save("Word2TxtOutput.txt")
                PDF2TextFile.close()
        mbox.showinfo(message= "File has been saved in Path: '"+ str(PDF2TextFile))
        global fullPdfPath
        # fullPdfPath=str(saveDir)+'/GUITextfrompdf.txt'
        fullPdfPath = PDF2TextFile


    def select_language(): #choosing the pdf language
        global languageis
        languageis = ""
        global language
        languageis = language.get()
        language_label = Label(window, text="You have chosen "+str(languageis), font=('Comic Sans MS', 9, 'italic'), fg='#ff006e', bg="#00072d", wraplength=200, justify="center")
        language_label.place(x=630, y=270)
    
    def clearAfterOneRun():
        global text
        text=""
        global summary_file
        summary_file=""
        global summary
        summary=""
        global fullPdfPath
        fullPdfPath=""
        global wordDir
        wordDir=""
        global txtAlreadyDir
        txtAlreadyDir=""
        global languageis
        languageis = ""
        global Word2TextFile
        Word2TextFile = ""
        mbox.showinfo(message= "Cleared everything, now you can start over")


    def summarize(): #Create summary
        global summary_file
        global summary
        global text
        if fullPdfPath != "":
            text = fullPdfPath
        elif wordDir != "":
            text = Word2TextFile
        else:
            text = txtAlreadyDir
        per=0.5
        if languageis == 'Russian':
            nlp = spacy.load('ru_core_news_lg')
        else:
            nlp = spacy.load('en_core_web_trf')
        if languageis =='Russian':
                if txtAlreadyDir !="":
                        with open(text, encoding='utf-8', mode = "r") as te: #problem here with decoding Windows-1251
                            print("Russian Loop: "+str(te))
                            article = te.read()
                            print(article)
                else:
                        with open(text.name, encoding='utf-8', mode = "r") as te: #problem here with decoding Windows-1251
                            print("Russian Loop: "+str(te))
                            article = te.read()
                            print(article)
        else:
                if txtAlreadyDir !="":
                    with open(text, encoding='utf-8', mode = "r") as te: #problem here with decoding Windows-1251
                        print("Russian Loop: "+str(te))
                        article = te.read()
                        print(article)
                else:
                        if text == str:
                            with open(text, encoding='utf-8', mode = "r") as te: #problem here with decoding Windows-1251
                                print("Russian Loop: "+str(te))
                                article = te.read()
                                print(article)
                        else:
                            with open(text.name, encoding='utf-8', mode = "r") as te: #problem here with decoding Windows-1251
                                print("Russian Loop: "+str(te))
                                article = te.read()
                                print(article)
        doc= nlp(article)
        tokens=[token.text for token in doc]
        word_frequencies={}
        for word in doc:
            if word.text.lower() not in list(STOP_WORDS):
                if word.text.lower() not in punctuation:
                    if word.text not in word_frequencies.keys():
                        word_frequencies[word.text] = 1
                    else:
                        word_frequencies[word.text] += 1
        max_frequency=max(word_frequencies.values())
        for word in word_frequencies.keys():
            word_frequencies[word]=word_frequencies[word]/max_frequency
        sentence_tokens= [sent for sent in doc.sents]
        sentence_scores = {}
        for sent in sentence_tokens:
            for word in sent:
                if word.text.lower() in word_frequencies.keys():
                    if sent not in sentence_scores.keys():                            
                        sentence_scores[sent]=word_frequencies[word.text.lower()]
                    else:
                        sentence_scores[sent]+=word_frequencies[word.text.lower()]
        select_length=int(len(sentence_tokens)*per)
        summary=nlargest(select_length, sentence_scores,key=sentence_scores.get)
        final_summary=[word.text for word in summary]
        summary=''.join(final_summary)
        summary_file = summary
        return summary_file
        

    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            window.destroy()

    def save():
            # function to call when user press
            # the save button, a filedialog will
            # open and ask to save file
        file = asksaveasfile(mode = 'wb', title="Select Location", filetypes=[('txt file', '*.txt')])
        if file is None:
            return
        else:
            # with open(file, 'w', encoding="utf-8") as f:
                file.write(summary_file.encode('utf-8'))
                file.close()

    def openGithub():
        webbrowser.open(url,new=new)
    
    def aboutInfo():
        mbox.showinfo("Disclaimer",
                                "This App has been created because I needed such an app at the time\n"
                                "\n"
                                "I did not find a script or a website which could summarize Russian text, so I decided to create one myself\n"
                                "\n"
                                "I have used many libraries here and got help from other people in the community\n"
                                "\n"
                                "with the help of this app you can do many things: \n"
                                "1. Convert PDF files to text file\n"
                                "2. Convert Word (.docx) files to text file\n"
                                "3. Create summary from the text files.\n"
                                "\n"
                                "Also Check out my GitHub")

if __name__ == "__main__":
    # window=Tk()
    window=customtkinter.CTk()
    #Adding a funny icon
    scriptdirname = os.path.dirname(__file__)
    iconimage = os.path.join(scriptdirname, "icons", "Santa.ico")
    window.iconbitmap(iconimage)

    #Menu bar
    menubar = Menu(window)

    PDFmenu = Menu(menubar, tearoff=0)
    PDFmenu.add_command(label="Start PDF Conversion to text", command=MainApplication.pdfTotext)
    menubar.add_cascade(label="PDF to text Convert", menu=PDFmenu)

    wordMenu = Menu(menubar, tearoff=0)
    wordMenu.add_command(label="Start Word (.docx) Conversion to text", command=MainApplication.wordtoText)
    menubar.add_cascade(label="Word to text Convert", menu=wordMenu)

    alreadyTxtMenu = Menu(menubar, tearoff=0)
    alreadyTxtMenu.add_command(label="Select text if already exists", command=MainApplication.choose_already_txt)
    menubar.add_cascade(label="Select Text", menu=alreadyTxtMenu)

    aboutMenu = Menu(menubar, tearoff=0)
    aboutMenu.add_command(label="About this app", command=MainApplication.aboutInfo)
    menubar.add_cascade(label="About", menu=aboutMenu)

    menubar.config(bg="#00072d")
    window.config(menu=menubar)
    MainApplication(window)
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window.configure(fg_color=("#00072d","#00072d"))
    canvas=Canvas()
    # window.mainloop()

    # Label
    # lbl=Label(window, text="---This part is dedicated to convert your files to text file (.txt)---", fg='#ff006e', bg="#00072d", font=("Comic Sans MS", 10, "italic"))
    # lbl.place(x=200, y=10)

    # Button for choosing file
    filename=""
    # file_btn=customtkinter.CTkButton(master= window, text="1. Choose PDF Document", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.choose_file, corner_radius=10)
    # file_btn.place(x=110, y=60)

    # ToolTip(file_btn, msg="Click here to select your PDF file")

    # Button to start pdf to text conversion
    # convert_btn=customtkinter.CTkButton(window, text="2. Convert PDF to Text", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.pdfTotext, corner_radius=10)
    # convert_btn.place(x=110, y=100)

    # ToolTip(convert_btn, msg="Click here to start converting your PDF to txt file")

    # # Button to start Word to text Conversion
    # convertWord_btn=customtkinter.CTkButton(window, text="2. Convert Word to Text", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.wordtoText, corner_radius=10)
    # convertWord_btn.place(x=410, y=60)

    # ToolTip(convertWord_btn, msg="Click here to start converting your Word Document (.Doc and .Docx) to text file")

    # # already_have_txt=""
    # btn_selecttextfile = customtkinter.CTkButton(window, text="2. Select text file",fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.choose_already_txt, corner_radius=10) #Choosing the pre-existing txt file
    # btn_selecttextfile.place(x=410, y=100)

    # ToolTip(btn_selecttextfile, msg="Click here to Select the text file if you already have it and there is no need to convert your file to text")

    options = ["Russian", "Enlish"]

    language = customtkinter.StringVar(value=options[0])
    # language.set(options[0]) # default value

    # Label
    lbl=Label(window, text="--- The following part is dedicated to summarizing your text ---", fg='#ff006e', bg="#00072d",  font=('Comic Sans MS', 10, 'italic'),wraplength=600, justify="center")
    lbl.place(x=200, y=160)

    # Label
    lbl=Label(window, text=" Please select the text language", fg='#ff006e', bg="#00072d",  font=('Comic Sans MS', 10, 'italic'),wraplength=300, justify="center")
    lbl.place(x=410, y=220)

    lang_options = customtkinter.CTkComboBox(master= window, variable=language, values=options, button_hover_color=("#3f37c9","#3f37c9"), dropdown_hover_color=("#3f37c9","#3f37c9"), fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d") )
    # lang_options.config(fg="#3d405b")
    lang_options.place(x=410, y=250)

    btn_language = customtkinter.CTkButton(master= window, text=" I'm happy with this language", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.select_language, corner_radius=10)
    btn_language.place(x=410, y=290)

        #Bind the tooltip with button
    ToolTip(btn_language, msg="Click here to apply your selected file language")

    btn_summary = customtkinter.CTkButton(window, text=" Summary", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.summarize, corner_radius=10)
    btn_summary.place(x=410, y=330)

    ToolTip(btn_summary, msg="Click here to Start summarizing your file")

    # Button for saving the summary
    btn_savefile = customtkinter.CTkButton(window, text = '   Save     ',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = lambda : MainApplication.save(), corner_radius=10)
    btn_savefile.place(x=410, y=370)
    ToolTip(btn_savefile, msg="Click here to save the summary that you created")

    #Button to clear all variables:
    btn_clearVariables = customtkinter.CTkButton(window, text = '   Clear     ',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = MainApplication.clearAfterOneRun, corner_radius=10)
    btn_clearVariables.place(x=410, y=410)
    ToolTip(btn_clearVariables, msg="Click here to clear all variables and start over")

    # Button for closing
    exit_button = customtkinter.CTkButton(window, text="    Exit    ",fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.close_app, corner_radius=10)
    exit_button.place(x=410, y=525)
    # exit_button.place(relx=0.5, rely=0.5, anchor=CENTER)
    ToolTip(btn_savefile, msg="Click here to exit the app")

    window.title('Summarizer, Pdf/word to Text App')
    window.geometry("800x600+10+20")
    #Adding the image to window
    imgPath = "./icons\Bg4.png"
    bgImg = ImageTk.PhotoImage(Image.open(imgPath))
    #The Label widget is a standard Tkinter widget used to display a text or image on the screen.
    bgImagePanel = Label(window, image = bgImg,bg="#00072d")
    #The Pack geometry manager packs widgets in rows or columns.
    # bgImagePanel.pack(side = "bottom", fill = "both", expand = "yes")
    bgImagePanel.place(x=60,y=190)
        # Created by
    account_bitmap = PhotoImage(file = "./icons\exclamation.png") 
    account_bitmap = account_bitmap.subsample(5, 5)
    lbl_exclamation = Label(window , image= account_bitmap, compound= TOP,bg="#00072d")
    lbl_exclamation.place(x=40, y=510)

    lbl_developed=Label(window, text="This app has been developed by ", fg='#ff006e', bg="#00072d",  font=("Comic Sans MS", 9, "italic"))
    lbl_developed.place(x=80, y=500)
    lbl_name=Label(window, text=" amir2628 ", fg='#fcbf49', bg="#00072d", font=("Comic Sans MS", 12, "italic"))
    lbl_name.place(x=80, y=525)

    # My Github profile

    btn_github = customtkinter.CTkButton(window, text = 'My GitHub',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = MainApplication.openGithub, corner_radius=10)
    btn_github.place(x=200, y=525)
    ToolTip(btn_github, msg="Click here to Check my GitHub profile")

    window.protocol("WM_DELETE_WINDOW", MainApplication.on_closing)

    window.mainloop()
