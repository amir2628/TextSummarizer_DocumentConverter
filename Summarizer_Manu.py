from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import *
# from tkPDFViewer import tkPDFViewer as pdf
import pathlib, os
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
import PyPDF2

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

    def choose_already_txt(): #Choosing the PDF file
        Tk().withdraw()
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

        Tk().withdraw()
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
                PDF2TextFile.close()
        mbox.showinfo(message= "File has been saved in Path: '"+ str(PDF2TextFile))
        global fullPdfPath
        fullPdfPath = PDF2TextFile


    def select_language(): #choosing the pdf language
        global languageis
        languageis = ""
        global language
        languageis = language.get()
        mbox.showinfo(message= "You have chosen "+str(languageis))

    def select_percentage():
        global perc
        perc=""
        global percent
        perc=percent.get()
        mbox.showinfo(message= "You have chosen "+str(perc))
    
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
        global percent
        percent = ""
        global perc
        perc=""
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
        per=(int(perc))/100
        print(per)
        if languageis == 'Russian':
            nlp = spacy.load('ru_core_news_lg')
        else:
            nlp = spacy.load('en_core_web_trf')
        if languageis =='Russian':
                if txtAlreadyDir !="":
                        with open(text, encoding='utf-8', mode = "r") as te:
                            print("Russian Loop: "+str(te))
                            article = te.read()
                            # print(article)
                else:
                        with open(text.name, encoding='utf-8', mode = "r") as te:
                            print("Russian Loop: "+str(te))
                            article = te.read()
                            # print(article)
        else:
                if txtAlreadyDir !="":
                    with open(text, encoding='utf-8', mode = "r") as te:
                        print("Russian Loop: "+str(te))
                        article = te.read()
                        # print(article)
                else:
                        if text == str:
                            with open(text, encoding='utf-8', mode = "r") as te:
                                print("Russian Loop: "+str(te))
                                article = te.read()
                                # print(article)
                        else:
                            with open(text.name, encoding='utf-8', mode = "r") as te:
                                print("Russian Loop: "+str(te))
                                article = te.read()
                                # print(article)
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
        print(len(sentence_tokens))
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
        file = asksaveasfile(mode = 'wb', title="Select Location", filetypes=[('txt file', '*.txt')])
        if file is None:
            return
        else:
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

    def helpDoc():
        helpWindow = customtkinter.CTk()
        # customtkinter.CTkFrame(window)
        HelpPDFDir = os.fspath(pathlib.Path(__file__).parent / 'Help.pdf')
        HelpPDFFile = PyPDF2.PdfReader(HelpPDFDir)
        Helppages = HelpPDFFile.pages[0]
        HelpPDFtext= Helppages.extract_text()
        helptextbox = customtkinter.CTkTextbox(helpWindow, width=1000, height=600, scrollbar_button_color=("#ff006e","#ff006e"))
        helptextbox.grid(row=0, column=0)
        helptextbox.insert("0.0", HelpPDFtext)  # insert at line 0 character 0
        helptextbox.configure(state="disabled")  # configure textbox to be read-only
        helptextbox.pack()
        helpWindow.mainloop()


if __name__ == "__main__":
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
    aboutMenu.add_command(label="Help", command=MainApplication.helpDoc)
    menubar.add_cascade(label="About", menu=aboutMenu)

    menubar.config(bg="#00072d")
    window.config(menu=menubar)
    MainApplication(window)
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")
    window.configure(fg_color=("#00072d","#00072d"))
    canvas=Canvas()
    filename=""
    options = ["Russian", "Enlish"]

    language = customtkinter.StringVar(value=options[0])

    # Label
    lbl=Label(window, text="--- The following part is dedicated to summarizing your text ---", fg='#ff006e', bg="#00072d",  font=('Comic Sans MS', 10, 'italic'),wraplength=600, justify="center")
    lbl.place(x=200, y=60)

    # Label
    lbl=Label(window, text=" Please select the text language", fg='#ff006e', bg="#00072d",  font=('Comic Sans MS', 10, 'italic'),wraplength=300, justify="center")
    lbl.place(x=350, y=100)

    lang_options = customtkinter.CTkComboBox(master= window, variable=language, values=options, button_hover_color=("#3f37c9","#3f37c9"), dropdown_hover_color=("#3f37c9","#3f37c9"), fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d") )
    lang_options.place(x=340, y=130)

    btn_language = customtkinter.CTkButton(master= window, text="✔️", width=30,
                                 fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.select_language, corner_radius=10)
    btn_language.place(x=520, y=130)
        #Bind the tooltip with button
    ToolTip(btn_language, msg="Click here to apply your selected file language")

    percentOptions = ["10", "15", "25", "40", "50", "60", "70", "80", "90"]
    percent = customtkinter.StringVar(value=percentOptions[4])
    percent_options = customtkinter.CTkComboBox(master= window, variable=percent, values=percentOptions, button_hover_color=("#3f37c9","#3f37c9"), dropdown_hover_color=("#3f37c9","#3f37c9"), fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d") )
    percent_options.place(x=340, y=170)


    btn_percent = customtkinter.CTkButton(master= window, text="✔️", width=30,
                                 fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.select_percentage, corner_radius=10)
    btn_percent.place(x=520, y=170)
        #Bind the tooltip with button
    ToolTip(btn_percent, msg="Click here to apply your selected summary percentage")

    btn_summary = customtkinter.CTkButton(window, text=" Summary", fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.summarize, corner_radius=10)
    btn_summary.place(x=410, y=210)
    ToolTip(btn_summary, msg="Click here to Start summarizing your file")

    # Button for saving the summary
    btn_savefile = customtkinter.CTkButton(window, text = '   Save     ',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = lambda : MainApplication.save(), corner_radius=10)
    btn_savefile.place(x=410, y=250)
    ToolTip(btn_savefile, msg="Click here to save the summary that you created")

    #Button to clear all variables:
    btn_clearVariables = customtkinter.CTkButton(window, text = '   Clear     ',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = MainApplication.clearAfterOneRun, corner_radius=10)
    btn_clearVariables.place(x=410, y=290)
    ToolTip(btn_clearVariables, msg="Click here to clear all variables and start over")

    # Button for closing
    exit_button = customtkinter.CTkButton(window, text="    Exit    ",fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command=MainApplication.close_app, corner_radius=10)
    exit_button.place(x=410, y=365)
    ToolTip(btn_savefile, msg="Click here to exit the app")

    window.title('Summarizer, Pdf/word to Text App')
    window.geometry("700x500+10+20")
    #Adding the image to window
    imgPath = "./icons\Bg4.png"
    bgImg = ImageTk.PhotoImage(Image.open(imgPath))
    #The Label widget is a standard Tkinter widget used to display a text or image on the screen.
    bgImagePanel = Label(window, image = bgImg,bg="#00072d")
    bgImagePanel.place(x=60,y=90)
        # Created by
    account_bitmap = PhotoImage(file = "./icons\exclamation.png") 
    account_bitmap = account_bitmap.subsample(5, 5)
    lbl_exclamation = Label(window , image= account_bitmap, compound= TOP,bg="#00072d")
    lbl_exclamation.place(x=40, y=350)

    lbl_developed=Label(window, text="This app has been developed by ", fg='#ff006e', bg="#00072d",  font=("Comic Sans MS", 9, "italic"))
    lbl_developed.place(x=80, y=340)
    lbl_name=Label(window, text=" amir2628 ", fg='#fcbf49', bg="#00072d", font=("Comic Sans MS", 12, "italic"))
    lbl_name.place(x=80, y=365)

    # My Github profile

    btn_github = customtkinter.CTkButton(window, text = 'My GitHub',fg_color=("#3d405b","#3d405b"), bg_color=("#00072d","#00072d"), command = MainApplication.openGithub, corner_radius=10)
    btn_github.place(x=200, y=365)
    ToolTip(btn_github, msg="Click here to Check my GitHub profile")

    window.protocol("WM_DELETE_WINDOW", MainApplication.on_closing)

    window.mainloop()
