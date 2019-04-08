# -*- coding: utf-8 -*-
"""
@author: Alex Anduss
Created on Wed Dec 05 20:22:02 2018
Refer to ReadMe for detailed install and run instructions.
"""
import Tkinter
from Tkinter import Tk, Label, Button, Entry, RIGHT, LEFT, BOTH, RAISED, Frame, E,W,N,S
import ttk
import tkFileDialog
from Functions_OCR import *
from Functions_Support import *

####################################################################################################################################################
#Creating a simple user interface using the Tkinter module for users to run the program.
#  This interface will allow the user to:
#    1. Understand how to run the program,
#    2. Specify the language to use,
#    2. Specify the folder to run OCR in,
#    3. Specify any keywords to search for in the OCR'ed text
####################################################################################################################################################
                        
class GUI:
    Path = "" 
    
    ########################################################################################
    # Define functions to be used for the GUI to work
    #####################################################   
    
    def greet(self):
        
        #If the folder to run in has not been assigned, force the folder selection before running.
        if(GUI.Path == ""):
           GUI.Path = tkFileDialog.askdirectory()
        language = self.combo.get()
                
        #Pull the keyword text from the Entry widget.
        text = self.entry.get()
        
        #Create a list from the string provided delimited by commas.
        keywordList = text.split(',')        
        
        #Output keywords to console for debugging
        for x, key in enumerate(keywordList):
            keywordList[x] = key.lower().replace(' ','')
            print(str(x) + " - " + key)
        
        #Run the main function and pass the path, language, and keyword list.
        run_OCR(GUI.Path, language, keywordList)
    
    def load(self):
        #Open text file for reading. 
        keywordFile = tkFileDialog.askopenfile()
        keywords = keywordFile.read()
        
        #print file contents for debugging.
        print(keywordFile.readlines())
        
        self.entry.insert(0,keywords)
        
            
    def choose(self):
        #Assing the folder the user chooses to the GUI.Path variable
        GUI.Path = tkFileDialog.askdirectory()
        
    def __init__(self, master):
        self.master = master       
        master.title("PDF OCR with Python")
        
        ##########################################################################
        # Create buttons and our combobox for user interaction
        #########################################################
               
        #Create button to load a list of keywords
        self.load_button = Button(master, text="Load Keywords", command=self.load, width = 12)
        self.load_button.grid(row = 2, column = 0)
        
        #Create Combo Box        
        self.combo = ttk.Combobox(width = 10)
        self.combo['values'] = ('eng','spa', 'chi_tra', 'ita', 'jpn', 'heb', 'kor', 'fra', 'deu', 'por', 'pol', 'rus')        
        self.combo.current(0)
        self.combo.pack(side=LEFT)
        self.combo.grid(row =4, column = 0)
        
        #Create a folder chooser
        self.chooser_button = Button(master, text="Set Folder", command=self.choose, width = 12)
        self.chooser_button.grid(row = 6, column = 0)
                
        #Create Run Button
        self.greet_button = Button(master, text="Run", command=self.greet, width = 12)
        self.greet_button.grid(row=8, column = 0)
                        
        ##########################################################################
        # Create Middle Labels for Padding
        ########################################
        
        # Create blank label for padding
        self.label = Label(master, text="      ")
        self.label.grid(row=1, column = 1, sticky = W)      
        
        self.label = Label(master, text="      ")
        self.label.grid(row=2, column = 1, sticky = W)
        
        self.label = Label(master, text="      ")
        self.label.grid(row=3, column = 1, sticky = W)
        
        self.label = Label(master, text="      ")
        self.label.grid(row=4, column = 1, sticky = W)
        
        self.label = Label(master, text="      ")
        self.label.grid(row=5, column = 1, sticky = W)
             
        ##########################################################################
        # Create Instruction Label(s)
        ######################################
        self.label = Label(master, text="Run Instructions:")
        self.label.grid(row=1, column = 3, sticky = W)
        
        self.label = Label(master, text="1. Place all PDFs for OCR into a folder.")
        self.label.grid(row=2, column = 3, sticky = W)
                
        self.label = Label(master, text="   1a. If PDFs are in different languages then create a folder for each language.")
        self.label.grid(row=3, column = 3, sticky = W)
        
        self.label = Label(master, text="2. Either load or enter the keywords in the text box below.")
        self.label.grid(row=4, column = 3, sticky = W)
        
        self.label = Label(master, text="   2a. Seperate words by commas. Use quotes for exact word searches")
        self.label.grid(row=5, column = 3, sticky = W)
        
        self.label = Label(master, text="3. Specify the language of the PDFs.")
        self.label.grid(row=6, column = 3, sticky = W)
        
        self.label = Label(master, text="4. Set the folder to run in.")
        self.label.grid(row=7, column = 3, sticky = W)
        
        self.label = Label(master, text="5. Click Run.")
        self.label.grid(row=8, column = 3, sticky = W)
        
        self.label = Label(master, text="6. A folder will be created for each PDF containing the original file and the OCR'ed text .")
        self.label.grid(row=8, column = 3, sticky = W)
                
        self.label = Label(master, text="7. A results excel file ('Keyword Results') will be created in your original folder.")
        self.label.grid(row=9, column = 3, sticky = W)
        
        self.label = Label(master, text="")
        self.label.grid(row=10, column = 3, sticky = W)
        
        ###########################################################################        
        #Create an Entry widget to enter/display keywords at bottom
        ############################################################
        
        self.label = Label(master, text="Keywords:")
        self.label.grid(row=11, column = 0)
        
        self.entry = Entry(master)
        self.entry.grid(row = 12, column = 1, columnspan = 6, sticky=W+E)
        
        self.label = Label(master, text="Example: privacy, 'term', fine, penalt, cost, exception")
        self.label.grid(row=11, column = 1, columnspan = 6, sticky = W)        
        ###########################################################################

# Create an instance of the GUI class, set window size, and set to run on execution
root = Tkinter.Tk()
root.geometry("600x300")
my_gui = GUI(root)
root.mainloop()
