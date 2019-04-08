# -*- coding: utf-8 -*-
"""
@author: Alex Anduss
Created on Wed Dec 05 20:22:02 2018
Refer to ReadMe for detailed install and run instructions.
"""

import os
from fnmatch import fnmatch
import shutil
from collections import defaultdict
from google.cloud import translate
import io
from xlwt import Workbook
import HTMLParser
from time import sleep
from Functions_Support import *

def run_OCR(path, language, keywordList):   

        #Create a variable, logPath, to preserve the original directory path     
        logPath = path
        
        #Initialize a Dictionary that will contain the PDF text for all pages of all files
        pdfList = defaultdict(list)
        
        #Set Root folder
        root = path 
        
        #only target PDF Files
        pattern = "*.pdf"
        
        #Make current directory root
        os.chdir(root)
    
        #Cycle through each pdf file in the path directory, create a new folder for that file, move the file to its new folder, and run OCR on it.
        #Translation will occur if the language specified is not English using Google Translate API.        
        for path, subdirs, files in os.walk(root):
            for name in files:
                #if the file is a pdf (i.e. matches the pattern) then OCR it
                if fnmatch(name, pattern):
                    
                    #Reset current directory at the root folder
                    os.chdir(root)
                    
                    #Get the filename without extension (i.e. '.pdf')
                    txt = os.path.splitext(name)[0]
                    
                    #If a folder with the name of the current pdf file does not already exist, continue with OCR. 
                    #If a Folder already exists then this would indicate that OCR has already been run for this file and it will attempt to run
                    #on the next file.
                    if not os.path.exists(txt):
                        
                        #Make a directory (i.e. folder) for the current pdf file.
                        os.mkdir(txt)
                        
                        #Go to the directory just made.
                        newDir = root + '\\%s' % txt
                        os.chdir(newDir)
                        
                        #Set starting and ending paths for file movement.
                        fromPath = root + '\\' + name
                        toPath = root + '\\' + txt + '\\' + name
                        
                        #Move the pdf file to their respective folders for further analysis.
                        shutil.move(fromPath, toPath)
                        
                        #Force quotes around name in case filename contains spaces.
                        tmpName = '"' + name + '"'
                        
                        #############################################################################################################
                        # Convert PDF To Image via ImageMagick Tool
                        #############################################################################################################
                        
                        #Convert each page of the current pdf file to a png file with a density (i.e. quality) of 300.
                        #A density of 300 seems to be what most online articles recommend for OCR via the ImageMagick tool. 
                        #This value can be tweaked if OCR performance is not optimal.
                        #Also note that the ImageMagick tool is a separate software installed on the Windows machine that is being accessed via the Command Line below using the OS Python module.
                        #ImageMagick commands begin with 'convert'. The ImageMagick tool must be installed on your windows machine for this code to function properly.
                        try:
                            os.system('convert -density 300 %s -colorspace Gray -alpha off page.png' % tmpName)
                        except Exception as e:
                            print('Error - ' + str(e) + '\n' + 'Ensure you have the ImageMagick tool installed on your windows machine and the MAGICK_HOME system variable set in your Windows Environment Variables')
                            
                        #Only target png Files in the current folder (i.e. The images of each pdf page that were just created).
                        pngPattern = "*.png"                
                        
                        #Create a temporary path pointing to the new folder created for the current pdf file.
                        tmpPath = root + '\\' + txt
                        
                        #############################################################################################################
                        # OCR via Tesseract
                        #############################################################################################################
                        #Run OCR on each page's .png file using the Tesseract tool. 
                        #This will result in a text file for each png file. Note that the naming convention will be the same but with a different extension. 
                        for path, subdirs, files in os.walk(tmpPath):
                            for pngName in files:                        
                                if fnmatch(pngName, pngPattern):
                                    try:
                                        #Create a string containing the windows command line command to run OCR via the Tesseract tool. The os.path.splitext command is used to ignore the '.png' extension.
                                        tesseractCMD = 'Tesseract {} {} -l {}'.format(pngName, os.path.splitext(pngName)[0], language)
                                        
                                        #Run the command created above via the windows command line
                                        os.system(tesseractCMD)
                                        
                                    except Exception as e:
                                        print('Error - ' + str(e) + '\n' + 'Ensure you have the tesseract and ghostscript tools installed on your windows machine')
                                                
                        #only target the text files (again, these each correspond to a page from the current pdf file) created by the Tesseract tool.
                        txtPattern = "*.txt"   
                       
                        #Create a list to hold each page of text and the name of each file. This will make it easy to perform keyword searching later on.
                        filenames = []
                        pageList = []
                                                
                        #Iterate through each text file and populate the filenames and pageList lists with each page's data.
                        for path, subdirs, files in os.walk(tmpPath):
                            for txtName in files:
                                if fnmatch(txtName, txtPattern):
                                    tmpFile = open(txtName, "r").readlines()
                                    pageList.append(tmpFile)
                                    filenames.append(txtName)
                                    
                        #Populate dictionary to keep track of the text on each page of each PDF.
                        #To print a page: pdfList["PDF_Name"][0][1]
                        pdfList[txt] = pageList                      
                                                
                        #The following code will output all OCR'ed data to one master text file named after the original PDF file. 
                        #If the specified language was not english, the text will be translated before being dumped in the master file.
                        
                        #Set path and name of master file.
                        outputPath = tmpPath + '\\%s.txt' % txt
                        
                        #Create/Open the Master file (referenced as 'outfile' in the code) for writing.
                        with open(outputPath, 'w') as outfile:
                            
                            # counter used for keeping track of what page we are on.
                            count = 1 
                            
                            #Iterate through each text file (corresponing to pages) created by the Tesseract tool in current directory.
                            for fname in filenames:    
                                
                                #Open current text file for reading. 
                                with io.open(fname, mode='r', encoding = 'utf-8') as infile: 
                                    
                                    #Assign page text to the 'paragraph' variable. Note that each page of the PDF was converted to separate 
                                    #  image files so each text file read in this loop corresponds to a page from the pdf.
                                    paragraph = infile.read() 
                                    
                                    #Set Encoding Type based on language being read in. Germanic-based languages seem to work only with 'ascii' while 
                                    #  non-germanic languages seem to do best with 'utf-8'.
                                    
                                    #Below are two lists indicating how the different languages are classified for reference when picking an encoding type.
                                    Non_Germanic_Langs = 'kor', 'jpn', 'chi_tra','rus'
                                    Germanic_Langs = 'eng','fra', 'ger', 'spa', 'por','ita', 'heb', 'deu', 'pol'
                                    
                                    #Choose an encoding type to use on text before sending it off to the google cloud for translation.
                                    if language in Germanic_Langs:
                                        encodingType = 'ascii'
                                    else:
                                        encodingType = 'utf-8'                                    
                                    
                                    #If the language was english ('eng') then no translation is needed so lets not waste our google credits.
                                    if(language == 'eng'):
                                        #write the page text straight to the master file
                                        outfile.write(paragraph.encode(encodingType, errors = 'ignore'))
                                    else:                                         
                                        
                                        # Instantiates a Google Translate client
                                        # NOTE: Your Google translate API key must be in the same folder as your script for this next line to run.
                                        try:
                                            translate_client = translate.Client()
                                        except Exception as e:
                                            print("Error: Ensure that your Google Translate API Key is in the same folder as your script")
                                        
                                        # Translates text to English ('en') using Google Translate's API. Note that the encoding type varies between 'utf-8' and 'ascii' depending on the language being read in.
                                        print("sending translation request")
                                        try:
                                            ##################################################################################################################
                                            # Translate Text Via Google Cloud Translation Services
                                            #################################################################################################################
                                            
                                            #Set target language
                                            target = 'en' #English
                                                                                       
                                            #Create a variable ('translation') with Google's translated text. 
                                            #Note that in the google translate query we provide the page text encoded with our chosen encoding type and say to ignore any pesky errors
                                            # that may pop up from unrecognized characters or punctuation. 
                                            # No reason to waste all our effort because it couldn't figure out what an apostrophe was. 
                                            #
                                            #Also note that the target language (i.e. convert everything to english) is provided. If you wanted to have all the text converted to 
                                            # another language, then you would want to change "target = 'en'" to a language code corresponding to your preffered language.
                                            #
                                            #You would also want to change the previous IF statement above to your preferred language.
                                            #  - Google's language keys: https://cloud.google.com/translate/docs/languages
                                            #  - Tesseract's language keys ('639 - 2/t' column): https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes
                                            #
                                            #Also note that you must have the Tesseract Language Library installed for Tesseract to use it. 
                                            # Languages libraries to be installed are specified during the Tesseract installation.
                                            # You can check which libraries you have installed by opening the windows console and typing 'tesseract --list-langs'
                                            
                                            #The below block of code attempts to retreive the translated paragraph text. If there is an error then it will
                                            #sleep for 5 seconds and try again to a maximum of 5 attempts. Such errors could include internet connectivity issues or bad handshakes during the communication.
                                                                                        
                                            translation = ''
                                            counter = 0
                                            while ((translation == 'Error' or translation == '') and (counter < 5)):
                                                counter += 1
                                                translation = translateText(paragraph, encodingType, target, translate_client)
                                                
                                                #If there was an error then wait 5 seconds and then the loop will try to request a translation form the google cloud translate server again.
                                                if(translation ==  'Error'):
                                                    print('Sleeping: ' + str(counter))                                                    
                                                    sleep(5)
                                                                                        
                                            #Get only the translated text from Google's returned information
                                            translatedText = translation['translatedText']
                                                                                        
                                            #Split the translated text returned by Google by periods so that each sentence can be written on a new line.
                                            pageList = translatedText.split(". ")
                                            
                                            #################################################################################################################
                                            # Print out the translated text to the master file. 
                                            #################################################################################################################
                                            
                                            #Also, Add the period back to each sentence and skip to next line
                                            for x, sentence in enumerate(pageList):
                                                #Add a period to the end of each sentence except the last because it will already have one.
                                                if(x < len(pageList)-1):
                                                    sentence = sentence + '.'
                                                
                                                #Write out sentence to the Master File
                                                outfile.write(sentence + '\n')
                                                
                                                #Populate the pdfList list with the sentences for each page. This list will be used for keyword searching.
                                                pdfList[txt][count-1][x] = sentence
                                            
                                        except Exception as e:
                                            print("Error: " + '\n' + str(e))
                                
                                #Print an end of page marker to make it easer to compare the OCR'ed text to the original pdf document.
                                outfile.write("\n")
                                outfile.write("End of Page %s \n" % count)
                                outfile.write("___________________________________________________________\n")
                                outfile.write("\n")            
                                count += 1
                                
        ###########################################################################################################################
        ## Search for Keywords and print results to excel file
        ###########################################################################################################################
        
        # Creating excel Document for outputing keyword match results and assign working sheet to use.
        wb = Workbook()
        sheet1 = wb.add_sheet('sheet1')
        
        #Creating Report Headers.
        sheet1.write(0,0, 'Keyword')
        sheet1.write(0,1, 'File Name')
        sheet1.write(0,2, 'Page #')
        sheet1.write(0,3, 'Sample of text with Keyword')
        
        #Loop through the list of keywords and print matches to the results excel file initialized above.
        #  The counter variable is used to keep track of the row number to print results to in Excel.
        counter = 1                
        for t, keyword in enumerate(keywordList):
            keyword = keyword.strip()
            counter = printExcel(counter, sheet1, keyword.lower(), pdfList)
        
        #Saving output file (Keyword Results) to the chosen folder.
        logPath = logPath + '/Keyword Results.xls'
        print('Saving Log to ' + logPath)
        wb.save(logPath)
        ##############################################################################################################################

def translateText(paragraph, encodingType, target, translate_client):
    #This function attempts to send the paragraph text to google cloud translate and receive the translated text which is stored in the 'translation' variable
    
    try:
        translation = translate_client.translate(paragraph.encode(encodingType, errors = 'ignore'), target_language = target)
    except Exception as e:
        translation = "Error"
        print("Err: " + str(e))
    
    return translation
                                            
