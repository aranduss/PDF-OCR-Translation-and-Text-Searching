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


def fixString(lineOutput):
    #Create Parser using HTML module to convert all named and numeric character references to unicode characters (e.g. '&gt;', '$#62;', etc.)
    parser = HTMLParser.HTMLParser()
    
    #Encode string with 'utf-8' and 'ascii' 
    lineOutput = repr(parser.unescape(lineOutput.decode('utf-8', errors = 'ignore').encode('ascii', errors = 'ignore')))
    
    #Remove line breaks ('\n') and unicode indicators "u'" from string
    lineOutput = lineOutput.replace('\\n', '').replace("u'", '')
    
    #Remove leading and trailing quotes
    lineOutput = lineOutput[1:-1]
    
    return lineOutput

def searchText(text, keyword, searchType):
    #This function is used to search the given text either inclusively or exclusively.
    #An inclusive search would return True for a search for 'a' in 'Another one bites the dust' and 'a fish will swish'
    #An exclusive search would return True only if 'a' was found by itself. For example it would return True for a search of 'a' in 'a fish will swish' but False for 'Another one bites the dust'
    
    if(searchType == 'inclusive'):
        #Inclusive Search
        return text.find(keyword) <> -1
    else:
        #Exclusive Search
        tempTextList = text.split()
        tempBool = False
        for x, key in enumerate(tempTextList):
            if(key == keyword):
                print(key + ' found')
                tempBool = True
                
        return tempBool

def printExcel(counter, sheet1, keyword, pdfList):
#This Function is used to search all text store in pdfList for the given keyword and print the results to sheet1 in the output file. 
#The Counter variable is used to keep track of what line to write on between function calls.
    
    #Determine wether an Inclusive or Exclusive search should be used based on whether the keyword has quotes around it.
    searchType = ''  
    if(keyword.startswith("'") and keyword.endswith("'") or keyword.startswith('"') and keyword.endswith('"')):
        searchType = 'exclusive'
        keyword = keyword[1:-1]        
        print('Modified Keyword = ' + keyword)
        
    else:
        searchType = 'inclusive'
    
    #Loop through PDFs
    for x, key in enumerate(pdfList):
        #Loop through Pages in PDF
        for y, key2 in enumerate(pdfList[key]):
           #Loop through Sentences in Pages
           for z, key3 in enumerate(pdfList[key][y]):                 
               try:
                    #Store page 
                    someText = pdfList[key][y][z].lower()
                    if(searchText(someText, keyword, searchType)):
                                                                           
                        #Write Keyword
                        sheet1.write(counter, 0, keyword)
                        
                        #File Name
                        sheet1.write(counter, 1, key)
                        
                        #Write Page Number
                        text = str(y+1)
                        sheet1.write(counter, 2, text)
                        
                        #Initialize LineOutput variable
                        lineOutput = fixString(pdfList[key][y][z])
                        
                        #Next we bring attention to the keyword in our reporting by using brackets and capitalization.
                        #Note that we are breaking out the sentence string using the Split() command, enumerating through each word inour sentence to find all 
                        # instances of the keyword. Once found, we capitalize using the Upper() command and add brackets [] to make the keywords more easily found 
                        # when reviewing the output.
                        wordList = lineOutput.split()
                        lineOutput = ''
                        for d, key4 in enumerate(wordList):
                            if(searchText(key4.lower(), keyword.lower(), searchType)):
                                newKey4 = '[' + key4.upper() + ']'
                                wordList[d] = newKey4                                
                        
                        #Rejoin wordList with spaces
                        lineOutput = ' '.join(wordList)
                        
                        #Write the sentence containing the keyword found
                        sheet1.write(counter,3, lineOutput)
                        
                        #Increment counter
                        counter = counter + 1
                            
               except Exception as e:
                   print('Error at line ' + str(z+1) + '\n' + str(e))
                   print(someText)                 
               
    return counter


                                            
