# PDF-OCR-Translation-and-Text-Searching
Python script that performs OCR on multiple PDF files, can translate via google translate API, and searches for keywords.

## About
1.	This python project was created to allow for the optical character recognition (OCR) of multiple PDF files with translation and keyword searching capabilities. It was built on and for a Windows PC machine. It uses external software tools for the OCR. ImageMagick is used for converting PDF pages to images. Tesseract is used to run the OCR on the created images. Ghostscript is used as an interpreter for PDF page description languages.

## Prerequisites
1.	ImageMagick v6.9.10-14-Q8-x86-dll
    -	Instructions:	http://docs.wand-py.org/en/latest/guide/install.html#install-imagemagick-on-windows
    -	Download:	https://imagemagick.org/script/download.php

2.	Tesseract w32 v4.0.0.20181030
    -	https://github.com/UB-Mannheim/tesseract/wiki
    -	Note that language selections are chosen during the installation process. 

3.	Ghostscript w32 v9.26
    -	https://www.ghostscript.com/download/gsdnld.html

4.	Python 2.6 (not tested with any other Python versions) and related.
    -	https://www.python.org/download/releases/2.6/
    -	Note: I used the anaconda IDE

5.	Install the following Python Modules
    -	Pyttk 0.3.2
        -	pip install pyttk
    -	Google-cloud-translate 1.3.1
        -	pip install google.cloud
    -	Xlwt 1.3.0
        -	pip install xlwt

6.	Create a Google Service Account key from the Google Cloud Platform.
    -	Here is a link that explains the process:	https://translatepress.com/docs/settings/generate-google-api-key/#createnewproject
    -	Move the created Service Account Key to the same directory as the OCR Python scripts.
 

## Running the Tests
1.	Move the example PDFs to be OCR’ed to a local dedicated folder.
    -	If the PDFs are in different languages, create a folder for each language.
2.	Place the three OCR python scripts into a local folder.
3.	Open and run the OCR_Main.py file from your Python 2.7 IDE
4.	In the user interface that appears:
  -	Either click the ‘Load Keywords’ button to load a text file with your keywords in the correct format or manually enter keywords in the text box.
    -	Keywords should be separated by commas.
    -	By default, searches will be inclusive. To indicate an exclusive search should be used, put single quotes around a keyword.
 	
|Search Type|Keyword	|Would Return|
| -------  |------- |------------| 
|Inclusive |term	  |term, termination, terminal, determined|
|Exclusive |'term'	|term|

  -	Use the dropdown menu to select the language the pdfs are in.
    -	Note that you must have the language libraries correlating to your selection installed in the Tesseract software. This can be done during the Tesseract installation process. If you realize you haven’t installed a language you need, you can simply uninstall and then reinstall Tesseract and select the needed languages during the install process.
    -	There are more languages supported by the Tesseract software than are listed in the drop down menu of this script. Additional 3 letter language codes can be added to the interface’s combo box by editing the code in the OCR_Main.py file on line #80. Remember to also add your language code to either the Germanic or Non-Germanic lists (whichever is appropriate) in the Functions_OCR.py file on lines #154 and #155.
  -	Use the ‘Set Folder’ button to select the directory to run in (i.e. the folder containing your PDF files).
  -	Ensure that none of the PDF files to be processed are open.
  -	Click run.
5.	Windows command consoles will appear throughout the running the program. Wait until all consoles have disappeared and an Excel file named ‘Keyword Results’ has appeared in the folder you selected. This file will contain the results of your keyword searches, indicating each instance your keyword was found, the file and page number it was found in, and the sentence it was found in.
  -	Also note that each of your PDF files will have been moved to a dedicated folder that will now contain:
    -	The original file
    -	A text file named after the original file containing all of the OCR’ed text from that file. It will contain the translated text if you indicated the PDF files were in a different language.
      -	If you wish to browse the converted text or to perform manual searching it will probably be easiest to use this text file.
        -	An image of each page from the file. Note that the page count begins with 0 (i.e. 0,1,2,3,4,5…).
        -	A text file containing text for each individual page. Note that the page count begins with 0.

## Author
•	Alex Anduss – All python scripts
