     _____                 _     _____           _     _               
    |  __ \               | |   |  __ \         | |   | |              
    | |  \/_ __ __ _ _ __ | |__ | |  \/_ __ __ _| |__ | |__   ___ _ __ 
    | | __| '__/ _` | '_ \| '_ \| | __| '__/ _` | '_ \| '_ \ / _ \ '__|
    | |_\ \ | | (_| | |_) | | | | |_\ \ | | (_| | |_) | |_) |  __/ |   
     \____/_|  \__,_| .__/|_| |_|\____/_|  \__,_|_.__/|_.__/ \___|_|   
                    | |                                                
                    |_|          Laura Moore (LMOORE12) - 2022         

## Description

GraphGrabber is a JLR tool which formats PDF reports into a PowerPoint style. It does this by converting PDFs to images then cropping them. The resultant image is then added to a slide alongside other images found in the PDF.

The report definitions are hardcoded for now and will remain that way unless there is a user need to build new functions for cropping and placing images.

Currently supports VT-01 (2 images), VT-07 (7 images), VT-12 (2 or 4 images), VT-15 (3 or 6 images)

Contains a copy of Pmw.py due to Python Megawidgets failing to import properly.

**Tkinter** is used to build the GUI, **PyMuPDF** handles PDF documents alongside **Pillow**, and **python-pptx** is used for operations on the PowerPoint files.

Builds with Pyinstaller and runs in Python 3.8.5 on:\
Win11\
Win10\
MacOS Big Sur (M1 Native)

-------------------------------------------------------

## User Guide

Creating a new deck is simple, but can break if you give it the wrong files. Here's a simple set of instructions:

### Basic Guide

1. Run GraphGrabber in its own folder.
2. Click Yes to generate required folders
3. Click Check Files to verify folders exist and are empty.
4. Place PDF reports into their respective folders. (Use Open Working Directory to visit in a file explorer).
5. Click Create Deck!
6. Output file will be found in the working directory with a Unix timestamp
7. Logfile is generated as GG.log in the working directory.
8. You can clear out all files from GraphGrabber folders with the Clear Folders button.

### Autosort

1. First select whether you want to place VT-12 CE reports into the Single or Three Phase folder (Default Single).
2. Click Autosort
3. Select the folder containing PDF reports
4. Submit
5. Autosort will attempt to place files into the right folders depending on their file name. Please check them!
6. Unsorted files will go into the Unsorted PDFs folder.

-------------------------------------------------------

## Building from source

Using [pyinstaller](https://pyinstaller.org/en/stable/usage.html) we can easily generate from source:

    pyinstaller --clean -F -y -n "GraphGrabber" --add-data="emptyDeck.pptx;." GraphGrabber.py

Running from source may require manual installation of python-pptx and PyMuPDF:

    pip install --upgrade python-pptx PyMuPDF

---

## Guide to Functions etc

**posDict**

Coordinates that determine where an image will be placed on the slide. Not pixels\
Tuple Order: Dist from Left , Dist from Top, Img Width, Img Height

**Example:** *"VT07MW": (Pt(1), Pt(70), Pt(233), Pt(176))*

---

**cropDict**

Similar to posDict but this one determines where in the PDF image that GraphGrabber crops a graph. There is an extra hardcoded value for certain magnetic plots which generate differently for some reason.

**Example:** *"upperOld": ((130, 138, 1000, 800))*

---

**nameDict**

A dictionary for converting from function name to full name for the report.

**Example:** *"VT12Single": "VT-12 Conducted Emissions (Single Phase)"*

---

**searchReplace(search_str, repl_str, input, output)**

Searches for a string within a text frame, within a shape, on any pptx slide. Replaces search_str with repl_str and saves the file.

**search_str (string):** The string to search for in the PowerPoint\
**repl_str (string):** The string to replace search_str with\
**input (string):** The filename to work on\
**output (string):** The filename to save as. Can be the same as input

**Example:** *searchReplace('Old text', 'New text!', 'newDeck.pptx', 'newDeck.pptx')*

---

**extractImages(PDFName, imageFolder)**

Opens the file specified as doc, sets up a zoom in using a 2x2 matrix. For each page in doc it loads the page then runs page.get_pixmap(matrix) to build the raster image. This image is appended to the extractedImages list for later.

**PDFName (string):** The name of the PDF to open and extract images from \
**imageFolder (string):** Name of the folder in which the PDF lives. Uses the current working directory.

**Example:** *extractImages('State 5 - Test Report.pdf', 'VT-01 3m')*

---

**cropGraph(targetImg, cropTuple, imName)**

Crops a single image file, targetImg, using the coordinates defined in cropTuple and adds to the croppedImages list. imName currently only used for tracking which image was cropped (likely a poor idea in hindsight).

**targetImg (Image Object):** The image to be cropped. In most cases this will be an image of a PDF page\
**cropTuple (Tuple):** Pixel coordinates to crop (Left Start, Top Start, Left End, Top End)\
**imName (string):** String to hold a name for the image. Only really used in logging

**Example:** *cropGraph(extractedImages[1], (130, 138, 1000, 800), 'Anything')*

---

**insertImage(oldFileName, newFileName, img, positionTuple, slideNumber)**

Inserts an image from the croppedImages array into the chosen slide of oldFileName using a position tuple. Saves as newFileName. Uses BytesIO() to make img mimic a real PNG file. Adds images with the python-pptx function: *slide.shapes.add_picture(temp, left, top, width, height)*

**oldFileName (string):** The pptx file name to insert into\
**newFileName (string):** The name your PowerPoint will be saved as\
**img (Image Object):** The image object to insert\
**positionTuple (Tuple):** A tuple of PowerPoint specific dimensions in the format (Dist from Left, Dist from Top, Img Width, Img Height)\
**slideNumber (int):** The slide number that is being worked on

**Example:** *insertImage(newDeck.pptx, newDeck.pptx, croppedImages[0], (Pt(1), Pt(70), Pt(233), Pt(176)), 1)*

---

**initialisePowerPoint(emptyDeckName, newDeckName)**

Will make a copy of a PowerPoint (emptyDeck.pptx usually) to serve as the template deck. Saves as newDeckName. If using a single-file executable it will check sys._MEIPASS to find the template in there instead. Uses the current working directory.

**emptyDeckName (string):** Name of the original deck you want to use as a template (no extension)\
**newDeckName (string):** Name of the new PowerPoint deck

**Example:** *initialisePowerPoint('emptyDeck', 'myNewFile')*

---

**reportFunction(PDFName, folderName, slideNumber, deckName)**

A set of functions such as VT12Single() which call the following functions in order to extract, crop, and insert all required images for a type of report.

**extractImages()** --> **cropGraph()** --> **insertImage()** --> Clear extractedImages & croppedImages lists

**reportFunction (function):** The function you want to call\
**PDFName (string):** The name of the PDF to work on\
**folderName (string):** The name of the folder your PDF is in\
**slideNumber (int):** The slide to work on\
**deckName (string):** The PowerPoint file name

**Example:** *VT12Single('State 5 - Test Report.pdf', 'VT-12 Single Phase', 17, 'myOutputDeck')*

    def VT12Single(PDFName, folderName, slideNumber, deckName):
        extractImages(PDFName, folderName)
        cropGraph(extractedImages[1], cropDict["upperOld"], "VT12SingleL1")
        cropGraph(extractedImages[1], cropDict["lowerOld"], "VT12SingleN")
        deckName = deckName + ".pptx"
        insertImage(deckName, deckName, croppedImages[0], posDict["VT12SingleL1"], slideNumber)
        insertImage(deckName, deckName, croppedImages[1], posDict["VT12SingleN"], slideNumber)
        extractedImages.clear()
        croppedImages.clear()
        logging.info(">>>>>>>>>>>> Finished VT-12 Single Phase for " + PDFName)

---

**setSlideCounter(num)**

Sets the current working slide (slidecounter) to num. Uses its own function to avoid constantly asking for the global slideCounter.

**num (int):** Integer to set the slidecounter to

**Example:** *setSlideCounter(5)*

---

**loopFolder(folderName, deckName, reportFunction)**

For any chosen FolderName, loopFolder() will loop through all the PDF files in the directory. For each file it will run **reportFunction()** --> **searchReplace()** before incrementing the slideCounter and moving on to the next file.

It updates a statusMessage as it goes, appending either the current slide or the Exception (e) to the logfile and the fileList GUI widget.

**folderName (string):** The name of the folder you wish to loop through\
**deckName (string):** The name of the PowerPoint deck you are working on\
**reportFunction (Function):** The function you are going to call to extract, crop, and insert images

**Example:** *loopFolder('VT-12 Single Phase', 'myDeckName', VT12Single)*

---

**getListboxValue()**

Returns the currently selected item in the GUI widget: **fileList**

---

**btnInitialisePowerPoint()**

Deprecated function for debug only right now. Runs InitialisePowerPoint('emptyDeck','newDeck')

---

**btnInitialiseFolders()**

Checks if the needed directories exist. If any are missing it will create them with os.mkdir(). Creates the following:

*VT-01 3m,\
VT-07,\
VT-12 Single Phase,\
VT-12 Three Phase,\
VT-15 Electric,\
VT-15 Magnetic,\
Unsorted PDFs*

---

**checkFolders()**

This is a folder check that is ran in multiple places including on startup and if a user tries to click a button that needs GraphGrabber's folders. It throws up a message asking the user whether they want to create folders or not.

---

**askForOutput()**

Simply a popup that will ask the user to enter their desired file name. Will remove special characters with the following regex expression:

    '[^A-Za-z0-9 ]+'

---

**btnClearFolders()**

Deletes every file using os.remove() in the folders created by btnInitialiseFolders(). Will not ask for confirmation. Deleted files are logged in GG.log

---

**btnVisitFolders()**

Uses os.getcwd() to open the current working directory. Method differs by OS:

*Windows: os.startfile(path)\
MacOS: subprocess.Popen(["open", path])\
Linux: subprocess.Popen(["xdg-open", path])*

---

**getInputBoxValue()**

Tkinter function for getting the user input text box. Returns userInput as a string

---

**btnCheckFiles()**

A surprisingly complex function that creates its own function: **loopInsertList(dir)** where dir is the directory you want to check the files in.

Will clear fileList and set listCounter to 0 before running **loopInsertList()** for each dir.

**loopInsertList()** will loop through each file and add it to the fileList GUI object. Between sections it will insert asterixes with the directory name. localCounter and listCounter are two integers used here. localCounter will count the files per directory and listCounter keeps track of the total number of files checked so far. This is used to properly find the locations of files and headers so that they can be deleted or edited to update the user.

**Example:** *loopInsertList("VT-07")*

Finally sets the progressBar value to 0 and maximum to the total file count (currently hardcoded as listCounter-6).

---

**btnGO()**

Will check that the folders exist, display files to user, initialise the PowerPoint, then loop through each folder. It uses the following expression append the current Unix timestamp onto the end of the file name:

    outputFileName = f"{outputFileName} {time.time():.0f}

---

**btnAutoSort()**

The auto sorter will user regex to try and recognise specific patterns in JLR reports. A file containing **BB** for example is likely to be a VT-01 report. VT-07 reports have no way of being sorted.

*Anything unsorted is added to the Usorted PDFs folder*
