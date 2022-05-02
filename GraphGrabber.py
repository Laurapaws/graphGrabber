from cgitb import text
from email.mime import image
from fileinput import filename
from re import search
import fitz
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
import os
import io
from io import BytesIO
import tempfile


posDict = {
    # Define coordinates for positioning. Format is test name then test type. e.g. a VT-07 test with a mediumwave plot
    # Tuple Order: Left , Top, Width, Height
    #VT07
    'VT07MW' : (Pt(1), Pt(70), Pt(233), Pt(176)),
    'VT07FM1' : (Pt(239), Pt(70), Pt(233), Pt(176)),
    'VT07FM2' : (Pt(477), Pt(70), Pt(233), Pt(176)),
    'VT07DAB1AV' : (Pt(1), Pt(275), Pt(175), Pt(130)),
    'VT07DAB1RMS' : (Pt(180), Pt(275), Pt(175), Pt(130)),
    'VT07DAB2AV' : (Pt(360), Pt(275), Pt(175), Pt(130)),
    'VT07DAB2RMS' : (Pt(540), Pt(275), Pt(175), Pt(130)),
    #VT01 3 Metre
    'VT01' : (Pt(1), Pt(70), Pt(233), Pt(176))
}

cropDict = {
    # Define coordinates for cropping. Old refers to the old style PDFs we use that crop in two static positions
    'upperOld' : ((130, 138, 1000, 800)),
    'lowerOld' : ((130, 820, 1000, 1482))
}

nameDict = {
    # Define names of functions and maps them to their name written on the slide
    'VT01Ten' : 'VT-01 Off-Board Emissions (10m)',
    'VT01Three' : 'VT-01 Off-Board Emissions (3m)',
    'VT07' : 'VT-07 On-Board Emissions',
    'VT-12' : 'VT-12 Conducted Emissions',
    'VT-15' : 'VT-15 ElectricField'
}

# Defining global variables
extractedImages = []
croppedImages = []
rejectedList = []
slideCounter = 0


def searchReplace(search_str, repl_str, input, output):
    # Attempts to search and replace on the entire file. Likely needs rewriting to be more robust and not need a template
    prs = Presentation(input)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                if(shape.text.find(search_str))!=-1:
                    text_frame = shape.text_frame
                    cur_text = text_frame.paragraphs[0].runs[0].text
                    new_text = cur_text.replace(str(search_str), str(repl_str))
                    text_frame.paragraphs[0].runs[0].text = new_text
    prs.save(output)
    print(search_str + ' replaced with ' + repl_str)


def extractImages(PDFName, image_folder):
    fileName = image_folder + '/' + PDFName
    doc = fitz.open(fileName)
    zoom = 2 # to increase the resolution
    mat = fitz.Matrix(zoom, zoom)
    noOfPages = doc.pageCount
    for pageNo in range(noOfPages):
        page = doc.load_page(pageNo) #number of page
        pix = page.get_pixmap(matrix = mat)
        extractedImages.append(pix)
        print('Converting PDFs to Image')


def cropGraph(targetImg, cropTuple, imName):
    targetPIL = targetImg.tobytes("PNG")
    im = Image.open(io.BytesIO(targetPIL))
    im1 = im.crop(box=cropTuple)
    croppedImages.append(im1)
    print(imName +  ' cropped')


def insertImage(oldFileName, newFileName, img, positionTuple, slideNumber):
    # Inserts an image from the croppedImages array into slideNumber using a position from posDict
    prs = Presentation(oldFileName)
    slide = prs.slides[slideNumber]
    left = positionTuple[0]
    top = positionTuple[1]
    width = positionTuple[2]
    height = positionTuple[3]
    temp = BytesIO()
    img.save(temp, "PNG")
    slide.shapes.add_picture(temp, left, top, width, height)
    prs.save(newFileName)
    print('Image inserted')


def initialisePowerPoint(emptyDeckName, newDeckName):
    # Sets up the empty, fresh PPTX file
    emptyDeckName = emptyDeckName + '.pptx'
    newDeckName = newDeckName + '.pptx'
    prs = Presentation(emptyDeckName)
    prs.save(newDeckName)
    print('Created new PowerPoint: ' + newDeckName)


def VT07(PDFName, folderName, slideNumber, deckName):
    extractImages(PDFName, folderName)
    cropGraph(extractedImages[1], cropDict['upperOld'], 'MW')
    cropGraph(extractedImages[1], cropDict['lowerOld'], 'FM1')
    cropGraph(extractedImages[2], cropDict['upperOld'], 'FM2')
    cropGraph(extractedImages[2], cropDict['lowerOld'], 'DAB1AV')
    cropGraph(extractedImages[3], cropDict['upperOld'], 'DAB1RMS')
    cropGraph(extractedImages[3], cropDict['lowerOld'], 'DAB2AV')
    cropGraph(extractedImages[4], cropDict['upperOld'], 'DAB2RMS')
    deckName = deckName + '.pptx'
    insertImage(deckName, deckName, croppedImages[0], posDict['VT07MW'], slideNumber)
    insertImage(deckName, deckName, croppedImages[1], posDict['VT07FM1'], slideNumber) 
    insertImage(deckName, deckName, croppedImages[2], posDict['VT07FM2'], slideNumber) 
    insertImage(deckName, deckName, croppedImages[3], posDict['VT07DAB1AV'], slideNumber) 
    insertImage(deckName, deckName, croppedImages[4], posDict['VT07DAB1RMS'], slideNumber) 
    insertImage(deckName, deckName, croppedImages[5], posDict['VT07DAB2AV'], slideNumber) 
    insertImage(deckName, deckName, croppedImages[6], posDict['VT07DAB2RMS'], slideNumber)
    extractedImages.clear()
    croppedImages.clear()
    print('Finished VT07 for ' + PDFName)

def VT01Three(PDFName, folderName, slideNumber, deckName):
    extractImages(PDFName, folderName)
    cropGraph(extractedImages[1], cropDict['upperOld'], 'MW')
    deckName = deckName + '.pptx'
    insertImage(deckName, deckName, croppedImages[0], posDict['VT07MW'], slideNumber)
    extractedImages.clear()
    croppedImages.clear()
    print('Finished VT01 3m for ' + PDFName)


def setSlideCounter(num):
    global slideCounter
    slideCounter = num
    print('Slide counter set to ' + str(slideCounter))


def loopFolder(folderName, deckName, reportFunction):

    directory = 'testFolder'
    global slideCounter
    
    for file in os.listdir(directory):
        if file.endswith(".Pdf") or file.endswith(".pdf"):
            print('Working on slide ' + str(slideCounter) + ', File Name: ' + file)
            reportFunction(file, folderName, slideCounter, deckName)
            searchString = '*' + str(slideCounter) + '*'
            replaceString = str(file)[:-4] + ' | ' + (nameDict[str(reportFunction.__name__)])
            searchReplace(searchString, replaceString, deckName + '.pptx', deckName + '.pptx')
            slideCounter = slideCounter + 1
    print('Finished with folder: ' + folderName)


initialisePowerPoint('emptyDeck', 'newDeck')
setSlideCounter(0)
loopFolder('testFolder','newDeck', VT01Three)
print('Finished all jobs...')
