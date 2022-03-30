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
    'VT07MW' : (Pt(1), Pt(70), Pt(233), Pt(176)),
    'VT07FM1' : (Pt(239), Pt(70), Pt(233), Pt(176)),
    'VT07FM2' : (Pt(477), Pt(70), Pt(233), Pt(176)),
    'VT07DAB1AV' : (Pt(1), Pt(275), Pt(175), Pt(130)),
    'VT07DAB1RMS' : (Pt(180), Pt(275), Pt(175), Pt(130)),
    'VT07DAB2AV' : (Pt(360), Pt(275), Pt(175), Pt(130)),
    'VT07DAB2RMS' : (Pt(540), Pt(275), Pt(175), Pt(130))
}

cropDict = {
    # Define coordinates for cropping. Old refers to the old style PDFs we use that crop in two static positions
    'upperOld' : ((130, 138, 1000, 800)),
    'lowerOld' : ((130, 820, 1000, 1482))
}

extractedImages = []
croppedImages = []

def searchReplace(search_str, repl_str, input, output):
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
        print('Converting PDFs to Image ... ')

def cropGraph(targetImg, cropTuple, imName):
    targetPIL = targetImg.tobytes("PNG")
    im = Image.open(io.BytesIO(targetPIL))
    im1 = im.crop(box=cropTuple)
    croppedImages.append(im1)
    #imName = 'testFolder/' + imName + '.png'
    #im1.save(imName)
    #print('Cropped and saved: ' + imName)

def insertImage(oldFileName, newFileName, img, positionTuple, slideNumber):
    #Be sure to call HALF the size you really want for the image. PowerPoint will auto resize
    prs = Presentation(oldFileName)
    slide = prs.slides[slideNumber]
    # img = 'testFolder/' + img
    left = positionTuple[0]
    top = positionTuple[1]
    width = positionTuple[2]
    height = positionTuple[3]
    temp = BytesIO()
    img.save(temp, "PNG")
    slide.shapes.add_picture(temp, left, top, width, height)
    prs.save(newFileName)

def initialisePowerPoint(emptyDeckName, newDeckName):
    emptyDeckName = emptyDeckName + '.pptx'
    newDeckName = newDeckName + '.pptx'
    prs = Presentation(emptyDeckName)
    prs.save(newDeckName)

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

def loopFolder(folderName, deckName):
    initialisePowerPoint('emptyDeck', deckName)
    directory = 'testFolder'
    slideCounter = 0
    for file in os.listdir(directory):
        if file.endswith(".Pdf") or file.endswith(".pdf"):
            print('Working on slide ' + str(slideCounter) + ', File Name: ' + file)
            VT07(file, folderName, slideCounter, deckName)
            searchString = '*' + str(slideCounter) + '*'
            replaceString = str(file)[:-4]
            replaceString = replaceString + ' - VT-07 On-Board Emissions'
            print(replaceString)
            searchReplace(searchString, replaceString, deckName + '.pptx', deckName + '.pptx')
            slideCounter = slideCounter + 1

loopFolder('testFolder','newDeck')

print('Finished all jobs...')
