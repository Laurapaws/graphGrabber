from cgitb import text
from email.mime import image
from fileinput import filename
from re import search
import fitz
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
import os


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
        output = image_folder + '/' + str(pageNo) + '.png'
        pix.save(output)
        print('Converting PDFs to Image ... ' + output)

def cropGraph(imPath, cropTuple, imName):
    im = Image.open(imPath)
    im1 = im.crop(box=cropTuple)
    #im1.show() #Don't need to show unless for testing
    imName = 'testFolder/' + imName + '.png'
    im1.save(imName)
    print('Cropped and saved: ' + imName)

def insertImage(oldFileName, newFileName, img, positionTuple, slideNumber):
    #Be sure to call HALF the size you really want for the image. PowerPoint will auto resize
    prs = Presentation(oldFileName)
    slide = prs.slides[slideNumber]
    img = 'testFolder/' + img
    left = positionTuple[0]
    top = positionTuple[1]
    width = positionTuple[2]
    height = positionTuple[3]
    slide.shapes.add_picture(img, left, top, width, height)
    prs.save(newFileName)
    print(img + ' pasted into ' + newFileName)
    os.remove(img)
    print(img + ' deleted')

def initialisePowerPoint(emptyDeckName, newDeckName):
    emptyDeckName = emptyDeckName + '.pptx'
    newDeckName = newDeckName + '.pptx'
    prs = Presentation(emptyDeckName)
    prs.save(newDeckName)

def dirtyCleanup(folderName):
    print('Deleting PDF Images')
    os.remove(folderName + '/0.png')
    os.remove(folderName + '/1.png')
    os.remove(folderName + '/2.png')
    os.remove(folderName + '/3.png')
    os.remove(folderName + '/4.png')
    os.remove(folderName + '/5.png')
    os.remove(folderName + '/6.png')
    print('Finished Deleting')

def VT07(PDFName, folderName, slideNumber, deckName):
    extractImages(PDFName, folderName)
    PNG1 = folderName + "/1.png"
    PNG2 = folderName + "/2.png"
    PNG3 = folderName + "/3.png"
    PNG4 = folderName + "/4.png"
    cropGraph(PNG1, cropDict['upperOld'], 'MW')
    cropGraph(PNG1, cropDict['lowerOld'], 'FM1')
    cropGraph(PNG2, cropDict['upperOld'], 'FM2')
    cropGraph(PNG2, cropDict['lowerOld'], 'DAB1AV')
    cropGraph(PNG3, cropDict['upperOld'], 'DAB1RMS')
    cropGraph(PNG3, cropDict['lowerOld'], 'DAB2AV')
    cropGraph(PNG4, cropDict['upperOld'], 'DAB2RMS')
    dirtyCleanup(folderName)
    deckName = deckName + '.pptx'
    insertImage(deckName, deckName,'MW.png', posDict['VT07MW'], slideNumber)
    insertImage(deckName, deckName,'FM1.png', posDict['VT07FM1'], slideNumber) 
    insertImage(deckName, deckName,'FM2.png', posDict['VT07FM2'], slideNumber) 
    insertImage(deckName, deckName,'DAB1AV.png', posDict['VT07DAB1AV'], slideNumber) 
    insertImage(deckName, deckName,'DAB1RMS.png', posDict['VT07DAB1RMS'], slideNumber) 
    insertImage(deckName, deckName,'DAB2AV.png', posDict['VT07DAB2AV'], slideNumber) 
    insertImage(deckName, deckName,'DAB2RMS.png', posDict['VT07DAB2RMS'], slideNumber)
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
