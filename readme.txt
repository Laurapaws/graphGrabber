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


-------------------------------------------------------


## User Guide

Creating a new deck is simple, but can break if you give it the wrong files. Here's a simple set of instructions:

### Basic Guide
1. Run GraphGrabber in its own folder.
2. Click Init Folders to build the required folder structure (first run only).
3. Click Check Files to verify folders exist and are empty.
4. Place PDF reports into their respective folders. (Use Open Working Directory to visit in a file explorer).
5. Type desired output file name.
6. Click Create Deck!
7. Output file will be found in the working directory with a Unix timestamp
8. Logfile is generated as GG.log in the working directory.
9. You can clear out all files from GraphGrabber folders with the Clear Folders button. Warning: This deletes instantly!

### Autosort
1. First select whether you want to place VT-12 CE into the Single or Three Phase folder (Default Single).
2. Click Autosort
3. Select the folder containing PDF reports
4. Submit
5. Autosort will attempt to place files into the right folders depending on their file name. Please check them!
6. Unsorted files will go into the Unsorted PDFs folder.


-------------------------------------------------------


## Guide to Functions etc.


**posDict**
Coordinates that determine where an image will be placed on the slide.
Tuple Order: Left , Top, Width, Height
    
*Example*
"VT07MW": (Pt(1), Pt(70), Pt(233), Pt(176))

**cropDict**
Similar to posDict but this one determines where in the PDF image that GraphGrabber crops a graph. There is an extra hardcoded value for 


