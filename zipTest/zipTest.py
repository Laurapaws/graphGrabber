import zipfile



archive = zipfile.ZipFile("/Users/laura/VSCode/GraphGrabber/graphGrabber/zipTest/testReport.docx")
#zipfile.ZipFile.printdir(archive)
for file in archive.filelist:
    if file.filename.startswith('word/media/') and file.file_size > 300:
        print(file.filename)
        archive.extract(file.filename)