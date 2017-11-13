#coding=utf8
import xmind
import os
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from xmind.core.topic import TopicElement

# this will consume pdf file and output it's xmind map
def PDF2XMIND(PDF_NAME,XMIND_NAME):
    # Open a PDF document.
    fp = open(PDF_NAME, 'rb')
    parser = PDFParser(fp)
    password = ''
    document = PDFDocument(parser, password)
    # Get the outlines of the document.
    outlines = document.get_outlines()

    w = xmind.load(XMIND_NAME) # load an existing file or create a new workbook if nothing is found
    s1=w.getPrimarySheet() # get the first sheet
    s1.setTitle("sheet1") # set its title
    r1=s1.getRootTopic() # get the root topic of this sheet
    r1.setTitle(PDF_NAME) # set its title


    preTop1 = TopicElement()
    preTop2 = TopicElement()
    # create a new branch if meet another level 1
    for (level,title,dest,a,se) in outlines:
        top = TopicElement()
        top.setTitle(title.encode('utf-8'))
        if level == 1:
            preTop = top
            r1.addSubTopic(top)
        else:
            if level == 2:
                preTop1.addSubTopic(top)
            else:
                preTop2.addSubTopic(top)

    xmind.save(w,XMIND_NAME)  # and we save


# for filename in os.listdir(r'C:\Users\yuantao01\OneDrive\KPI\pdf'):
#     if filename.endswith('pdf'):
#         PDF2XMIND(filename,filename+'.xmind')
PDF_NAME = "Java.pdf"
XMIND_NAME = "Java.xmind"
PDF2XMIND(PDF_NAME,XMIND_NAME)
