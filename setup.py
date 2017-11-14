#coding=utf8
import xmind
import os
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from xmind.core.topic import TopicElement
"""
todo 
    bugfix 
        1.throwable when no outline 
        2.clean when open an old map
        3.make outline more editable when have 3.4.5.6.7.8...n level content
"""
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
    pretop1 = TopicElement()
    pretop2 = TopicElement()
    # create a new branch if meet another level 1
    for (level,title,dest,a,se) in outlines:
        top = TopicElement()
        top.setTitle(title.encode('utf-8'))
        if level == 1:
            pretop1 = top
            r1.addSubTopic(top)
        else:
            if level == 2:
                pretop2 = pretop1
                pretop1.addSubTopic(top)
            else:
                pretop2.addSubTopic(top)

    xmind.save(w,XMIND_NAME)  # and we save


PREFFIX = "pdfFiles/"
SUFFIX = ".xmind"
for str in [
    "Java_7hxjsyzjsj.pdf",
    "JVM_srlj.pdf",
    "MySQL_gxn.pdf",
    "Python_sjfx.pdf",
    "Spring Recipes A Problem-Solution Approach.pdf",
    "tlcl-en-cn.pdf"
]:
    PDF_NAME = PREFFIX+str
    XMIND_NAME = PREFFIX+str+SUFFIX
    PDF2XMIND(PDF_NAME,XMIND_NAME)
    print str+" convert completed !"
