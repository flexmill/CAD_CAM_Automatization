# -*- coding: utf-8 -*-
"""
Created on Mon Jun 27 11:25:01 2022

@author: j.erlacher
"""

import os
import sys
from pycatia import catia


# adress catpart 
sys.path.insert(0, os.path.abspath('..\\pycatia'))

caa = catia()  

documents = caa.documents

documents.open(r'C:\Users\j.erlacher\RO-RA Aviation Systems GmbH\DIH2-FlexMill - Dokumente\20_work\04_Design\Structure Rod End\TEST\1F-B-00382-006.CATPart')
    
partdoc = caa.active_document

documents.open(r'C:\Users\j.erlacher\RO-RA Aviation Systems GmbH\DIH2-FlexMill - Dokumente\20_work\04_Design\Structure Rod End\structure_rodEnd.CATDrawing')

drawingdoc = caa.active_document
sheets = drawingdoc.sheets
sheet = sheets.item("Main_View")
views = sheet.views

for y in range (1,views.count):
    view = views.item(y)
    view.activate()
    blinks = view.generative_behavior
    blinks2 = view.generative_links
    blinks2.remove_all_links
    mypart = caa.documents.item("1F-B-00382-006.CATPart")
    blinks.document = mypart
    print(mypart.name)
