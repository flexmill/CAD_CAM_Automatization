# -*- coding: utf-8 -*-
"""
@author: j.erlacher
"""

import os
import getpass
import datetime
import sys
from pycatia import catia
from os.path import expanduser
import shutil


def replace (path,source,file_number):
    
    documents.open(path+'\\'+'1F-B-003XX-000_ROD_END.CATDrawing')
    drawingdoc = caa.active_document
    
    documents.open(path+'\\'+'RodEnd_Structure.CATPart')
    partdoc = caa.active_document
    partdoc.save_as(path+'\\'+file_number)
    partdoc.close()
    drawingdoc.save_as(path+'\\'+str(file_number)[:-8]+'_ROD_END')
    drawingdoc.close()
    
    os.remove(path+'\\'+file_number)
    shutil.move(source+'\\'+file_number,path+'\\'+file_number)
    
    documents.open(path +'\\'+str(file_number)[:-8]+'_ROD_END.CATDrawing')
    drawingdoc = caa.active_document
    sheets = drawingdoc.sheets
    sheet = sheets.item('1F-B-003XX-000')
    sheet.force_update()
    drawingdoc.save()

# adress catpart 
sys.path.insert(0, os.path.abspath('..\\pycatia'))
caa = catia()  

home = expanduser ("~")

path= home + r'\RO-RA Aviation Systems GmbH\DIH2-FlexMill - Dokumente\20_work\04_Design\Structure Rod End\ordner1'
source = home + r'\RO-RA Aviation Systems GmbH\DIH2-FlexMill - Dokumente\20_work\04_Design\Structure Rod End\ordner2'

files = os.listdir(source)
files_CATPart = [i for i in files if i.endswith('.CATPart')]

for i in range (len(files_CATPart)):
    file = source+'\\'+files_CATPart[i]
    file_number = files_CATPart[i]

    documents = caa.documents
    documents.open(file)
    
    partdoc = caa.active_document
    part = partdoc.part
    paras = part.parameters
    
    #for ID_t in range(6,10):
        #paraID.value = float(ID_t)
        #part.update()
    
    #define paras import for drawing
    paranr = paras.item("NUMBER").value
    paramat = paras.item("MATERIAL").value
    paramatspec = paras.item("MATERIAL_SPEC").value
    parathread = paras.item("TRD").value
    #part.update()
    
    partdoc.close()
    
    replace(path,source,file_number)
    
    #opening catdrawing
    print("OPEN DRAWING...")
    
    drawingdoc = caa.active_document
    sel1 = drawingdoc.selection
    
        
    #input of the drawing parameters
    dwgparas = drawingdoc.parameters
    dwgnr = dwgparas.item("NUMBER")
    dwgmat = dwgparas.item("MATERIAL")
    dwgnr.value = paranr
    dwgmat.value = paramat+" // "+paramatspec
    
    #selection of the correct sheet
    sheets = drawingdoc.sheets
    sheet = sheets.item(dwgnr.value)
    
        
    #declaration of thread tolerances and right/ left handed thread
    views = sheet.views
    view = views.item("A1-A1")
    view.activate()
    dimension = view.dimensions.item("THREAD")
    di = dimension.get_value()
    
    thread_h = str(parathread[6:])
    di.set_bault_text(1,"",thread_h,"","")
    
    #add surface note in general notes
    background = views.item(2)
    background.activate()
    _tables = background.tables
    _table = _tables.item(2)
    if str(dwgmat.value[:2]) == "EN":
        surface = dwgparas.item("Surface")
        e_surface = dwgparas.item("e_surface")
        surface.value = "SCHWEFELSÄUREANODISIEREN LAUT MIL-PRF-8625 TYP II, KLASSE, SCHICHTDICKE 12+/-3µm, NACHVERDICHTEN MIT DEIONISIERTEN HEIßWASSER"
        e_surface.value = "SULFURIC ACID ANODIZING  IN ACCORDANCE WITH MIL-PRF-8625 TYPE II, CLASS 1, COAT THICKNESS 12+/-3µm, REPRESSING WITH BOILING DEIONIZED WATER"
        
    elif str(dwgmat.value[:2]) == "17" :
        surface = dwgparas.item("Surface")
        e_surface = dwgparas.item("e_surface")
        surface.value = "PASSIVIERT LAUT AMS2700, METHODE 1, TYP 8, KLASSE 3"
        e_surface.value = "PASSIVATED IN ACCORDANCE WITH AMS2700, METHOD 1, TYPE 8, CLASS 3"
        
    
    #oberflächenschutz
    #viewn = views.item("A1-A1")
    #viewn.activate()
    
    #diadim = viewn.dimensions.item("durchmessser")
    #diad = diadim.get_value()
    #dia = (diad.value - 1)/2
    
    #view = views.item("Vorderansicht")
    #view.activate()
    #factory_2d = view.factory_2d
    
    #circle = factory_2d.create_closed_circle(0,0, dia)
    #sel1 = drawingdoc.selection
    #sel1.add(circle)
    
    #visprops = drawingdoc.selection.vis_properties
    #visprops.set_real_line_type(5, 1)
    
    #fill out of name drawn and checked + date
    detail_sh = sheets.item("Details Zeichnungskopf").activate()
    detail_s = sheets.active_sheet
    detail_views = detail_s.views
    detail_v = detail_views.item("XX-X-XXXXX-XXX").activate()
    detail_schriftkopf = detail_views.active_view
    
    #user
    user = getpass.getuser()
    user_1 = user[0]
    user_2 = user[2]
    k_user = str(user_1+user_2)
    
    #date
    today = datetime.datetime.now().strftime('%d.%m.%Y')
    date_1 = today[:6]
    date_2 = today[-2:]
    date=date_1+date_2
    
    
    #fill in in the titleblock
    i = 1
    for i in range (1,detail_schriftkopf.texts.count):
        texts_n = detail_schriftkopf.texts.item(i)
        if texts_n.name == "Drawn":
            drawn = str(user[2:].upper())
            texts_n.text =drawn
            
    j = 1
    for j in range (1,detail_schriftkopf.texts.count):
        texts_k = detail_schriftkopf.texts.item(j)
        
        if texts_k.name == "K_Drawn":
            texts_k.text = k_user.upper()
            
    k = 1
    for k in range (1,detail_schriftkopf.texts.count):
        texts_d = detail_schriftkopf.texts.item(k)
        
        if texts_d.name == "Datum":
            texts_d.text = date
            
    #checked = input("Prüfer eingeben:")
    #m = 1
    #for m in range (1,detail_schriftkopf.texts.count):
        #texts_c = detail_schriftkopf.texts.item(m)
        
        #if texts_c.name == "Checked":
            #texts_c.text = checked.upper()
            
    sheet = sheets.item(dwgnr.value).activate()
    sheet = sheets.active_sheet
    
    drawingdoc.save()
    #drawingdoc.export_data(path+"\\"+ str(dwgnr.value) +"_ROD_END"+r'.dxf', "dxf")
    drawingdoc.export_data(path+"\\"+ str(dwgnr.value) +"_ROD_END"+r'.pdf', "pdf", overwrite=True)
    
    
    print("CLOSING DRAWING")
    
    drawingdoc.close()

print("SCRIPT END")
    
    
        





