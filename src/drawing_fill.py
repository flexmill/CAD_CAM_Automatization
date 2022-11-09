# -*- coding: utf-8 -*-
"""
@author: j.erlacher
"""

import openpyxl
import os
import getpass
import datetime
import sys
from pycatia import catia
import time

def artikel(path,paranr,paraname, paramat,paramatspec, RML,_RM,e_surface):
    print("PARAMETER ARE EXTRACED FOR ERP...")
    wb = openpyxl.load_workbook(path+'\\'+'Artikelanlage.xlsx')
    ws = wb.active
    data = []
    data = (paranr,paraname.value,paramat,RML.value,_RM,paramatspec,e_surface.value)
    
    
    counter = 1
    for i in range(len(data)):
        ws.cell(counter,2).value = str(data[i])
        counter = counter+1
        
    wb.save(path+'\\'+paranr+'.xlsx')

def IDL(caa,part,idl ):
    
    #declaration of the individual sketches and shape of part
    boobs = part.bodies
    boob = boobs.item("Drehkörper")
    _rodend = boobs.item("RODEND")
    _sketches = boob.sketches
    _sketches2 = _rodend.sketches
    no_idl = _sketches.item("_NO-IDL")
    _idl = _sketches2.item("_IDL")
    _shapes = _rodend.shapes
    _NOIDLTDR = _shapes.item("_NO-IDL_M10x1.25-RH")
    _IDLTDR = _shapes.item("_IDL_M10x1.25-RH") 
    _slot = _shapes.item("_SLOT_1")
    _rodsketch = _rodend.sketches
    _slotprofile = _rodsketch.item("_SLOTPROFILE")
    _slotcurve = _rodsketch.item("_SLOTCURVE")
    _slotpattern = _shapes.item("_SLOTPATTERN")
    
    if idl == 1:
        
        print("IDL IS ADDED...")
        #inactivation of sketch and thread of no IDL    
        part.inactivate(no_idl)
        part.inactivate(_NOIDLTDR)
        
        #replacemenet of the sketches of the rod end shaft body
        #a separate catscript is executed in CATIA
        macroname = "IDL.CATScript"
        modul = "CATMain"
        path_2 = r'N:\05 MAKROS UND TOOLS\01 Makros\FlexMill'
        system_service = caa.system_service
        system_service.execute_script(path_2,1, macroname,modul, [])
        
        
        #thread and slot for IDL is activated
        part.activate(_IDLTDR)
        part.activate(_slot)
        part.activate(_slotprofile)
        part.activate(_slotcurve)
        part.activate(_slotpattern)
        
        
    elif idl == 0:
        print("IDL IS REMOVED...")
        #inactivation of sketch and thread of IDL    
        part.inactivate(_idl)
        part.inactivate(_IDLTDR)
        
        #replacemenet of the sketches of the rod end shaft body
        #a separate catscript is executed in CATIA
        macroname = "NO_IDL.CATScript"
        modul = "CATMain"
        path_2 = r'N:\05 MAKROS UND TOOLS\01 Makros\FlexMill'
        system_service = caa.system_service
        system_service.execute_script(path_2,1, macroname,modul, [])
        
        
        #thread and slot for IDL is activate
        part.activate(_NOIDLTDR)
        part.inactivate(_slot)
        part.inactivate(_slotprofile)
        part.inactivate(_slotcurve)
        part.inactivate(_slotpattern)
    
def rod_spffile(paras,path_r,idl):
    print(".SPF FIlE IS CREATED...")
    os.chdir(path_r)
    # declaration of parameters needed for spf file
    paranr = paras.item("PART-NUMBER").value
    paranumber = '"'+ paranr+'"'
    paramat = paras.item("Material").value
    BOD = paras.item("BOD").value
    ROR = paras.item("ROR").value
    RBW = paras.item("RBW").value
    FASE = paras.item("FASE").value
    PD = paras.item("PD").value
    WINK = paras.item("WINK").value
    _FRKO = paras.item("FRKO").value
    if idl == 1:
        _VAR = 2
    elif idl == 0:
        _VAR = 1    
    _RM = ROR*2+ 0.5
    if str(paramat[:2]) == "EN":
        _MATE = 0
    elif str(paramat[:2]) == "17":
        _MATE = 1
    else:
        print("no correct material selected")
        
     #creation of two lists
    datapart =[]
    datapart=(paranumber,_VAR,_MATE,_RM,BOD,ROR,RBW,FASE,PD,WINK,_FRKO)
    data = []
    data=("_Partnumber","_VAR","_MATE","_RM","_BOD","_ROR","_RBW", "_FASE","_PD","_WINK","_FRKO")
    
    #merge the lists and play them out as spf file
    count_line=1
    while count_line <= len(datapart):
    
        spf_file=open(str(paranr)+r'.spf','w')
    
        count=0
        text=''
        for line in data:
            text=text+data[count]+'='+str(datapart[count])
            text=text+'\n'
            count=count+1
        text=text+'RET \n'  
        spf_file.write(str(text))
        spf_file.close()
        count_line+=1
    return(_RM)
        
def fork_spffile(paras, path_r, idl):
    print(".SPF FIlE IS CREATED...")
    os.chdir(path_r)
    # declaration of parameters needed for spf file
    paranr = paras.item("PART-NUMBER").value
    paranumber = '"'+ paranr+'"'
    paramat = paras.item("MATERIAL").value
    FIW = paras.item("FIW").value
    FD = paras.item("FD").value
    VDD = paras.item("FFR").value
    FCR = paras.item("FCR").value
    BUOD = paras.item("BUOD").value
    FMR = paras.item("FMR").value
    FDRI = paras.item("FDRI").value
    FDRA = paras.item("FDRA").value
    FOWA = paras.item("FOWA").value
    FIWA = paras.item("FIWA").value
    FWT = (FOWA-FIWA)
    FBL = paras.item("HD").value
    FTO = paras.item("FTO").value
    FDW = paras.item("FDW").value
    FTC = paras.item("FTC").value
    FTB = paras.item("FTB").value
    FLD = paras.item("FLD").value
    FLE = paras.item("FLE").value
    FLF = paras.item("FLF").value
    FLG = paras.item("FLG").value
    FLH = paras.item("FLH").value
    FLI = paras.item("FLI").value
    FLA = paras.item("FLA").value
    FLB = paras.item("FLB").value
    FLC = paras.item("FLC").value
    
    if idl == 1:
        _VAR = 4
    elif idl == 0:
        _VAR = 3    
    _RM = VDD*2+ 0.5
    if str(paramat[:2]) == "EN":
        _MATE = 0
    elif str(paramat[:2]) == "17":
        _MATE = 1
    else:
        print("no correct material selected")
     #creation of two lists
    datapart =[]
    datapart=(paranumber,_VAR,_MATE,_RM,FIW,FD,VDD,FCR,BUOD,FMR,FDRI,FDRA,FOWA,FWT,FBL,FTO,FDW,FTC,FTB,FLD,FLE,FLF,FLG,FLH,FLI,FLA,FLB,FLC)
    data = []
    data=("_Partnumber","_VAR","_MATE","_RM","_FIW","_FD","_VDD","_FCR","_BUOD","_FMR","_FDRI","_FDRA","_FOWA","_FWT","_FBL","_FTO","_FDW","_FTC","_FTB","_FLD","_FLE","_FLF","_FLG","_FLH","_FLI","_FLA","_FLB","_FLC")
    
    #merge the lists and play them out as spf file
    count_line=1
    while count_line <= len(datapart):
    
        spf_file=open(str(paranr)+r'.spf','w')
    
        count=0
        text=''
        for line in data:
            text=text+data[count]+'='+str(datapart[count])
            text=text+'\n'
            count=count+1
        text=text+'RET \n'  
        spf_file.write(str(text))
        spf_file.close()
        count_line+=1
    return(_RM)
    
    
   


# adress catpart 
sys.path.insert(0, os.path.abspath('..\\pycatia'))

caa = catia()  

path= r'C:\Users\j.erlacher\RO-RA Aviation Systems GmbH\DIH2-FlexMill - Dokumente\20_work\04_Design'

documents = caa.documents
partdoc = caa.active_document
part = partdoc.part
paras = part.parameters

#diff between rod and fork
paraname = paras.item("NAME").value 
if paraname[9] == "F":
    A_type = "FORK"
elif paraname[9] == "R":
    A_type = "ROD"    

#definie essential parameter
paranr = paras.item("PART-NUMBER").value
paramat = paras.item("MATERIAL").value
idln = paras.item("IDLH").value
paramatspec = paras.item("MAT_SPEC").value
parathread = paras.item("TRD").value
RML = paras.item("RML")
paraname = paras.item("NAME")

part.update()

# inquiry IDL
if A_type == "ROD":
    idlthread = paras.item("idlthread")
    if idln == True:
        if idlthread.value == False:
            idl = 1
            IDL(caa,part,idl)
        else:
            idl = 1
            
    elif idln.value == False:
        if idlthread.value == True:
            idl = 0
            IDL(caa,part,idl)
        else:
            idl = 0
    file_name = r'test_rodEnd.CATDrawing'
    path_r = path + "\\Interior Rod End\\"
    _RM = rod_spffile(paras, path_r, idl)
    
    
elif A_type == "FORK":
    if idln == True:
        idl = 1
            
    elif idln == False:
        idl = 0
    starttime = time.time()
    lasttime = starttime
    lapnum = 1
    file_name = r'test_fork.CATDrawing'
    path_r = path + "\\Interior Fork End\\"
    _RM = fork_spffile(paras, path_r, idl)
    laptime=((time.time()-starttime),2)
    print(str(laptime))
    
    

#for ID_t in range(6,10):
    #paraID.value = float(ID_t)
    #part.update()

#drawing paras



#opening catdrawing
print("OPEN DRAWING...")

documents.open(path+"\\MASTER_DRAWING_INTERIOR.CATDrawing")
drawingdoc = caa.active_document
sel1 = drawingdoc.selection

#selection of the correct sheet
sheets = drawingdoc.sheets
ridlsheet = sheets.item("ROD_IDL")
rnoidlsheet = sheets.item("ROD_NO_IDL")
fidlsheet = sheets.item("FORK_IDL")
fnoidlsheet = sheets.item("FORK_NO_IDL")

if A_type == "ROD":
    if idl == 1:
        sel1.clear()
        sel1.add(rnoidlsheet)
        sel1.add(fnoidlsheet)
        sel1.add(fidlsheet)
        sel1.delete()
        sheet = ridlsheet

    elif idl == 0:
        sel1.add(ridlsheet)
        sel1.add(fnoidlsheet)
        sel1.add(fidlsheet)
        sel1.delete()
        sheet = rnoidlsheet
    
elif A_type == "FORK":
    if idl == 1:
        sel1.clear()
        sel1.add(fnoidlsheet)
        sel1.add(rnoidlsheet)
        sel1.add(ridlsheet)
        sel1.delete()
        sheet = fidlsheet

    elif idl == 0:
        sel1.add(fidlsheet)
        sel1.add(rnoidlsheet)
        sel1.add(ridlsheet)
        sel1.delete()
        sheet = fnoidlsheet

sheet.force_update()

#input of the drawing parameters
dwgparas = drawingdoc.parameters
dwgnr = dwgparas.item("NUMBER")
dwgmat = dwgparas.item("MATERIAL")
dwgname = dwgparas.item("NAME")
dwgnr.value = paranr
dwgmat.value = paramat+" // "+paramatspec
dwgname = paraname

    
#declaration of thread tolerances and right/ left handed thread
views = sheet.views
view = views.item("Main_View")
view.activate()
dimension = view.dimensions.item("THREAD")
di = dimension.get_value()

thread_h = str(parathread[-6:])
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
    surface.value = "PASSIVIERT LAUT AMS2700, METHODE 1, KLASSE 4"
    e_surface.value = "PASSIVATED IN ACCORDANCE WITH AMS2700, METHOD 1, CLASS 4"

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

drawingdoc.save_as(path_r +"\\"+ str(dwgnr.value)+"_ROD_END")
drawingdoc.export_data(path_r+"\\"+ str(dwgnr.value) +"_ROD_END"+r'.dxf', "dxf")
drawingdoc.export_data(path_r+"\\"+ str(dwgnr.value) +"_ROD_END"+r'.pdf', "pdf", overwrite=True)

artikel(path,paranr,paraname, paramat,paramatspec, RML,_RM,e_surface)

print("CLOSING DRAWING")

drawingdoc.close()

print("SCRIPT END")
    
    
        





