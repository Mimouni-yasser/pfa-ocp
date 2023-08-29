import win32com.client
import ctypes
import pywintypes
import schedule
from enum import Enum
import sys
import os.path
import datetime

from pathlib import Path
import win32com

gen_py_path = os.getcwd() + '/gen_py'

Path(gen_py_path).mkdir(parents=True, exist_ok=True)
win32com.__gen_path__ = gen_py_path

arr_type = ctypes.c_int * 10

QUALITY_CHANGED = False #REMOVE THIS WHEN QUALITY CHANGED IS IMPLEMENTED

class s():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    BOLD = '\033[1m'
    WHITE  = '\33[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'

    
class sheet(Enum):
    RPT1 = 1
    RPT2 = 2
    RP3 = 3
    RP5 = 4
    RPTS = 5
    CRIBLAGE = 6
    Arret_RPT1 = 7
    Arret_RPT2 = 8
    Arret_RP3 = 9
    Arret_RP5 = 10
    Arret_RPTS = 11
    Arret_Criblage = 12
    Arrets = 13
class col_index(Enum):
    quality = 3
    code = 4
    repere_destokage = 5
    destination = 6
    repere_stokage = 7
    bascule = 8
    index_bascule = 10
    index_horaire = 12
    post = 14
    heures = 2
    

class entery():
    count = 16
    def __init__(self, quality: str, code: str, repere_destokage:str, destination:str, repere_stokage:str, bascule:str, index_bascule: int, index_horaire: int, post: str):
        self.quality = quality
        self.code = code
        self.repere_destokage = repere_destokage
        self.repere_stokage = repere_stokage
        self.destination = destination
        self.bascule = bascule
        self.index_bascule = index_bascule
        self.index_horaire = index_horaire
        self.post = post
        self.insert_offset = entery.count + 1
        entery.count = entery.count + 1
        self.heure = datetime.datetime.now().strftime("%H:00")
        self.color = '0xFFFFFF'
        
    def update_index(self, bascule, horaire):
        self.index_bascule = bascule
        self.index_horaire = horaire


        
def write_entery_to_excel(current_enteries):
    
    worksheet = workbook.Sheets(1)
    for sheet in workbook.Sheets:
        sheet.Range("R10").Value = datetime.datetime.now().strftime("%d/%m/%Y")
        
    for entery in current_enteries:
        worksheet.Cells(entery.insert_offset, col_index.quality.value).Value = entery.quality
        worksheet.Cells(entery.insert_offset, col_index.code.value).Value = entery.code
        worksheet.Cells(entery.insert_offset, col_index.repere_destokage.value).Value = entery.repere_destokage
        worksheet.Cells(entery.insert_offset, col_index.destination.value).Value = entery.destination
        worksheet.Cells(entery.insert_offset, col_index.repere_stokage.value).Value = entery.repere_stokage
        worksheet.Cells(entery.insert_offset, col_index.bascule.value).Value = entery.bascule
        worksheet.Cells(entery.insert_offset, col_index.index_bascule.value).Value = entery.index_bascule
        worksheet.Cells(entery.insert_offset, col_index.index_horaire.value).Value = entery.index_horaire
        worksheet.Cells(entery.insert_offset, col_index.post.value).Value = entery.post
        entery.heure = datetime.datetime.now().strftime("%H:00")
        worksheet.Cells(entery.insert_offset, col_index.heures.value).Value = entery.heure
    #save workbook
    
def temp_save():
    time = datetime.datetime.now().strftime("%H-%M-%S")
    date = datetime.datetime.now().strftime("%d-%m-%Y")
    workbook.SaveCopyAs(os.getcwd() + "/TEMP SAVES/temp-save at "+date + " " + time + ".xlsm")
    
def next_row(changed: bool, excel_entery: entery):
    if(changed):
        excel_entery.color = '0xFFFF00'
        excel_entery.heure = ' '
    else:
        excel_entery.color = '0xFFFF00'
        excel_entery.heure = datetime.datetime.now().strftime("%H:00")
    excel_entery.insert_offset = excel_entery.insert_offset + 1


def check_quality(current_entry: entery):
    global QUALITY_CHANGED
    if(current_entry.quality != 'wtf'):
        pass
    else:
        QUALITY_CHANGED = True
    
def save_workbook():
    for sheet in workbook.Sheets:
        sheet.Range("R10").Value = datetime.datetime.now().strftime("%d/%m/%Y")
        sheet.Range("Q15").Value = datetime.datetime.now().strftime("%d/%m/%Y")
    workbook.SaveAs(os.getcwd() + '/rapport journalier {0}.xlsx'.format(datetime.datetime.now().strftime("%d-%m-%Y")))



#check sys args
if __name__ == "__main__":
    #check if excel fileexists
    
    try:
        excel_file = sys.argv[1]
    except IndexError:
        excel_file = os.getcwd() + '/book.xlsm'
        
    if not os.path.isfile(excel_file):
        print(s.RED + "Excel file not found" + s.RESET)
        exit()


    try:
        OPC_server = sys.argv[3]
    except IndexError:
        OPC_server = 'Schneider-Aut.OFS.2'
        

    try:
        wrapper = sys.argv[4]
    except IndexError:
        wrapper = 'Graybox.OPC.DAWrapper'

    try:
        item = sys.argv[2]
    except IndexError:
        item = 'DevExample_1!Energie_1_Active_BPBIS'

    
    arr = arr_type(1) 
    arr[1] = 2 
    arr[2] = 3 


    try:
        client = win32com.client.Dispatch(wrapper) #this wrapper provides functions that allow interfacing with OPC-DA through COM.
    except:
        print(s.RED + "cannot create win32 client" + s.RESET) 
        exit()

    try:
        result = client.Connect(OPC_server) #the server in question
    except pywintypes.com_error as e:
        print(s.RED + "Client connect failed with error {0},\nCheck the server is running and the server name is spelt correctly".format(e.hresult) + s.RESET)
        client.Disconnect()
        exit()
    print(s.YELLOW + 'testing connectivety with OPC server...' + s.RESET)
    try:
        result = client.GetItemProperties(item, 2, arr) # the 2 means reading 2 items, the arr has the item property IDs (2 for value 3 for quality)
    except pywintypes.com_error as e:
        print(s.RED + "Method call failed with HRESULT {}".format(e.hresult))
        print("\n error decoded: {0}", format(client.GetErrorString(e.hresult)) + s.RESET)
        client.Disconnect()
        exit()
        
    if(result[1] != (0, 0)):
        print("error {1} getting result, exiting...".format(result[1]))
        client.Disconnect()
        exit()
    else:
        print(s.GREEN + "all good, setting up excel... (values read = {0})".format(result[0]) + s.RESET)

    try:
        excel = win32com.client.Dispatch('Excel.Application')
    except:
        print(s.RED + "error occured while creating the excel com client"+s.RESET)
        client.Disconnect()
        
    print(s.GREEN + "excel COM client succesfully created" + s.RESET)

    print(s.YELLOW + "openning excel file..."+s.RESET)

    try:
        workbook = excel.Workbooks.Open(excel_file)
    except pywintypes.com_error as e:
        print(s.RED + "error occured while loading file to excel instance: " + s.RESET)
        print(e)
        client.Disconnect()
        excel = None
        exit()

    try:
        os.mkdir("TEMP SAVES")
    except FileExistsError:
        pass
    
    
    print(s.GREEN + "excel file loaded, configuration finished sucessfully." + s.RESET)
    excel.Visible = True
    print(s.YELLOW + "Starting read-write loop" +s.RESET)

    #schedule.every(3).seconds.do(write)
    excel_entery = entery("quality", "code", "repere_destokage", "destination", "repere_stokage", "bascule", 0, 0, "post")
    excel_entery.insert_offset = 17
    schedule.every(3).seconds.do(write_entery_to_excel, [excel_entery])
    schedule.every(30).minutes.do(temp_save)
    schedule.every().hour.do(next_row, False, excel_entery)
    #schedule.every(5).seconds.do(check_quality, excel_entery)
    schedule.every().day.at("00:00").do(save_workbook, [excel_entery])

    while True:
        
        #TODO read index bascule and index horaire from server (get the var names)
        try:
            result = client.GetItemProperties(item, 2, arr) # the 2 means reading 2 item properties, the arr has the item property IDs (2 for value 3 for quality)
        except pywintypes.com_error as e:
            print(s.RED + "Method call failed with HRESULT {}".format(e.hresult))
            print("\nerror decoded: {0}", format(client.GetErrorString(e.hresult)) + s.RESET)
            client.Disconnect()
            exit()  
        if(result[1] == (0, 0)):
            excel_entery.update_index(result[0][0], result[0][1])
        else:
            print(s.RED + "error {1} getting result, exiting...".format(result[1]))
            client.Disconnect()
            workbook.Close(False)
            exit()
        
        if(QUALITY_CHANGED):
            next_row(True, excel_entery)
            write_entery_to_excel([excel_entery])
            QUALITY_CHANGED = False
            
            
        schedule.run_pending()
