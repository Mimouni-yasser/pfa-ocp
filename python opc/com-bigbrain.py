import win32com.client
import ctypes
import pywintypes
import schedule
from time import sleep
from enum import Enum



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

class entery():
    insert_offset = 0
    def __init__(self, q: str, code: str, ref_des:str, dest:str, ref_sto:str, bascule:str, index_bascule: int, index_horaire: int, post: str):
        self.quality = q
        self.code = code
        self.repere_destokage = ref_des
        self.repere_stokage = ref_sto
        self.destination = dest
        self.bascule = bascule
        self.index_bascule = index_bascule
        self.index_horaire = index_horaire
        self.post = post
    
    def update_index(self, bascule, horaire):
        self.index_bascule = bascule
        self.index_horaire = horaire
    
    
# def write_entery_to_excel(current_enteries: list[entery]):
#     workbook.Sheets[]

#TODO: wrap everything in a main function and make it callable with sys args


arr_type = ctypes.c_int * 10

if __name__ == "__main__":
    
    arr = arr_type(1) #safe array to access data from the COM interface, no reason to be 10 elements long just did it 
    arr[1] = 2 #for whatever stupid fucking reason, the COM interface starts arrays from 1 instead of 0; thus we put the propertyID of the Value in the 1st (which is the second) element of the array
    arr[2] = 3 #reading item quality too, just u know for shits and giggles


    try:
        client = win32com.client.Dispatch('Graybox.OPC.DAWrapper') #this wrapper provides functions that allow interfacing with OPC-DA through COM.
    except:
        print(s.RED + "cannot create win32 client" + s.RESET) 
        exit()

    #TODO: this should probably be a system argument. 

    try:
        result = client.Connect('Schneider-Aut.OFS.2') #the server in question
    except pywintypes.com_error as e:
        print(s.RED + "Client connect failed with error {0},\nCheck the server is running and the server name is spelt correctly".format(e.hresult) + s.RESET)
        client.Disconnect()
        exit()
        
        
    print(s.YELLOW + 'testing connectivety with OPC server...' + s.RESET)
    try:
        result = client.GetItemProperties('DevExample_1!Energie_1_Active_BPBIS', 2, arr)
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
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    except:
        print(s.RED + "error occured while creating the excel com client"+s.RESET)
        client.Disconnect()
        
    print(s.GREEN + "excel COM client succesfully created" + s.RESET)

    print(s.YELLOW + "openning excel file..."+s.RESET)

    try:
        workbook = excel.Workbooks.Open("C:/Users/Yasser Mimouni/Desktop/pfa me/python opc/book.xlsx")
    except pywintypes.com_error as e:
        print(s.RED + "error occured while loading file to excel instance: " + s.RESET)
        print(e)
        client.Disconnect()
        excel = None

    print(s.GREEN + "excel file loaded, configuration finished sucessfully." + s.RESET)
    excel.Visible = True
    ########################
    # sleep(5)
    # client.Disconnect()
    # workbook.Close(False)
    # excel = None
    ######################
    print(s.YELLOW + "Starting read-write loop" +s.RESET)

    #schedule.every(3).seconds.do(write)
    excel_entery = entery('', '','', '', '', 0, 0, 0, '')
    schedule.every().hour.do(write_entery_to_excel, excel_entery)

    while True:
        
        schedule.run_pending()

