import win32com.client
import ctypes
import pywintypes
import schedule
from time import sleep
import random



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

class row():
    insert_at = None
    
    
def write():
    cell = workbook.Sheets(1).Range('J22')
    result = client.GetItemProperties('API!Energie_Active2_PLT3', 2, arr)
    cell.Value = result[0][0]
    

#TODO: wrap everything in a main function and make it callable with sys args


arr_type = ctypes.c_int * 10
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
    result = client.GetItemProperties('API!Energie_Active2_PLT3', 2, arr)
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
    workbook = excel.Workbooks.Open(r"C:\Users\lenovo\Desktop\RP5\RAPPORT DE PRODUCTION\book.xlsm")
except pywintypes.com_error as e:
    print(s.RED + "error occured while loading file to excel instance: " + s.RESET)
    print(e)
    client.Disconnect()
    excel = None

print(s.GREEN + "excel file loaded, configuration finished good sucessfully." + s.RESET)
excel.Visible = True
########################
# sleep(5)
# client.Disconnect()
# workbook.Close(False)
# excel = None
######################
print(s.YELLOW + "Starting read-write loop" +s.RESET)

schedule.every(3).seconds.do(write)

while True:
    schedule.run_pending()

