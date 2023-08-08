import win32com.client
import ctypes
import pywintypes
import openpyxl
from time import sleep
import random


wb = openpyxl.load_workbook('book.xlsm')
sh = wb.active

#TODO: wrap everything in a main function and make it callable with sys args


arr_type = ctypes.c_int * 10
arr = arr_type(1) #safe array to access data from the COM interface, no reason to be 10 elements long just did it 

arr[1] = 2 #for whatever stupid fucking reason, the COM interface starts arrays from 1 instead of 0; thus we put the propertyID of the Value in the 1st (which is the second) element of the array
arr[2] = 3 #reading item quality too, just u know for shits and giggles


try:
    client = win32com.client.Dispatch('Graybox.OPC.DAWrapper') #this wrapper provides functions that allow interfacing with OPC-DA through COM.
except:
    print("cannot create win32 client ") 
    exit()

#TODO: this should probably be a system argument. 

try:
    result = client.Connect('Schneider-Aut.OFS.2') #the server in question
except pywintypes.com_error as e:
    print("Client connect failed with error {0},\nCheck the server is running and the server name is spelt correctly".format(e.hresult))
    client.Disconnect()
    exit()
    
    
try:
    result = client.GetItemProperties('DevExample_1!Energie_1_Active_BPBIS', 2, arr)
except pywintypes.com_error as e:
    print("Method call failed with HRESULT {}".format(e.hresult)) #don't understand shit about these error messages
    #TODO: add error message decoding ...
    #!nevermind can't seem to get the HRESULT :(
    #!HA got it
    print("\n error decoded: {0}", format(client.GetErrorString(e.hresult)))
    
    client.Disconnect()
    exit()
    
if(result[1] != (0, 0)):
    print("error {1} getting result, exiting...".format(result[1]))
    client.Disconnect()
    exit()
    
for i in range(2):
    print("value of propertyID {0} is {1}".format(i, result[0][i]))
    
for i in range(17, 40):
    result = client.GetItemProperties('DevExample_1!Energie_1_Active_Digue', 2, arr)
    print(result[0][0])
    sh["J{0}".format(i)] = result[0][0]
    sleep(random.uniform(0.5, 3))
    
client.Disconnect()
wb.save('book.xlsx')
    
# while True:
#     if(q_c != q_a): break
#     elif(indice_h > 100): break
#     else:
#         sleep(3.6)
#         indice_h +1 