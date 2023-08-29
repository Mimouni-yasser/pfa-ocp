import win32com.client
import ctypes
import pywintypes
import schedule
import argparse
from enum import Enum
import sys
import os.path
import datetime
from pathlib import Path
import win32com


insert_offset = 17

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

    
sheet = {
    'RPT1': 1,
    'RPT2': 2,
   'RP3' : 3,
    'RP5' : 4,
    'RPTS' : 5,
    'CRIBLAGE' : 6,
    'Arret_RPT1' : 7,
    'Arret_RPT2' : 8,
    'Arret_RP3' : 9,
    'Arret_RP5' : 10,
    'Arret_RPTS' : 11,
    'Arret_Criblage' : 12,
    'Arrets' : 13
}
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
    

class excel_entry():
    count = 16
    def __init__(self, index_bascule, index_horaire , post, quality =' ', code=' ', repere_destokage =' ', destination = ' ', repere_stokage = ' ', bascule = ' '):
        self.quality = quality
        self.code = code
        self.repere_destokage = repere_destokage
        self.repere_stokage = repere_stokage
        self.destination = destination
        self.bascule = bascule
        self.index_bascule = index_bascule
        self.index_horaire = index_horaire
        self.post = post
        self.heure = datetime.datetime.now().strftime("%H:00")
        self.color = '0xFFFFFF'
        
    def update_index_and_write(self, bascule, horaire, worksheet):
        self.index_bascule = bascule
        self.index_horaire = horaire
        
        worksheet.Cells(insert_offset, col_index.quality.value).Value = self.quality
        worksheet.Cells(insert_offset, col_index.code.value).Value = self.code
        worksheet.Cells(insert_offset, col_index.repere_destokage.value).Value = self.repere_destokage
        worksheet.Cells(insert_offset, col_index.destination.value).Value = self.destination
        worksheet.Cells(insert_offset, col_index.repere_stokage.value).Value = self.repere_stokage
        worksheet.Cells(insert_offset, col_index.bascule.value).Value = self.bascule
        worksheet.Cells(insert_offset, col_index.index_bascule.value).Value = self.index_bascule
        worksheet.Cells(insert_offset, col_index.index_horaire.value).Value = self.index_horaire
        worksheet.Cells(insert_offset, col_index.post.value).Value = self.post
        self.heure = datetime.datetime.now().strftime("%H:00")
        worksheet.Cells(insert_offset, col_index.heures.value).Value = self.heure


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
                    prog='python-enabled reporting for OCP',
                    description='read from OFS put into excel automatically',
                    epilog=' ')
    parser.add_argument('--item', dest='itemID', help='ID de(s) variable(s) à lire', action='extend', nargs="+", type=str)
    parser.add_argument('--file', help = 'path absolue de fichier Exel de reference, par defaut book.xlsm dans le path courant', default='book.xlsm')
    parser.add_argument('--server', '-s', help = 'nom de serveur OPC-DA, par defaut = "Schneider-Aut.OFS.2"', default="Schneider-Aut.OFS.2")
    parser.add_argument('--wrapper', '-W', help='autowrapper pour OPC-DA, default "Graybox.OPC.DAWrapper"', default='Graybox.OPC.DAWrapper')
    parser.add_argument('--sheet', '-Se', help= 'nom de fiche excel sur laquelle les resultats sont affichés', default='RPT1')
    args = parser.parse_args()
    
    if args.file == "book.xlsm":
        excel_file_path = os.getcwd() + "\\" + args.file
    else:
        excel_file_path = args.file
        
    DAwrapper = args.wrapper
    server = args.server
    itemIDs = args.itemID
    
    print(s.BLUE + s.BOLD + "Initialisation de client COM de serveur OPC..."+s.RESET)
    
    try:
        client = win32com.client.Dispatch(DAwrapper)
    except pywintypes.com_error as e:
        print(s.RED + "créeation de client COM à echoue" + s.RESET) 
        print(e)
        exit()
    
    try:
        client.Connect(server)
    except pywintypes.com_error as e:
        print(s.RED + "initalisation de connection avec le serveur COM à echoue" + s.RESET) 
        print(e)
        exit()
    
    print(s.GREEN + s.BOLD + "Fin d'Initialisation de client COM et connection reussit avec serveur..."+s.RESET)
    
    print(s.BLUE + s.BOLD + "Initialisation de client COM de serveur Excel..."+s.RESET)
    
    try:
        excel = win32com.client.Dispatch('Excel.Application')
    except pywintypes.com_error as e:
        print(s.RED + "Creaction de client COM Excel echoue"+s.RESET)
        client.Disconnect()
        exit()
    
    workbook = excel.Workbooks.Open(excel_file_path)
    excel.Visible = True
    
    
    excel_sheet = sheet.get(args.sheet)
    if excel_sheet is None:
        print(s.RED + "fiche choisir n'exist pas dans le fichier, continuer? (y/n)")
        i = input()
        if i == "n": exit()
    worksheet = worksheet = workbook.Sheets(excel_sheet)
        
    print(s.GREEN + s.BOLD + "Fin d'Initialisation de client COM et connection reussit avec serveur..."+s.RESET)
    
    
    
    client.Disconnect()
    
        
    
        
    print(args)