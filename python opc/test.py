import subprocess
import os
import openpyxl
from time import sleep

cmd = "opc -r -s Schneider-Aut.OFS.2 DevExample_1!Energie_2_Active_PB"

opc_env = os.environ.copy()
opc_env['OPC_CLASS'] = 'OPC.Automation'
opc_env['OPC_MODE'] = 'dcom'

workbook = openpyxl.load_workbook('exp.xlsx')
worksheet = workbook.active


#print(opc_env)

for i in range(1,6):
    returned_value = subprocess.run(cmd, env=opc_env, capture_output=True, shell = False) 
    tmp =  returned_value.stdout.split()[1].decode('utf-8')
    print('returned value:', tmp)
    sleep(1)
    worksheet.cell(row=1, column=i).value = tmp
    workbook.save('exp.xlsx')

workbook.save('exp.xlsx')


