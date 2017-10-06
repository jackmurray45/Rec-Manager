import xcelreader
import xcelwriter

import datetime
import openpyxl
from coordinateClass import Coordinate
from timeformatter import *


def first_selection() -> str:
    while(True):
        print('Please enter a letter to make a selection:\n\n',
          'E: Equipment related\n',
          'L: Liability related\n',
          'X: Close application\n',
          )
        result = input('\nInput: ')
        if result.upper() not in ['E', 'L', 'X']:
            print('Please enter valid input')
        else:
            return result.upper()

def equipment_sheet() -> str:
    while(True):
        print('Please enter a letter to make a selection:\n\n',
              'P: Ping Pong\n',
              'POOL: Billards use\n',
              'XBOX: XBOX \n',
              'PS4: PS4 \n',
              'BG: Board Games \n',
              'S: Sports Ball \n',
              'Y: Yoga Mats \n',
              'D: DVDs \n',
              'M: Miscellaneous \n',
              'B: Go Back \n',
              'X: Close program \n'
              )
        result = input('\nInput: ')
        if result.upper() not in ['P', 'B', 'XBOX', 'PS4', 'BG', 'S','Y', 'D', 'M','X', 'POOL']:
            print('Please enter valid input')
        else:
            return result.upper()
    
def equipment_selection() -> str:
        while(True):
            print('Please enter a letter to make a selection:\n\n',
              'A: Add equipment use\n',
              'T: Track equipment use\n',
              'R: Returned equipment use \n',
              'B: Go Back \n',
              'X: Close program \n'
              )
            result = input('\nInput: ')
            if result.upper() not in ['A', 'T', 'R', 'B', 'X']:
                print('Please enter valid input')
            else:
                return result.upper()



def equipment_add_ui(equip: str, file:'file',filename:str, coord: Coordinate):
    name = input('\nResident name: ')
    hall = input('Hall: ')
    #game = input('Game: ')
    quantity = input('Quantity: ')
    if equip == 'Ping Pong' or equip == 'Billards':
        xcelreader.add_usage_pb(file,filename, equip, coord, name, hall, quantity,
                                datetime.datetime.now())
    elif equip == 'XBOX' or equip == 'PS4':
        xcelreader.add_usage_vg(file,filename, equip, coord, datetime.datetime.now(), name, hall, quantity,
                                game, datetime.datetime.now())
    
    
    



def liability_selection() -> str:
    while(True):
        print('Please enter a letter to make a selection:\n\n',
              'A: Add resident to file\n',
              'L: Look up resident\n',
              'R: Remove resident \n',
              'B: Go Back \n',
              'X: Close program \n'
              )
        result = input('\nInput: ')
        if result.upper() not in ['A', 'L', 'R','B','X']:
            print('Please enter valid input')
        else:
            return result.upper()

def liability_type() -> str:
    while(True):
        print('Please enter a letter to choose a liability form:\n\n',
              'G: General\n',
              'GAME: Game\n',
              'D: DVD \n',
              'B: Go Back \n',
              'X: Close program \n'
              )
        result = input('\nInput: ')
        if result.upper() not in ['G', 'GAME', 'D','B','X']:
            print('Please enter valid input')
        else:
            return result.upper()





def input_handler(file:'files', filename:str, coords: {str: Coordinate}):
    close = False
    while(not close):
        first = first_selection()
        if first == 'E':
            while(True):
                equip_sheet = equipment_sheet()
                if equip_sheet == 'P':
                    current_sheet = 'Ping Pong'
                    equip = equipment_selection()
                    if equip == 'A':
                        equipment_add_ui(current_sheet,file, filename,coords[current_sheet]) 
                elif equip_sheet == 'POOL':
                    current_sheet = 'Billards'
                    equip = equipment_selection()
                elif equip_sheet == 'XBOX':
                    current_sheet = 'XBOX'
                    equip = equipment_selection()
                elif equip_sheet == 'PS4':
                    current_sheet = 'PS4'
                    equip = equipment_selection()
                elif equip_sheet == 'BG':
                    current_sheet = 'Board Games'
                    equip = equipment_selection()
                elif equip_sheet == 'S':
                    current_sheet = 'Sports Balls'
                    equip = equipment_selection()
                elif equip_sheet == 'Y':
                    current_sheet = 'Yoga Mats'
                    equip = equipment_selection()
                elif equip_sheet == 'D':
                    current_sheet = 'DVDS'
                    equip = equipment_selection()
                elif equip_sheet == 'M':
                    current_sheet = 'Miscellaneous'
                    equip = equipment_selection()
                elif equip_sheet == 'B':
                    break
                elif equip_sheet == 'X':
                    return










    
        elif first == 'L':
            while(True):
                liability_sheet = liability_type()
                if liability_sheet == 'G':
                    current_sheet = 'General'
                elif liability_sheet == 'GAME':
                    current_sheet = 'Technology'
                elif liability_sheet == 'D':
                    current_sheet = 'DVD'
                elif liability_sheet == 'B':
                    break
                elif liability_sheet == 'X':
                    return
        elif first == 'X':
            return

        
            

def set_file(file:bool) -> 'file type compatible with openpyxl':
    while(True):
        if(file):
            filename = input('Please enter file you would like to write to for equipment: ')
        else:
            filename = input('Please enter file you would like to write to for liability forms: ')
        try:
            return (openpyxl.load_workbook(filename+'.xlsx'),filename+'.xlsx')
        except:
            print('\nexcel file does not exist. Would you like to make this into a new one?')
            while(True):
                response = input('\n Y: Yes\n N: No\n Input: ')
                if 'N' in response.upper():
                    break
                elif 'Y' in response.upper():
                    xcelwriter.create_rec_sheet(filename+'.xlsx')
                    return (openpyxl.load_workbook(filename+'.xlsx'),filename+'.xlsx')
            continue
            

 
    
def find_bottoms(file:'file')->dict:
    result = {}
    sheets = file.get_sheet_names()
    for i in sheets:
        result[i] = xcelreader.find_bottom(file, i)
    return result
        
    



if __name__ == '__main__':
    equipment_file_list = set_file(True)
    equipment_file = equipment_file_list[0]
    equipment_file_name = equipment_file_list[1]
    #liability_file = set_file(False)

    equipment_sheets = find_bottoms(equipment_file)
    #liability_sheets = find_bottoms(liability_file)
    

    
    Ca_name = input('Please enter your name: ')
    print('Hello ' + Ca_name.split(' ')[0] + '! What would you like to do?\n')





    while(True):
        input_handler(equipment_file,equipment_file_name, equipment_sheets)
        break
        
        
    

