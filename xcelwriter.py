import xlsxwriter

'''
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx',{'constant_memory':True})
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet('test')

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Hello')

# Text with formatting.
worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

workbook.read('A:A')

workbook.close()
'''


def create_rec_sheet(file_name: str):
    workbook = xlsxwriter.Workbook(file_name, {'constant_memory':True})
    pingpong = workbook.add_worksheet('Ping Pong').set_tab_color('red')
    billards = workbook.add_worksheet('Billards').set_tab_color('orange')
    xbox = workbook.add_worksheet('XBOX').set_tab_color('lime')
    ps4 = workbook.add_worksheet('PS4').set_tab_color('blue')
    boardGames = workbook.add_worksheet('Board Games').set_tab_color('purple')
    sportsBalls = workbook.add_worksheet('Sports Balls').set_tab_color('yellow')
    yogaMats = workbook.add_worksheet('Yoga Mats').set_tab_color('#AFEEEE')
    dvds = workbook.add_worksheet('DVDS').set_tab_color('brown')
    miscellaneous = workbook.add_worksheet('Miscellaneous').set_tab_color('pink')
    workbook.close()



