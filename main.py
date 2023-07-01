from datetime import datetime
import xlsxwriter
############### Input to file name #########################

file_name[0] = file_name[0] + '.xlsx'


############################################################





##############################################################

############  creating workbook and worksheet ##############
workbook = xlsxwriter.Workbook(file_name[0])
worksheet = workbook.add_worksheet('Data')

###########################################################


############################ Formats #############################
cell_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'})

num_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'num_format': '###0.00'})


date_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter'})
#######################################################################



############################ file header format   ##########################################
worksheet.set_column('B:B', 15)
worksheet.merge_range('A1:H1', 'client name', cell_format)
worksheet.merge_range('A2:H2', 'address', cell_format)
worksheet.merge_range('A3:H3', 'SALES DETAILS', cell_format)
worksheet.merge_range('A4:H4', f'FROM {date_1[0]} T0 {date_1[1]}', cell_format)
worksheet.merge_range('A5:H5', 'GST NUMBER :', cell_format)
worksheet.write('A6', 'SL.', cell_format)
worksheet.write('B6', 'DATE', cell_format)
worksheet.write('C6', 'INVOICE', cell_format)
worksheet.write('D6', 'GROSS', cell_format)
worksheet.write('E6', 'SALE ', cell_format)
worksheet.write('F6', 'SGST', cell_format)
worksheet.write('G6', 'CGST', cell_format)
worksheet.write('H6', 'TOTAL', cell_format)
worksheet.write('A7', 'NO.', cell_format)
worksheet.write('B7', 'OF ', cell_format)
worksheet.write('C7', 'NUMBER', cell_format)
worksheet.write('D7', 'SALES', cell_format)
worksheet.write('E7', 'AMOUNT', cell_format)
worksheet.write('F7', '9%', cell_format)
worksheet.write('G7', '9%', cell_format)
worksheet.write('H7', 'SALES', cell_format)
worksheet.write('A8', '', cell_format)
worksheet.write('B8', 'SALE', cell_format)
worksheet.write('C8', '', cell_format)
worksheet.write('D8', 'IN', cell_format)
worksheet.write('E8', 'IN', cell_format)
worksheet.write('F8', 'IN', cell_format)
worksheet.write('G8', 'IN', cell_format)
worksheet.write('H8', 'IN', cell_format)
worksheet.write('A9', '', cell_format)
worksheet.write('B9', '', cell_format)
worksheet.write('C9', '', cell_format)
worksheet.write('D9', 'RUPEES', cell_format)
worksheet.write('E9', 'RUPEES', cell_format)
worksheet.write('F9', 'RUPEES', cell_format)
worksheet.write('G9', 'RUPEES', cell_format)
worksheet.write('H9', 'RUPEES', cell_format)
worksheet.write('A10', '', cell_format)
worksheet.write('B10', '', cell_format)
worksheet.write('C10', '', cell_format)
worksheet.write('D10', '', cell_format)
worksheet.write('E10', '84.74576%', cell_format)
worksheet.write('F10', '9%', cell_format)
worksheet.write('G10', '', cell_format)
worksheet.write('H10', '', cell_format)
worksheet.write('A11', '', cell_format)
worksheet.write('B11', '', cell_format)
worksheet.write('C11', '', cell_format)
worksheet.write('D11', '', cell_format)
worksheet.write('E11', 'ON-4', cell_format)
worksheet.write('F11', 'ON-5', cell_format)
worksheet.write('G11', 'ON-5', cell_format)
worksheet.write('H11', '', cell_format)
worksheet.write('A12', '(1)', cell_format)
worksheet.write('B12', '(2)', cell_format)
worksheet.write('C12', '(3)', cell_format)
worksheet.write('D12', '(4)', cell_format)
worksheet.write('E12', '(5)', cell_format)
worksheet.write('F12', '(6)', cell_format)
worksheet.write('G12', '(7)', cell_format)
worksheet.write('H12', '(8)', cell_format)


#########################################



############################ writing data to worksheet ##################


sr = 1
row = 12
col = 0




for i in range(0, len(bill_no)):
    mrp = input("Ammount: ")
    worksheet.write(row, col, sr, cell_format)
    worksheet.write(row, col+1, date[i], date_format)
    worksheet.write(row, col+2, bill_no[i], cell_format)
    worksheet.write(row, col+3, f'=ROUND({mrp[i]},2)', num_format)
    worksheet.write_formula(row, col+4, f'=SUM(D{row+1}*84.74576/100)', num_format)
    worksheet.write_formula(row, col+5, f'=SUM(E{row+1}*9/100)', num_format)
    worksheet.write_formula(row, col+6, f'=SUM(F{row+1})', num_format)
    worksheet.write_formula(row, col+7, f'=SUM(E{row+1}:G{row+1})', num_format)
    sr+=1
    row+=1


worksheet.write(row+2, col+1, "Total", cell_format)
worksheet.write(row+2, col+2, " = ", cell_format)
worksheet.write(row+2, col+3, f'=SUM(D13:D{row+1})', num_format)
worksheet.write(row+2, col+4, f'=SUM(E13:E{row+1})', num_format)
worksheet.write(row+2, col+5, f'=+E{row+3}*9/100', num_format)


workbook.close()
#########################################################################


