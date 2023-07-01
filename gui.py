from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
import xlsxwriter
from os.path import expanduser

def new_func7():
    date = []
    amount = []
    bill_list = []
    bill_start = []
    file_name = []
    date_1 = []
    return date,amount,bill_list,bill_start,file_name,date_1

def show_entry_fields(event=None):
    date.append(date_entry.get())
    amount.append(float(amount_entry.get()))
    bill_list.append(int(bill_entry.get()))
    v.set(bill_list[-1]+1)
    amount_entry.delete(0, 'end')

def from_to(event=None):
    file_name.append(file_entry.get())
    date_1.append(from_entry.get())
    date_1.append(to_entry.get())
    bill_start.append(int(bill_start_entry.get()))
    v.set(bill_start[0])
    date_entry.insert(0, date_1[0])

def generate_file():
    home = expanduser("~")
    ############  creating workbook and worksheet ##############
    workbook = xlsxwriter.Workbook(f'{home}/Desktop/{file_name[0]}.xlsx')
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
    worksheet.merge_range('A1:H1', 'M/S. BADRUDDIN MULLA SHAMSUDDIN AND SONS', cell_format)
    worksheet.merge_range('A2:H2', 'SADAR BAZAR, RAIPUR', cell_format)
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




    for i in range(0, len(bill_list)):
        worksheet.write(row, col, sr, cell_format)
        worksheet.write(row, col+1, date[i], date_format)
        worksheet.write(row, col+2, bill_list[i], cell_format)
        worksheet.write(row, col+3, f'=ROUND({amount[i]},2)', num_format)
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
    messagebox.showinfo("Success", "File Generated Successfully")
    #########################################################################

def new_func5():  #window creation and styling 
    master = Tk()
    master.geometry("1000x500+250+150")
    master.title("GST Filling")
    master.config(bg='#161a1d')
    return master

def set_style():
    style = Style()
    style_label = Style()
    style.configure('TButton', font = ('calibri', 14, 'bold'), borderwidth = '4', background="#161a1d", foreground='#000000')
    style_label.configure('BW.TLabel', font=('Times', 18), background="#161a1d", foreground='#f5f3f4')

def set_label(master):
    Label(master, text="File Name", style='BW.TLabel').grid(row=0, column=0, padx=(50,20), pady=(25,5))
    Label(master, text="From", style='BW.TLabel').grid(row=1, column=0, padx=(50,20), pady=(25,5))
    Label(master, text="To", style='BW.TLabel').grid(row=2, column=0, padx=(50,20), pady=(25,5))
    Label(master, text="Bill Start", style='BW.TLabel').grid(row=3, column=0, padx=(50,20), pady=(25,5))
    Label(master, text="Date", style='BW.TLabel').grid(row=0, column=2, padx=(30,20), pady=(25,5))
    Label(master, text="Bill No", style='BW.TLabel').grid(row=1, column=2, padx=(30,20), pady=(25,5))
    Label(master, text="Amount", style='BW.TLabel').grid(row=2, column=2, padx=(30,20), pady=(25,5))

def set_entry_field(master, v):
    file_entry = Entry(master, font=('Times', 12))
    from_entry = Entry(master, font=('Times', 12))
    to_entry = Entry(master, font=('Times', 12))
    bill_start_entry = Entry(master, font=('Times', 12))
    date_entry = Entry(master, font=('Times', 12))
    amount_entry = Entry(master, font=('Times', 12))
    bill_entry = Entry(master, text=v, font=('Times', 12))
    return file_entry,from_entry,to_entry,bill_start_entry,date_entry,amount_entry,bill_entry

def set_grid(file_entry, from_entry, to_entry, bill_start_entry, date_entry, amount_entry, bill_entry):
    file_entry.grid(row=0, column=1, padx=(5,20), pady=(25,5), ipady=3)
    from_entry.grid(row=1, column=1, padx=(5,20), pady=(25,5), ipady=3)
    to_entry.grid(row=2, column=1, padx=(5,20), pady=(25,5), ipady=3)
    bill_start_entry.grid(row=3, column=1, padx=(5,20), pady=(25,5), ipady=3)
    date_entry.grid(row=0, column=3, padx=(5,20), pady=(25,5), ipady=3)
    bill_entry.grid(row=1, column=3, padx=(5,20), pady=(25,5), ipady=3)
    amount_entry.grid(row=2, column=3, padx=(5,20), pady=(25,5), ipady=3)

def set_buttons(show_entry_fields, from_to, generate_file, master): #Buttons
    Button(master, text='Set', command=from_to, style='TButton').grid(row=8, column=0, sticky=W, padx=(75, 50),pady=(40,10), ipadx=7, ipady=5)
    Button(master, text='Quit', command=master.quit).grid(row=8, column=1, sticky=W, padx=(50, 50),pady=(40,10), ipadx=7, ipady=5)
    Button(master, text='Enter', command=show_entry_fields).grid(row=8, column=3, sticky=W, padx=(50, 50),pady=(40,10), ipadx=7, ipady=5)
    Button(master, text='Generate File', command=generate_file).grid(row=8, column=2, sticky=W, padx=(50, 50) ,pady=(40,10), ipadx=7, ipady=5)

def binding_keys(show_entry_fields, from_to, master):
    master.bind('<Return>', show_entry_fields)
    master.bind('<Control-s>', from_to)


if __name__ == "__main__":
    date, amount, bill_list, bill_start, file_name, date_1 = new_func7()
    master = new_func5()
    set_style()
    set_label(master)
    v = IntVar() 
    file_entry, from_entry, to_entry, bill_start_entry, date_entry, amount_entry, bill_entry = set_entry_field(master, v)
    set_grid(file_entry, from_entry, to_entry, bill_start_entry, date_entry, amount_entry, bill_entry)
    set_buttons(show_entry_fields, from_to, generate_file, master)
    binding_keys(show_entry_fields, from_to, master)

    mainloop( )
