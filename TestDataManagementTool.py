from tkinter import *
from tkinter import ttk #for button
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import pypyodbc
import xlrd
root = Tk()
##root.configure(background='blue')

root.title("TDM Tool Stores")
root.iconbitmap('bbylogo_AEu_icon.ico')
Label(root, text = "Enter SKU").pack() #pack is the geomectry manager

SKUReceived = StringVar() 
SkuText = ttk.Entry(root,textvariable=SKUReceived)
SkuText.pack(padx = 10, pady = 10)
Label(root, text = "OR").pack()
buttonBulkUpload = ttk.Button(root,text = "Browse file for SKUs...")
buttonBulkUpload.pack(padx = 10, pady = 10)

def bulkUpload():
    root.fileName = filedialog.askopenfilename(filetypes =( ("Excel files", "*.xls;*.xlsx"),("All Files","*.*")))
    path = 'SKU List.xlsx'
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_index(0)
    #cell = worksheet.cell(0,0)
    print (worksheet.nrows)
    x= []
    for rownum in range(worksheet.nrows):
        x.append(worksheet.cell(rownum, 0))
    stringSKU = str(x)
    SKUlistFill = stringSKU.replace('number:','')
    SKUlistFill2 = str(SKUlistFill.replace('.0',''))
    SKUlistFill3 = str(SKUlistFill2.replace(']',''))
    SKUlistFill4 = str(SKUlistFill3.replace('[',''))
    print(SKUlistFill4)
    SKUReceived.set(SKUlistFill4)


#Store Server
Label(root, text="Enter the Server(Format- BBY01990S001,63518)").pack(side=TOP)
serverlisted = StringVar()
ServerName = ttk.Entry(root, textvariable=serverlisted)
ServerName.pack(padx = 10, pady = 10)

#Disposition Combo Box
Label(root, text = "Select Disposition").pack()
dispositionsListed = StringVar()
comboboxDisp = ttk.Combobox(root, textvariable = dispositionsListed)
comboboxDisp.pack(padx = 10, pady = 10)
comboboxDisp.config(value = ('1.2', '2.1', '2.2')) #put all the disposition values avialable

#Entitlement Combo Box
Label(root, text = "Select Entitlement").pack()
entitlementListed = StringVar()
comboboxEnt = ttk.Combobox(root, textvariable = entitlementListed)
comboboxEnt.pack(padx = 10, pady = 10)
comboboxEnt.config(value = ('M42.5 & M64-RTV', '2a', '3a', '4a')) #put all the entitlment values

#Button
button = ttk.Button(root,text = "Execute")
button.pack(padx = 10, pady = 10)

#Button2
button2 = ttk.Button(root,text = "Status")
button2.pack(padx = 10, pady = 10)


#SKU TextBox
#SKURecieved = StringVar()
#ttk.TextBox(root, textvariable = )

def SelectQuery():
    va1 = str(serverlisted.get())
    try:
        va1 = str(serverlisted.get())
        connection = pypyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=POSFDN001;Trusted_Connection=yes' %(va1)) # Creating a windows authentication connection
        print('connected')
        cur = connection.cursor()
        var = str(SKUReceived.get())
        sql= "Select * from ITEMRETURNS WHERE ITEMID in ( %s )" %(var) 
        cur.execute(sql)
        results = cur.fetchone() # Converted the Array output to String.
        var1= str(results[0])
        var2= str(results[1])
        var3= str(results[2])
##        print(results)
        connection.close()
    except:
        print('Issue')
    Label(root, text= 'Itemid=' + var1 + ', ' + 'MediaDCProcessFlag=' + var2 + ', '+ 'RTValwaysFlag='+ var3 + ', ').pack(side=RIGHT)

#Execution Function
def SampleQuery():
   
    if (dispositionsListed.get() == '1.2'):
        try:
            va2 = str(serverlisted.get())
            connection = pypyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=POSFDN001;Trusted_Connection=yes' %(va2))
            print('connected')
            cur = connection.cursor()
            var = str(SKUReceived.get())
            sql= "UPDATE ItemReturns SET ReturnableFlag = 'True' , ResellFlag='True' WHERE ITEMID in ( %s )" %(var)
            cur.execute(sql)
            cur.commit()
            cur.fetchall
            connection.close()
        except:
            print('Connection Issue')
    
##    else:
##        print('Program Failed')
    elif entitlementListed.get() == 'M42.5 & M64-RTV':
        try:
            connection = pypyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=POSFDN001;Trusted_Connection=yes' %(va2))
            print('connected')
            cur = connection.cursor()
            var = str(SKUReceived.get())
            sql= "UPDATE dbo.ItemReturns SET RTVopenBoxDayQty = 9999, RTVdefectiveDayQty = 9999, ServiceFactoryWarrantyInFlag = 'True', ServiceFactoryWarrantyOutFlag = 'False', SupportBBfactoryWarrantyFlag = 'False', ReturnPolicyDayQty = -30, RapidExchangeMfgWarrantyFlag = 'False', RapidExchangeCallVendorFlag = 'False', RapidExchangeUnderPSPflag = 'False', OpenBoxReturnFlag = 'True', ReturnableFlag = 'True', RapidExchangeUnderPRPflag = 'False', RapidExchangeFeeFlag = 'False' where ItemID in ( %s)" %(var)
            cur.execute(sql)
            cur.commit()
            cur.fetchall
            connection.close()
        except:
            print('Connection Issue')
    else:
	    print("Invalid Option")
    Label(root, text = "EXECUTED!!!").pack()
##
##
##	if entitlementListed.get() == '1a':
##		print('1.1')
##	elif entitlementListed.get() == '2a':
##		print('1.2')
##	else:
##		print("Invalid Option")
##	print("Executed")
##	Label(root, text = "EXECUTED!!!").pack()
##
def DialogueBox():
    messagebox.showinfo('Dialogue','Are you sure you want to Run an Update?')
    SampleQuery()
button.config(command = DialogueBox)
button2.config(command = SelectQuery)
buttonBulkUpload.config(command = bulkUpload)
