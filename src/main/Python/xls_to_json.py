import xlrd
import re
from collections import OrderedDict
import simplejson as json
import time
i=0
failed=[]
print ("________________________________________________________________________")
print ("This program would convert you XML document into Individual JSON files!")
print ("________________________________________________________________________")
print ("")
print ("")
print ("")
user = input("Please press y/Y to start conversion: ")

if(user=="y" or user=="Y"):
    # Open the workbook and select the first worksheet
    try:
        wb = xlrd.open_workbook('Billing_Questions.xls')
        sh = wb.sheet_by_index(0)
        print ("Gathering Questions...")
        time.sleep(1)
        
        # List to hold dictionaries
        title_list = []
        # Iterate through each row in worksheet and fetch values into dict
        for rownum in range(1, sh.nrows):
            title = OrderedDict()
            row_values = sh.row_values(rownum)
   #        title['id'] = int(row_values[0])
            title['title'] = row_values[0]
            title['body'] = row_values[1]
            title['textHtml'] = row_values[2]
            title_list.append(title)   
            j = json.dumps(title_list)
            parsed = json.loads(j)
            p=json.dumps(parsed, indent=4, sort_keys=True)
            p=p[6:]
            p=p[:-2]
            name=p
            name.replace("?"," ")
            f= open("Billing_" + str(row_values[0])+'.json',"w+")
            f.write(p)
            f.close()
            title_list=[]
            i=i+1
    except:
        failed.append(row_values[0])
print("")
print ("converted " + str(i) +" files out of " + str(rownum)+". Thank you" )
info=input("please press i for more info or else press e for exit:")
if(info=="i" or info=="I"):
    print ("")
    print("following test titles failed:")
    if(len(failed) == 0):
        print ("Null")
    else:
        print ("")
        print(failed)  
else:
    print("You may now close the program! Thank you for using me")


