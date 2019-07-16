import pandas as pd
import json
import xlsxwriter

try:
    import pandas as pd
except:
    raise exceptions.ValidationError('Warning ! python-pandas module missing. Please install it.')
# if file
inv_file  = 'Accounts.json'
with open(inv_file) as inv_file:
    dict_inv_file = json.load(inv_file)
inv_frame = pd.DataFrame.from_dict(dict_inv_file, orient='index')
inv_frame.reset_index(inplace=True)

# else no file

index = []
columns  =[]
inv_frame = pd.DataFrame(index=index, columns=columns)

loc_file = 'ceaarnongst-Locations-export.json'
with open(loc_file) as loc:
    dict_location_file = json.load(loc)
location_frame = pd.DataFrame.from_dict(dict_location_file, orient='index')
location_frame.reset_index(inplace=True)

customer_file = 'ceaarnongst-Customers-export.json'
with open(customer_file) as coust:
    dict_coustmer_file = json.load(coust)
customer_frame = pd.DataFrame.from_dict(dict_coustmer_file, orient='index')
customer_frame.reset_index(inplace=True)

product_file  = 'ceaarnongst-Products-export.json'
with open(product_file) as pro:
    dict_product_file = json.load(pro)
product_frame = pd.DataFrame.from_dict(dict_product_file, orient='index')
product_frame.reset_index(inplace=True)


with pd.ExcelWriter('Excel Master.xlsx',  engine='xlsxwriter') as writer:  # doctest: +SKIP
	inv_frame.to_excel(writer,sheet_name='Invoices',index=False)
	location_frame.to_excel(writer, sheet_name='Location',index=False)
	customer_frame.to_excel(writer,sheet_name='Customers',index=False)
	product_frame.to_excel(writer,sheet_name='Products',index=False)
	workbook  = writer.book
	worksheet1 = writer.sheets['Invoices']
	worksheet2 = writer.sheets['Location']
	worksheet3 = writer.sheets['Customers']
	worksheet4 = writer.sheets['Products']
	
	# worksheet.protect()
	# unlocked = workbook.add_format({'locked': 0})
	bold = workbook.add_format({'bold': True})

	# worksheet1.write(0,0,'Invoices',bold)
	worksheet2.write(0,0,'Location',bold)
	worksheet3.write(0,0,'Customers',bold)
	worksheet4.write(0,0,'Products',bold)
	
	worksheet1.set_column('A:A', 25)
	worksheet1.set_column('B:E', 12)

	
	worksheet2.set_column('A:A', 30)
	worksheet2.set_column('B:K', 16)

	worksheet3.set_column('A:T', 18)

	worksheet4.set_column('A:A', 30)
	worksheet4.set_column('B:AA', 20)
	
	# To close the workbook
	workbook.close()
