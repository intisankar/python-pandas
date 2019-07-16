import pandas as pd
import json
import xlsxwriter
# To check the pandas
try:
    import pandas as pd
except:
    raise exceptions.ValidationError('Warning ! python-pandas module missing. Please install it.')
# load file path
file = 'Accounts.json'
with open(file) as file:
    dict_train = json.load(file)
# converting json dataset from dictionary to dataframe
train = pd.DataFrame.from_dict(dict_train, orient='index')
train.reset_index(inplace=True)
# To delete unwanted data
desire = train.drop(columns=['doDisable','accountnumber','sortTimeStamp','runningtotal','isActive','isPaymentType'])
# to change the columns names
cols = list(desire.columns)
# to reorder the columns
a, b, c, d  = cols.index('groupName'), cols.index('openingbalance'), cols.index('acctNumber'), cols.index('ifsccode')
cols[1], cols[2],cols[3], cols[4] = cols[a], cols[b],cols[c], cols[d] 
desire = desire[cols]
# to write the xlsx writer object 
with pd.ExcelWriter('All account categories.xlsx',  engine='xlsxwriter') as writer:  # doctest: +SKIP
	desire.to_excel(writer, sheet_name='Sheet1',index=False)
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	# To enable the write protected
	worksheet.protect()
	# To unlock and bold format defination
	unlocked = workbook.add_format({'locked': 0})
	bold = workbook.add_format({'bold': True})
	row = 1
	col = 2
	# To apply the unlock for openingbalance column @ rewriting
	for item in range(len(desire['openingbalance'])):
		# to override Names of Indexes
		worksheet.write(0,0,'Account Name',bold)
		worksheet.write(0,1,'Group Name',bold)
		worksheet.write(0,2,'Opening Balance',bold)
		worksheet.write(0,3,'Account Number',bold)
		worksheet.write(0,4,'IFSC Code',bold)
		worksheet.write(row, col,desire['openingbalance'][item],unlocked)
		row += 1
	# To set the cell width
	worksheet.set_column('A:A', 30)
	worksheet.set_column('B:E', 16)
	# To close the workbook
	workbook.close()
