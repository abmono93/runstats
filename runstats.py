import csv, time, gspread, os
from oauth2client.service_account import ServiceAccountCredentials

"""uncomment this function to update cells individually
def my_update_cell(collumn_to_write, data_to_write):
	written = False
	while not written:
		try:
			worksheet.update_cell(last_row, collumn_to_write, data_to_write)
			written = True
		except gspread.exceptions.APIError:
			print("Waiting for Google's API call limit to reset...")
			time.sleep(100)"""

# get the original info from the Activities.csv file
rows = []
with open('../Downloads/Activities.csv', 'rb') as csvfile:
	filereader = csv.reader(csvfile)
	for row in filereader:
		rows.insert(0, row)
os.remove('../Downloads/Activities.csv')
		
# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
	 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('secretfilename.json', scope)
gc = gspread.authorize(credentials)

# use client to open the worksheet and pull the dates that have been added
sh = gc.open("RunningStats")
worksheet = sh.worksheet("Raw_Data")
dates_list = worksheet.col_values(1)

# go through the rows from the csv file and list which ones aren't on the worksheet yet
to_add = []
for row in rows:
	if row[0].lower() == "running":
		if row[1][11] == '0':
			date = row[1][:11] + row[1][12:]
		else:
			date = row[1]
		row[1] = date
		if date not in dates_list:
			to_add.append(row)
		"""use this to see cells updated one by one
			my_update_cell(1, date)
			my_update_cell(3, row[4])
			my_update_cell(4, row[6])
			my_update_cell(5, row[11])"""

# see how many rows need updating, then put the information in the cells locally and
# batch update the worksheet
last_row = len(dates_list)
num_rows_to_add = len(to_add)
if num_rows_to_add > 0:
	s = ""
	if num_rows_to_add > 1: s = "s"
	print("Adding " + str(num_rows_to_add) + " row" + s + " to spreadsheet...")
	cells_range = "A" + str(last_row + 1) + ":E" + str(last_row + num_rows_to_add)
	cell_list = worksheet.range(cells_range)
	# keep track of how many rows have been finished
	i = 0
	# add info to cells one by one based on the collumn and which row we're on
	for cell in cell_list:
		if cell.col == 1:
			cell.value = to_add[i][1]
		elif cell.col == 2:
			#this cell doesn't use anything from the csv file, just the row number
			row_num = str(cell.row)
			cell.value = '=DATE(LEFT(A'+row_num+',4), RIGHT(LEFT(A'+row_num+', 7), 2), RIGHT(LEFT(A'+row_num+', 10), 2))'
		elif cell.col == 3:
			cell.value = to_add[i][4]
		elif cell.col == 4:
			cell.value = to_add[i][6]
		elif cell.col == 5:
			cell.value = to_add[i][11]
			i += 1
	worksheet.update_cells(cell_list, value_input_option = 'USER_ENTERED')
else:
	print("No new rows found to add to spreadsheet")
			
			
