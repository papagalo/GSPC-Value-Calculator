import tkinter as tk
import pdb
import sys
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

# Create a workbook object
wb = Workbook()

sys.stdout = open('gspcVSdavt results.txt', 'w')

# load existing spreadsheet
wb = load_workbook(filename='GSPCdateAsText.xlsx', data_only=True)
ws = wb.active
root = tk.Tk()
root.title("GvD Calculator")
root.geometry("300x225")
v = tk.IntVar()

# Declare ranges to load lists
date_range = ws['A2':'A2145']
gspc_close_range = ws['B2':'B2145']
gspc_dc_range = ws['C3':'C2145']
davt_close_range = ws['E2':'E2145']
davt_dc_range = ws['F3':'F2145']

# Declare lists
dateList = []
gspcCloseList = []
gspcDailyChangeList = []
davtCloseList = []
davtDailyChangeList = []

# def show_warning():
# 	print("Error")

def testCommand():
	print(2 + 2)


# Iterates through the passed dailyChange list, building a string of output to
# display streaks
def findStreaksDailyChange(dates, gspcDC, davtDC):
	# pdb.set_trace()
	streakLength = eStreak.get()
	minValue = eMin.get()

	# try:
	#if streakLength != '':
	if v.get() == 1:
		minV = float(minValue)
		streak = int(streakLength)
		# except ValueError:
		# show_warning()
		# return false

		streakDates = ""
		runningDates = ""
		runningStreakCounter = 0
		index = 0

		for val in gspcDC:
			index += 1
			if val >= minV:
				runningDates += dates[index] + '\n'
				runningStreakCounter += 1
			else:
				if runningStreakCounter >= streak:
					streakDates += runningDates + '\n'
				runningDates = ""
				runningStreakCounter = 0
	
		print('Dates where the gspc daily change was higher than ' + 
			minValue + '% for ' + streakLength + ' or more days in a row:\n')
		print(streakDates)
		sys.stdout.close()
	else:
		minV = float(minValue)
		streak = int(streakLength)
		# except ValueError:
		# show_warning()
		# return false

		streakDates = ""
		runningDates = ""
		runningStreakCounter = 0
		index = 0
	
		for val in davtDC:
			index += 1
			if val >= minV:
				runningDates += dates[index] + '\n'
				runningStreakCounter += 1
			else:
				if runningStreakCounter >= streak:
					streakDates += runningDates + '\n'
				runningDates = ""
				runningStreakCounter = 0
		print('Dates where the davt daily change was higher than ' + 
			minValue + '% for ' + streakLength + ' or more days in a row:\n')
		print(streakDates)
		sys.stdout.close()

	# if v.get() == 1:
	# 	print("1")
	# elif v.get() == 2: 
	# 	print("2")


# "Load" functions load their respective lists with data from the Excel sheet
def load_dates(worksheet):
	for cell in date_range:
		for x in cell:
			dateList.append(x.value)
	
def load_gspc_close(worksheet):
	for cell in gspc_close_range:
		for x in cell:
			gspcCloseList.append(x.value)

def load_gspc_daily_change(worksheet):
	for cell in gspc_dc_range:
		for x in cell:
			gspcDailyChangeList.append(x.value * 100)

def load_davt_close(worksheet):
	for cell in davt_close_range:
		for x in cell:
			davtCloseList.append(x.value)

def load_davt_daily_change(worksheet):
	for cell in davt_dc_range:
		for x in cell:
			davtDailyChangeList.append(x.value * 100)

def load_everything(worksheet):
	load_dates(ws)
	load_gspc_close(ws)
	load_gspc_daily_change(ws)
	load_davt_close(ws)
	load_davt_daily_change(ws)


# Begin the actual program now

# Call all the load functions
load_everything(ws)


# Set up the two radio buttons
chooseChart_label = tk.Label(root, text = "Choose your chart:")
chooseChart_label.grid(row = 0, column = 0, padx = 10, 
	pady = (10, 5), sticky = "W")

gspcRadio = tk.Radiobutton(root, text = 'gspc', variable = v,
	value = 1, anchor = "w")
davtRadio = tk.Radiobutton(root, text = 'davt', variable = v, 
	value = 2, anchor = "w")

gspcRadio.grid(row = 1, column = 0)
davtRadio.grid(row = 1, column = 1)


# Set up the min daily change label and entry box
minDailyChange_label = tk.Label(root, text = "Minimum Daily Change:")
minDailyChange_label.grid(row = 2, column = 0, padx = 10, pady = 10)

eMin = tk.Entry(root, width = 10, borderwidth = 4)
eMin.grid(row = 2, column = 1, padx = 10, pady = 10, columnspan = 2)


# Set up the min daily change label and entry box
streak_label = tk.Label(root, text = "How many days in a row?")
streak_label.grid(row = 3, column = 0, padx = 10, pady = 10)

eStreak = tk.Entry(root, width = 10, borderwidth = 4)
eStreak.grid(row = 3, column = 1, padx = 10, pady = 10, columnspan = 2)

submitButton = tk.Button(root, text = "Submit", padx = 40, pady = 20, 
	width = 1, height = 1, command = lambda: findStreaksDailyChange(dateList, 
		gspcDailyChangeList, davtDailyChangeList))
submitButton.grid(row = 4, column = 0)


# Create an exit button that kills the program
exitButton = tk.Button(root, text = "Exit", padx = 40, pady = 20, 
	width = 1, height = 1, command = root.destroy)
exitButton.grid(row = 4, column = 1)





# # this will create a label widget
# l1 = tk.Label(root, text = "Height")
# l2 = tk.Label(root, text = "Width")
  
# # grid method to arrange labels in respective
# # rows and columns as specified
# l1.grid(row = 0, column = 0, sticky = "W", pady = 2)
# l2.grid(row = 1, column = 0, sticky = "W", pady = 2)
  
# # entry widgets, used to take entry from user
# e1 = tk.Entry(root)
# e2 = tk.Entry(root)
  
# # this will arrange entry widgets
# e1.grid(row = 0, column = 1, pady = 2)
# e2.grid(row = 1, column = 1, pady = 2)










# defines the submit command used for myButton below
# def submit():
# 	greet = nameit(myBox.get())
# 	myLabel.config(text=greet)

# myBox = Entry(root)
# myBox.pack(pady=20)


# myLabel = Label(root, text="", font=("Helvetica", 18))
# myLabel.pack(pady=20)

# myButton = Button(root, text="Submit Name", command = submit)
# myButton.pack(pady=20)

root.mainloop()