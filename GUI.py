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
differentials = []

# def show_warning():
# 	print("Error")

def testCommand():
	print(2 + 2)


# Iterates through the passed dailyChange list, building a string of output to
# display streaks
def findStreakDifferential():
	streakLength = eStreak.get()
	minValue = eMin.get()

	minV = float(minValue)
	streak = int(streakLength)

	streakDates = calc_Streak(minV, streak)

	print('Dates where the differential ("^gspc" daily change - ' +
		'daily account" daily change) was \nhigher than ' + 
			minValue + '% for ' + streakLength + ' or more days in a row:\n')
	print(streakDates)
	sys.stdout.close()

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

def load_differentials(gspcDailyChange, davtDailyChange):
	for i in range(2143):
		differentials.append(gspcDailyChange[i] - davtDailyChange[i])

# Call all the load functions
def load_everything(worksheet):
	load_dates(ws)
	load_gspc_close(ws)
	load_gspc_daily_change(ws)
	load_davt_close(ws)
	load_davt_daily_change(ws)
	load_differentials(gspcDailyChangeList, davtDailyChangeList)
	# for x in differentials:
	# 	print(x)

def calc_Streak(minV, streak):
	streakDates = ""
	runningDates = ""
	runningStreakCounter = 0
	index = 0

	for val in differentials:
		index += 1
		if val >= minV:
			# Save the date at the index point in a temp var
			runningDates += dateList[index] + '\n'
			runningStreakCounter += 1
		else:
			# Streak is broken. Save it to streakDates (what gets printed
			# 	later) and reset temp vars
			if runningStreakCounter >= streak:
				streakDates += runningDates + '\n'
			runningDates = ""
			runningStreakCounter = 0

	return streakDates


# Load the arrays with data
load_everything(ws)

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
	width = 1, height = 1, command = lambda: findStreakDifferential())
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


# Set up the two radio buttons
# chooseChart_label = tk.Label(root, text = "Choose your chart:")
# chooseChart_label.grid(row = 0, column = 0, padx = 10, 
# 	pady = (10, 5), sticky = "W")

# gspcRadio = tk.Radiobutton(root, text = 'gspc', variable = v,
# 	value = 1, anchor = "w")
# davtRadio = tk.Radiobutton(root, text = 'davt', variable = v, 
# 	value = 2, anchor = "w")

# gspcRadio.grid(row = 1, column = 0)
# davtRadio.grid(row = 1, column = 1)


root.mainloop()