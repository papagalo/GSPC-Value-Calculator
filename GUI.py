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
root.geometry("400x225")
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

# Iterates through the passed dailyChange list, building a string of output to
# display streaks
def findStreakDifferential():
	streakLength = eStreak.get()
	minValue = eMin.get()
	futureDays = eFuture.get()

	minV = float(minValue)
	streak = int(streakLength)
	future_days = int(futureDays)

	streakDates = calc_Streak(minV, streak, future_days)

	print('Dates where the differential ("^gspc" daily change - ' +
		'daily account" daily change) was \nhigher than ' + 
			minValue + '% for ' + streakLength + ' or more days in a row:\n')
	print(streakDates)
	sys.stdout.close()
	root.destroy
	root.quit()

def calc_Streak(minV, streak, future_days):
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

# End definitions
# Begin program logic

# Load the arrays with data
load_everything(ws)

# Set up the GUI

# Set up the min daily change label and entry box
minDailyChange_label = tk.Label(root, text = "Minimum Daily Change:")
minDailyChange_label.grid(row = 1, column = 0, padx = 10, pady = 10)

eMin = tk.Entry(root, width = 10, borderwidth = 4)
eMin.grid(row = 1, column = 1, padx = 10, pady = 10, columnspan = 2)

# Set up the min daily change label and entry box
streak_label = tk.Label(root, text = "How many days in a row?")
streak_label.grid(row = 2, column = 0, padx = 10, pady = 10)

eStreak = tk.Entry(root, width = 10, borderwidth = 4)
eStreak.grid(row = 2, column = 1, padx = 10, pady = 10, columnspan = 2)

# Set up the min daily change label and entry box
future_label = tk.Label(root, text = "Display difference in % change over" +
	" next X days:")
future_label.grid(row = 3, column = 0, padx = 10, pady = 10)

eFuture = tk.Entry(root, width = 10, borderwidth = 4)
eFuture.grid(row = 3, column = 1, padx = 10, pady = 10, columnspan = 2)

# Create submit button
submitButton = tk.Button(root, text = "Submit", padx = 40, pady = 20, 
	width = 1, height = 1, command = lambda: findStreakDifferential())
submitButton.grid(row = 4, column = 0)

# Create an exit button that kills the program
exitButton = tk.Button(root, text = "Exit", padx = 40, pady = 20, 
	width = 1, height = 1, command = root.destroy)
exitButton.grid(row = 4, column = 1)


root.mainloop()