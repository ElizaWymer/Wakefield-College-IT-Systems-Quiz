import math
import tkinter
import tkinter.messagebox
from tkinter import *
import pandas

window = tkinter.Tk()
window.geometry("605x288")
window.resizable(width = False, height = False)
frame = Frame(window)
frame.grid()

banner = tkinter.Label(frame, text = "IT Systems Quiz Scoreboard Page 1", font = "calibri 18 bold", bg = "#99ccff", width = 46, height = 1)
banner.grid(columnspan = 4, row = 1, sticky = "w")

prevButton = tkinter.Button(frame, text = "Previous\nPage", font = "calibri 15 bold", command = "Previous Page", bg = "#97d077", width = 10)
prevButton.grid(column = 1, row = 8, sticky = "w", padx = (64, 0), pady = 12)

hubButton = tkinter.Button(frame, text = "Return to\nUser Hub", font = "calibri 15 bold", command = "Next Page", bg = "#f19c99")
hubButton.grid(column = 1, row = 8, sticky = "w", padx = (223, 0))

nextButton = tkinter.Button(frame, text = "Next\nPage", font = "calibri 15 bold", command = "", bg = "#97d077", width = 10)
nextButton.grid(column = 1, row = 8, sticky = "w", padx = (366, 0))

usernames = []
scores = []
rawTimes = []
displayTimes = []

spreadsheetName = "Wakefield College IT Systems Quiz Data.xlsx"
scoreboardDetails = pandas.read_excel(spreadsheetName, "Scoreboard")
scoreboardDataFrame = pandas.DataFrame(scoreboardDetails, columns = ["Username", "Score", "Raw Time", "Display Time"])

for x in scoreboardDataFrame["Username"]:
    usernames.append(str(x))

for x in scoreboardDataFrame["Score"]:
    scores.append(int(x))

for x in scoreboardDataFrame["Raw Time"]:
    rawTimes.append(int(x))

for x in scoreboardDataFrame["Display Time"]:
    displayTimes.append(str(x))

usernames = [x for _, x in sorted(zip(scores, usernames), reverse = True)]
rawTimes = [x for _, x in sorted(zip(scores, rawTimes), reverse = True)]
displayTimes = [x for _, x in sorted(zip(scores, displayTimes), reverse = True)]
scores = sorted(scores, reverse = True)

oc_set = set()
res = []
for idx, val in enumerate(scores):
    if val not in oc_set:
        for x in res:
            try:
                if rawTimes[x] < rawTimes[y]:
                    scores[y] -= 0.5
                elif rawTimes[y] < rawTimes[x]:
                    scores[x] -= 0.5
            except:
                pass
            y = x
        x = None
        y = None
        oc_set.add(val)
        res.clear()         
    res.append(idx)

usernames = [x for _, x in sorted(zip(scores, usernames), reverse = True)]
rawTimes = [x for _, x in sorted(zip(scores, rawTimes), reverse = True)]
displayTimes = [x for _, x in sorted(zip(scores, displayTimes), reverse = True)]
scores = sorted(scores, reverse = True)
for x in scores:
    scores[scores.index(x)] = round(x)

print("\n" + str(scores))
print(usernames)
print(rawTimes)
print(displayTimes)

numberOfPages = math.ceil(len(usernames) / 4)
print(numberOfPages)
page = 1
#if page == 1: previous button locked
#if page == number of pages: next button locked


values = [0, 1, 2, 3]

def PreviousPage():
    #Miscellaneous.DestroyFrame()
    page -= 1
    for x in values:
        values[values.index(x)] = x - 4
    Template()

def NextPage():
    #Miscellaneous.DestroyFrame()
    page += 1
    for x in values:
        values[values.index(x)] = x + 4
    Template()

def Template():
    usernameHeader = tkinter.Label(frame, text = " Username ", font = "calibri 15 bold", bg = "#99ccff", width = 20, borderwidth = 1, relief = "solid")
    usernameHeader.grid(column = 1, row = 2, sticky = "w", padx = (65,0), pady = (25,0))
    scoreHeader = tkinter.Label(frame, text = " High Score ", font = "calibri 15 bold", bg = "#99ccff", width = 10, borderwidth = 1, relief = "solid")
    scoreHeader.grid(column = 1, row = 2, sticky = "w", padx = (269, 0), pady = (25,0))
    timeHeader = tkinter.Label(frame, text = " Time ", font = "calibri 15 bold", bg = "#99ccff", width = 10, borderwidth = 1, relief = "solid")
    timeHeader.grid(column = 1, row = 2, sticky = "w", padx = (373,0), pady = (25,0))

    firstLine = tkinter.Label(frame, text = usernames[values[0]], font = "calibri 15 bold", width = 20, borderwidth = 1, relief = "solid") 
    firstLine.grid(column = 1, row = 3, padx = (65,0), sticky = "w")
    sLine = tkinter.Label(frame, text = scores[values[0]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    sLine.grid(column = 1, row = 3, sticky = "w", padx = (269,0))
    tLine = tkinter.Label(frame, text = displayTimes[values[0]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    tLine.grid(column = 1, row = 3, sticky = "w", padx = (373,0))

    tomLine1 = tkinter.Label(frame, text = usernames[values[1]], font = "calibri 15 bold", width = 20, borderwidth = 1, relief = "solid") 
    tomLine1.grid(column = 1, row = 4, padx = (65,0), sticky = "w")
    tomLine2 = tkinter.Label(frame, text = scores[values[1]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    tomLine2.grid(column = 1, row = 4, sticky = "w", padx = (269,0))
    tomLine3 = tkinter.Label(frame, text = displayTimes[values[1]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    tomLine3.grid(column = 1, row = 4, sticky = "w", padx = (373,0))

    pLine = tkinter.Label(frame, text = usernames[values[2]], font = "calibri 15 bold", width = 20, borderwidth = 1, relief = "solid") 
    pLine.grid(column = 1, row = 5, padx = (65,0), sticky = "w")
    pLine2 = tkinter.Label(frame, text = scores[values[2]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    pLine2.grid(column = 1, row = 5, sticky = "w", padx = (269,0))
    pLine3 = tkinter.Label(frame, text = displayTimes[values[2]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    pLine3.grid(column = 1, row = 5, sticky = "w", padx = (373,0))

    aLine = tkinter.Label(frame, text = usernames[values[3]], font = "calibri 15 bold", width = 20, borderwidth = 1, relief = "solid") 
    aLine.grid(column = 1, row = 6, padx = (65,0), sticky = "w")
    aLine2 = tkinter.Label(frame, text = scores[values[3]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    aLine2.grid(column = 1, row = 6, sticky = "w", padx = (269,0))
    aLine3 = tkinter.Label(frame, text = displayTimes[values[3]], font = "calibri 15 bold", width = 10, borderwidth = 1, relief = "solid")
    aLine3.grid(column = 1, row = 6, sticky = "w", padx = (373,0))

Template()
window.mainloop()