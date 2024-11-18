rawTime = 155
minutes = []

splitTime = str(rawTime / 60).rsplit(".")
for x in splitTime:
    minutes.append(x)
seconds = round(int(minutes[1]) / (10**len(minutes[1]))  * 60)

if int(minutes[0]) < 10:
    quizTime = "0" + str(minutes[0])
else:
    quizTime = str(minutes[0])
if seconds < 10:
    quizTime += ":0" + str(seconds)
else:   
    quizTime += ":" + str(seconds)

print(quizTime)
