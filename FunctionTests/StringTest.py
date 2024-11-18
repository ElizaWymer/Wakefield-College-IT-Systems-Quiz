newAnswer = "Answer Bananswer"
splitAnswer = newAnswer.split()
newHint = ""

for x in splitAnswer:
    subAnswer = x[1:len(x)]
    subHint = x[0]
    for x in subAnswer:
        subHint += "-"
    newHint += subHint + " "

print(newHint)
