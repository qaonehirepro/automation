data = {"choice1": "MS QP Automation Question1 Answer1", "choice2": "MS QP Automation Question1 Answer2","choice3": "MS QP Automation Question1 Answer3", "choice4": "MS QP Automation Question1 Answer4"}

# q2 = {"choice1": "MS QP Automation Question2 Answer1", "choice2": "MS QP Automation Question2 Answer2","choice3": "MS QP Automation Question2 Answer3", "choice4": "MS QP Automation Question2 Answer4"}
# q3 = {"choice1": "MS QP Automation Question13 Answer1", "choice2": "MS QP Automation Question3 Answer2","choice3": "MS QP Automation Question3 Answer3", "choice4": "MS QP Automation Question3 Answer4"}
# q4 = {"choice1": "MS QP Automation Question4 Answer1", "choice2": "MS QP Automation Question4 Answer2","choice3": "MS QP Automation Question4 Answer3", "choice4": "MS QP Automation Question4 Answer4"}
# q5 = {"choice1": "MS QP Automation Question5 Answer1", "choice2": "MS QP Automation Question5 Answer2","choice3": "MS QP Automation Question5 Answer3", "choice4": "MS QP Automation Question5 Answer4"}
# print(data)

a = "MS QP Automation Question1 Answer4"
position = 0
for key, value in data.items():
    position = position + 1
    if a == value:
        print(key)
        print(position)
