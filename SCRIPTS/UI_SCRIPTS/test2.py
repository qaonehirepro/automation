answer_options = ('option0', 'option1', 'option2', 'option3')
# a = "option0,option1,option2, option3"
a = "option1"
if a:
    b = a.split(',')
    for i in b:
        i = i.strip()
        print(i)
else:
    print("This is else")
