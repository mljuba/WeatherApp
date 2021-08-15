def myFun(*argv):
    i = 0
    text = ""
    for arg in argv:
        i = i + 1
    i = i-1
    for arg in argv:
        if i > 0:
            text = text + arg + "-"
            i -= 1
        else:
            text = text + arg
    print(text)


myFun('Hello', 'Welcome', 'to', 'GeeksforGeeks')