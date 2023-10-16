# readlines.py
f = open("input/keys.txt", 'r')

lines = f.readlines()
index = -1

def get_key():
    global index
    index += 1
    return lines[index]

def save():
    f2 = open("output/keys.txt", 'w')
    lines_out = lines[index+1:]
    f2.writelines(lines_out)
    f2.close()
    f.close()