import xlwings as xw
from time import sleep
from pynput import keyboard

wb = xw.Book('Book2.xlsx')

for sheet in wb.sheets:
    print(sheet)

sht1 = wb.sheets['Sheet1']
sht1.range('A1').color = (255,0,0)

num = 1

downKey = False
upKey = False
rightKey = False
leftKey = False

def on_press(key):
    global upKey, downKey, rightKey, leftKey
    if key == keyboard.Key.up:
        print("Up Key")
        upKey = True
    if key == keyboard.Key.down:
        print("Down Key")
        downKey = True
    if key == keyboard.Key.right:
        print("Right Key")
        rightKey = True
    if key == keyboard.Key.left:
        print("Left Key")
        leftKey = True

# ...or, in a non-blocking fashion:
listener = keyboard.Listener(
    on_press=on_press)
listener.start()

ypos = 0
xpos = 0

sht1.range(chr(65 + xpos) + str(ypos+1)).color = (255,0,0)


while True:
    sleep(.1)

    oldypos = ypos
    oldxpos = xpos

    if upKey:
        ypos = ypos - 1
    elif downKey:
        ypos = ypos + 1
    elif rightKey:
        xpos = xpos + 1
    elif leftKey:
        xpos = xpos - 1
    
    if ypos < 0:
        ypos = 0
    if xpos < 0:
        xpos = 0


    if xpos != oldxpos or ypos != oldypos:
        sht1.range(chr(65 + oldxpos) + str(oldypos+1)).color = (255,255,255)
        sht1.range(chr(65 + xpos) + str(ypos+1)).color = (255,0,0)

    upKey = False
    downKey = False
    leftKey = False
    rightKey = False


