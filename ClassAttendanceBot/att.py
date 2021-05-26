import pyautogui
import webbrowser
import time
import os

def attendclass(url,time_hour,duration = 1,print_waitTime=False):
    n = 20
    pyautogui.FAILSAFE =True
    wait_time = 10
    if time_hour not in range(0,25):
        print("Invalid time format")
    if time_hour == 0:
        time_hour = 24
    callsec = (time_hour*3600 )
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec

    if currhr == 0:
        currhr = 24

    currtotsec = (currhr*3600)+(currmin*60)+(currsec)
    lefttm = callsec-currtotsec

    if lefttm <= -3600:
        lefttm = 86400+lefttm
    sleeptm = max(lefttm-wait_time,0)
    if print_waitTime :
        print(f"In {sleeptm} seconds meet.google.com will open and after {wait_time} seconds meeting will be joined")
    time.sleep(sleeptm)
    webbrowser.open(url)
    time.sleep(5)
    pyautogui.click(1761, 180)
    time.sleep(10)
    x, y = pyautogui.locateCenterOnScreen('kiit.png',confidence = 0.5 , grayscale =True)
    pyautogui.click(x, y)
    time.sleep(10)
    pyautogui.hotkey('ctrl', 'e')
    pyautogui.hotkey('ctrl', 'd')
    xy4 = None
    xy5 = None
    while xy4 or xy5 is None:
        xy4=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'form_roll.png'),grayscale=True)
        xy5=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'breakout.png'),grayscale=True)
    if xy4!=None:
        pyautogui.click(xy4)
    else:
        pyautogui.click(xy5)
    time.sleep(5)
    xy1 = None 
    i=0
    while xy1 is None and i<n:
        i = i+1
        xy1=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'msg.png'),confidence= 0.7 ,grayscale=True)
    pyautogui.click(xy1)
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec
    i=0
    xy2 = None
    while i<(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec))):
        i = i + 200
        # print(i)
        time.sleep(200)
        xy2 = pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'form.png'),grayscale=False)
        if xy2!=None:
            # print(xy2)
            pyautogui.click(xy2)
            time.sleep(5)
            xy4 = None
            while xy4 is None:
                print("Hi")
                pyautogui.hotkey('fn', 'f5')
                time.sleep(100)
                xy4=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'form_roll.png'),grayscale=True)
            time.sleep(1)
            print("Game over")
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.write('1828061')
            time.sleep(1)
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(5)
            pyautogui.hotkey('alt','f4')
            break
        # print(xy2)
        if len(list(pyautogui.locateAllOnScreen('meet_roll.png'))) > 5:
            pyautogui.write('1828061')
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.hotkey('alt', 'f4')
    width,height = pyautogui.size()
    pyautogui.click(width/2,height/2)
    xy = None 
    i=0
    while xy is None and i<n:
        i = i+1
        xy=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'end.png'),grayscale=True)
    pyautogui.click(xy)


def attendclasszoom(url,time_hour,duration = 1,passcode = None,print_waitTime=False):
    n = 20
    pyautogui.FAILSAFE =True
    pyautogui.press('win')
    pyautogui.write('zoom')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.hotkey('alt' , 'f4')
    wait_time = 10
    if time_hour not in range(0,25):
        print("Invalid time format")
    if time_hour == 0:
        time_hour = 24
    callsec = (time_hour*3600)
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec

    if currhr == 0:
        currhr = 24

    currtotsec = (currhr*3600)+(currmin*60)+(currsec)
    lefttm = callsec-currtotsec

    if lefttm <= -3600:
        lefttm = 86400+lefttm
    sleeptm = max(lefttm-wait_time,0)
    if print_waitTime :
        print(f"In {sleeptm} seconds zoom.com will open and after {wait_time} seconds meeting will be joined")
    time.sleep(sleeptm)
    webbrowser.open(url, new=1)
    time.sleep(1)
    xy=None
    i=0
    while xy is None and i<n:
        i = i+1
        xy=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'launch.png'),grayscale=True)
    pyautogui.click(xy)
    time.sleep(5)
    if passcode!=None:
        pyautogui.write(passcode)
        pyautogui.press('enter')
    xy1 =None
    while xy1 is None:
        xy1=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'join.png'),grayscale=True)  
    pyautogui.click(xy1)
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    pyautogui.hotkey('alt', 'f4')
    xy2 =None
    i=0
    while xy2 is None and i<n:
        i = i+1
        xy2=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'expand.png'),grayscale=True)
    pyautogui.click(xy2)
    pyautogui.hotkey('alt', 'h')
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec
    print(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec)))
    xy3 = None
    i = 0
    while i<(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec))):
        i = i + 200
        print(i)
        time.sleep(200)
        xy3 = pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'formzoom.png'),grayscale=False)
        if xy3!=None:
            pyautogui.click(xy3)
            time.sleep(5)
            xy4 = None
            xy5 =None
            while xy4 or xy5 is None:
                print("Hi")
                pyautogui.hotkey('fn', 'f5')
                time.sleep(100)
                xy4=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'form_roll.png'),grayscale=True)
                xy5=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'breakout.png'),grayscale=True)
                if len(list(pyautogui.locateAllOnScreen('roll.png'))) > 5:
                    pyautogui.write('1828061')
                    pyautogui.press('enter')
                    time.sleep(5)
                    pyautogui.hotkey('alt','f4')
                    exit()
                
            time.sleep(1)
            print("Game over")
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.write('1828061')
            time.sleep(1)
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(5)
            pyautogui.hotkey('alt','f4')
            break

        # UNTESTED CODE HERE TO ENTER ROLL WHEN ROLL NUMBERS FOUND
        
    pyautogui.hotkey('alt', 'q')
    time.sleep(1)
    pyautogui.click(x=1015, y=480)

def attendtnp(time_hour,duration = 1,print_waitTime=True):
    n=10
    wait_time = 10
    if time_hour not in range(0,25):
        print("Invalid time format")
    if time_hour == 0:
        time_hour = 24
    callsec = (time_hour*3600)
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec

    if currhr == 0:
        currhr = 24

    currtotsec = (currhr*3600)+(currmin*60)+(currsec)
    lefttm = callsec-currtotsec

    if lefttm <= -3600:
        lefttm = 86400+lefttm
    sleeptm = max(lefttm-wait_time,0)
    if print_waitTime :
        print(f"In {sleeptm} seconds zoom.com will open and after {wait_time} seconds meeting will be joined")
    time.sleep(sleeptm)
    webbrowser.open('https://web.whatsapp.com',new = 1)
    time.sleep(5)
    xy5 = None
    while xy5 is None:
        xy5 = pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'searchbarwhatsapp.png'),grayscale=True,confidence = 0.9)  
    pyautogui.click(xy5)
    pyautogui.write("CAAS")
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    xy6 =None
    while xy6 is None:
        xy6=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'whatsappsearch.png'),grayscale=True)  
    pyautogui.click(xy6)
    time.sleep(1)
    pyautogui.write("zoom")
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    x =None
    y= None
    while x is None:
        x , y =pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'whatsappzoomjoin.png'),grayscale=True)  
    pyautogui.click(x,y)

    time.sleep(1)
    xy=None
    i=0
    while xy is None and i<n:
        i = i+1
        xy=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'launch.png'),grayscale=True)
    pyautogui.click(xy)
    xy1 =None
    while xy1 is None:
        xy1=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'join.png'),grayscale=True)  
    pyautogui.click(xy1)
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    pyautogui.hotkey('alt', 'f4')
    xy2 =None
    i=0
    while xy2 is None and i<n:
        i = i+1
        xy2=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'expand.png'),grayscale=True)
    pyautogui.click(xy2)
    pyautogui.hotkey('alt', 'h')
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec
    print(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec)))
    xy3 = None
    i = 0
    while i<(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec))):
        i = i + 200
        print(i)
        time.sleep(200)
        xy3 = pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'formzoom.png'),grayscale=False)
        if xy3!=None:
            pyautogui.click(xy3)
            break

        ## UNTESTED CODE HERE TO ENTER ROLL WHEN ROLL NUMBERS FOUND
        # if len(list(pyautogui.locateAllOnScreen('roll.png'))) > 5:
        #     pyautogui.write('1828061')
        #     pyautogui.hotkey('alt', 'q')
        #     time.sleep(1)
        #     pyautogui.click(x=1015, y=480)
    # time.sleep(duration*3600 + callsec - ((currhr*3600)+(currmin*60)+(currsec)))
    pyautogui.hotkey('alt', 'q')
    time.sleep(1)
    pyautogui.click(x=1015, y=480)

def fillform(url,time_hour,duration = 1,print_waitTime=False):
    n = 20
    pyautogui.FAILSAFE =True
    wait_time = 10
    if time_hour not in range(0,25):
        print("Invalid time format")
    if time_hour == 0:
        time_hour = 24
    callsec = (time_hour*3600 )
    curr = time.localtime()
    currhr = curr.tm_hour
    currmin = curr.tm_min
    currsec = curr.tm_sec

    if currhr == 0:
        currhr = 24

    currtotsec = (currhr*3600)+(currmin*60)+(currsec)
    lefttm = callsec-currtotsec

    if lefttm <= -3600:
        lefttm = 86400+lefttm
    sleeptm = max(lefttm-wait_time,0)
    if print_waitTime :
        print(f"In {sleeptm} seconds meet.google.com will open and after {wait_time} seconds meeting will be joined")
    time.sleep(sleeptm)
    webbrowser.open(url)
    time.sleep(5)
    xy = None
    while xy is None:
        print("Hi")
        pyautogui.hotkey('fn', 'f5')
        time.sleep(100)
        xy=pyautogui.locateCenterOnScreen(os.path.join(os.getcwd(),'form_roll.png'),grayscale=True)
    time.sleep(1)
    print("Game over")
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write('1828061')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.hotkey('alt','f4')

 
if __name__ == '__main__':
    # attendclass('https://meet.google.com/lookup/dgs2od5lk2?authuser=1&hs=179' , 12 ,duration=0.75,print_waitTime=True)
    # attendclasszoom('https://us02web.zoom.us/j/88931134636' ,16.5 ,duration=1 ,passcode='enkiit', print_waitTime=True) 
    #url must contain https: or else script opens Edge/ Explorer
    # attendtnp(18.75)
    # attendclasszoom('https://us02web.zoom.us/j/9574661784?pwd=UXZHT0FQNEpQYTVZcmlqZlBoSkZ4dz09', 16,duration=0.8, print_waitTime=True)

    # attendclass('https://meet.google.com/jgn-fkrt-msp',10,duration=1,print_waitTime=True)
    # attendclass('https://meet.google.com/ahr-bbye-hbw?pli=1&authuser=1' , 9 ,duration=2,print_waitTime=True)
    attendclasszoom('https://us02web.zoom.us/j/2223322244?pwd=empRNDZQdnh0bWY3M20zYW4ySDd6UT09', 12,duration=0.75, print_waitTime=True)
    # fillform('https://docs.google.com/forms/d/e/1FAIpQLSfdfLxb3DGx3Uyk0JLESxLq9zFIlh4epw_54rV_c2nT57WAaA/formrestricted',12)