import pyautogui
from pyautogui import ImageNotFoundException
from time import sleep
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

#function used to open visual studio and set up WSL server
def visual():
    pyautogui.press('winleft')
    sleep(0.3)
    pyautogui.write("Visual Studio Code")
    sleep(0.2)
    pyautogui.press('enter')
    sleep(1)
    pyautogui.click(463,915)
    sleep(0.5)
    pyautogui.click(pyautogui.locateOnScreen('images/wsl.png',confidence=0.6))
#function used to open spotify and the playlist I selected
def music():
    #allows me to set volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    #if my volume is muted or at 0 the code unmutes the volume 
    # print((volume.GetMute()) == 1 or (round(volume.GetMasterVolumeLevelScalar()*100) == 0))
    if((volume.GetMute()) == 1):
        pyautogui.press('volumemute')
    volume.SetMasterVolumeLevel(-10.2, None)
    pyautogui.press('winleft')
    sleep(0.3)
    # error handling
    try:
        pyautogui.click(pyautogui.locateOnScreen('images/wind.png'))
    except ImageNotFoundException:
        print("Error Processing Image")
    sleep(0.2)
    pyautogui.write('Spotify')
    pyautogui.press('enter')
    sleep(5)
    try:
        pyautogui.click(pyautogui.locateOnScreen('images/play.png',confidence=0.7))
        pyautogui.click(pyautogui.locateCenterOnScreen("images/shuf.png",confidence=0.5))
        pyautogui.click(pyautogui.locateCenterOnScreen("images/playbutton.png",confidence=0.8))
        pyautogui.click(pyautogui.locateCenterOnScreen("images/min.png"))
    except ImageNotFoundException:
        print("Error Processing Image")
    sleep(2)
#main function
def main():
    music()
    # sleep used to delay the next code to run
    sleep(3)
    visual()

main()