import pyautogui
import pyperclip

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    screenWidth, screenHeight = pyautogui.size()
    print(screenWidth, screenHeight)

    currentMouseX, currentMouseY = pyautogui.position()
    print(currentMouseX, currentMouseY)
    pyperclip.copy('The text to be copied to the clipboard.')
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
