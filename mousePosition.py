# Displays the mouse cursor's current position.
# Trial before identifying text coordinate.

import pyautogui


def current_move():
    print('Press Ctrl-C to quit.')
    try:
        while True:
            x, y = pyautogui.position()
            positionStr = 'x: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
            print(positionStr, end = '')
            print('\b' * len(positionStr), end='', flush=True)
            print('\n')

    except KeyboardInterrupt:
        print('\nDone.')


def current():
    print(pyautogui.position())

    location = pyautogui.locateAllOnScreen('capture.png')
    print(location)
