import subprocess
import os
import win32gui
import win32con
import re
import psutil
import time
import pyautogui

pyautogui.FAILSAFE = True  # When fail-safe mode is True, moving the mouse to the upper-left will raise a
# pyautogui.FailSafeException that can abort the program


# Constant Variables
background_Windows = ['Default IME',
                      'MSCTFIME UI',
                      'TouchPad object helper window',
                      'Touchpad driver tray icon window']

root_folder = r'C:\mxdData\todo'
mxdCounter = 0


# Start of Method Definitions
def printHeader():
    header_metadata = ['    $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$',
                       '    $$             ____________________________                  $$',
                       '    $$            |    Speedy Starfish v1.0    |                 $$',
                       '    $$            |____________________________|                 $$',
                       '    $$             Last updated: October 10, 2019                $$',
                       '    $$                                                           $$',
                       '    $$             an Auto-clicker/Read/Write Bot                $$',
                       '    $$                           ,                               $$',
                       '    $$                        __/ \__                            $$',
                       r"    $$                        \     /                            $$",
                       r'''    $$                        /_   _\                            $$''',
                       '    $$                          \ /                              $$',
                       r"    $$                           '                               $$",
                       '    $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$']
    print('\n'.join(header_metadata))


def printLoading():
    counter = 0
    while counter < 60:
        bar = ""
        for i in range(counter):
            bar += '*'
        time.sleep(0.5)
        print('[' + bar + ']', end='\r')
        time.sleep(0.4)
        counter += 1


def printStartAutoclick():
    print("----------------------------------------------------------")
    print("\t[...Initiating [pyautogui] Script...]")
    print("If the script must be stopped while running, move Mouse to the Top-Left Corner to Force Quit...")
    print("----------------------------------------------------------")
    pyautogui.PAUSE = 1
    return


def printFileDirectory(mxdPathList):
    print("\n###############################################################################################\n")
    print("     --------------------------------------------------------------------------------------------")
    print("    |     LOCATED [{}] .MXD FILE(S) INSIDE OF {}".format(len(mxdPathList), root_folder) + "      |")
    print("     --------------------------------------------------------------------------------------------")
    for path in mxdPathList:
        print(path)
        time.sleep(0.01)
    print("----------------------------------------------------------")


def readFileDirectory(basepath):
    mxdPathList = []
    global mxdCounter
    with os.scandir(basepath) as entries:
        for entry in entries:
            if entry.is_dir():
                print("Looking into : {} ...".format(entry.path))

                mxdPathList = mxdPathList + (readFileDirectory(entry))
            else:
                if entry.is_file() and entry.name.endswith('.mxd'):
                    mxdPathList.append(entry.path)
                    mxdCounter += 1
                    print('\t\t...Found {} .mxds so far...'.format(mxdCounter), end='\r')
                    time.sleep(0.1)
    return mxdPathList


def openFileWArcMap(file_location):
    # Open Specific Folder
    file = file_location
    ArcMapExe = r'C:\Program Files (x86)\ArcGIS\Desktop10.5\bin\ArcMap.exe'
    subprocess.Popen("%s %s" % (ArcMapExe, file))
    return


def openFileWithApp(file_location):
    # Open Specific Folder
    file = file_location
    acrobatPath = r'C:\WINDOWS\system32\notepad.exe'
    subprocess.Popen("%s %s" % (acrobatPath, file))
    return


def initialCheckForArcMap():
    cb = lambda x, y: y.append(x)
    wins = []
    win32gui.EnumWindows(cb, wins)

    # now check to see if any match our regexp:
    tgtWin = -1
    passed_Windows = []
    for win in wins:
        txt = win32gui.GetWindowText(win)

        if len(txt) > 3 and txt not in background_Windows and txt not in passed_Windows:
            passed_Windows.append(txt)

        if re.search(r'ArcMap', txt):
            tgtWin = win
    if tgtWin >= 0:
        print('          ... ArcMap.exe is Already Running! ...')

        win32gui.ShowWindow(tgtWin, win32con.SW_MAXIMIZE)
        print('----------------------------------------------------------')
        return True
    return False


def waitForArcMap():
    escape = False
    while not escape:
        cb = lambda x, y: y.append(x)
        wins = []
        win32gui.EnumWindows(cb, wins)
        print("waiting...")
        # now check to see if any match our regexp:
        tgtWin = -1
        passed_Windows = []
        for win in wins:
            txt = win32gui.GetWindowText(win)

            if len(txt) > 3 and txt not in background_Windows and txt not in passed_Windows:
                passed_Windows.append(txt)

            if re.search(r'ArcMap', txt):
                tgtWin = win
                escape = True
                print('Found!!')
        if tgtWin >= 0:
            print('          ... ArcMap.exe is Already Running! ...')
            win32gui.ShowWindow(tgtWin, win32con.SW_MAXIMIZE)
            print('----------------------------------------------------------')
            return True
        pyautogui.PAUSE(3.5)
    return escape


def checkForArcMap():
    # start by getting a list of all the windows:
    arcMap_bool = False
    nameRe = r'ArcMap'
    print("----------------------------------------------------------")
    print("... Searching Windows to see if ArcMap 10.5.1 is open ...")
    print("----------------------------------------------------------")
    while not arcMap_bool:

        cb = lambda x, y: y.append(x)
        wins = []
        win32gui.EnumWindows(cb, wins)

        # now check to see if any match our regexp:
        tgtWin = -1
        passed_Windows = []

        for win in wins:
            txt = win32gui.GetWindowText(win)

            if len(txt) > 3 and txt not in background_Windows and txt not in passed_Windows:
                passed_Windows.append(txt)

            if re.search(nameRe, txt):
                tgtWin = win
                print(' -------------------------')
                print("| Currently Open Windows  |")
                print(' -------------------------')
                print("\n".join(passed_Windows))
                time.sleep(0.25)
                print('----------------------------------------------------------')
                arcMap_bool = True
                break
        if tgtWin >= 0:
            win32gui.ShowWindow(tgtWin, win32con.SW_MAXIMIZE)
    return arcMap_bool


def forceFrontWindow(application):
    cb = lambda x, y: y.append(x)
    wins = []
    win32gui.EnumWindows(cb, wins)

    # now check to see if any match our regexp:
    tgtWin = -1
    passed_Windows = []

    for win in wins:
        txt = win32gui.GetWindowText(win)

        if len(txt) > 3 and txt not in background_Windows and txt not in passed_Windows:
            passed_Windows.append(txt)

        if re.search(application, txt):
            tgtWin = win
            break
    if tgtWin >= 0:
        print('----------------------------------------------------------')
        print('\t\tExpanding to Full Window ')
        window = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window, win32con.SW_MAXIMIZE)
        print('----------------------------------------------------------')


def checkIfProcessRunning(processName):
    print(' ----------------------------------------------------------')
    print('| Checking if {}.exe is Running and Fully Responding...|'.format(processName))
    print(' ----------------------------------------------------------')
    counter = 0
    process_bool = False
    iteration_counter = 0  # counter to prevent infinite while loop error, if ArcMap does not load indefinitely
    process_counter = 0

    while not process_bool and iteration_counter < 20:
        '''
        Check if there is any running process that contains the given name processName.
        '''
        # Iterate over the all the running process
        for proc in psutil.process_iter():
            process_counter += 1
            print('\t...Searching... Looped through: {} Processes'.format(process_counter), end='\r')
            time.sleep(0.03)
            print('\t...         ... Looped through: {} Processes'.format(process_counter), end='\r')
            try:
                # Check if process name contains the given name string.
                if processName.lower() in proc.name().lower():
                    print("{}.exe is currently [RUNNING] and [RESPONDING]. It is ready to start AutoClick()...".format(
                        processName))
                    process_bool = True
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        iteration_counter += 1
        print(
            "\t Looped through all Running Processes. No instance of {}.exe has been found yet... ".format(processName))
    return process_bool


def raiseWindowNamed(nameRe):
    # start by getting a list of all the windows:
    cb = lambda x, y: y.append(x)
    wins = []
    win32gui.EnumWindows(cb, wins)
    # now check to see if any match our regexp:
    tgtWin = -1
    passed_Windows = []

    for win in wins:
        txt = win32gui.GetWindowText(win)
        print("..Searching through all running Windows...", end='\r')
        if len(txt) > 3 and txt not in background_Windows and txt not in passed_Windows:
            passed_Windows.append(txt)

        if re.search(nameRe, txt):
            tgtWin = win
            print('----------------------------------------------------------')
            print("\n".join(passed_Windows))
            print('----------------------------------------------------------')
            print("    LOCATED: " + win32gui.GetWindowText(tgtWin))
            break
    if tgtWin >= 0:
        print('Expanding  Window')
        win32gui.SetForegroundWindow(tgtWin)
        win32gui.ShowWindow(tgtWin, win32con.SW_MAXIMIZE)
        print('----------------------------------------------------------')


def openDataFrameProperties():
    print("Starting Autoclick Script - Please DO NOT touch Mouse...")
    # Fail-safe Force to Middle of Screen
    pyautogui.moveTo(500, 500, 0.5)
    # Right Click Layers (DataFrame)
    pyautogui.moveTo(61, 155, 1)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='right')
    pyautogui.PAUSE = 0.5
    # Left Click (select) Properties
    pyautogui.moveTo(120, 588, 1)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')


def ApplyPreserveLayerChange():
    # Left Click General Tab (fail-safe)
    pyautogui.moveTo(766, 255, 1.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click CheckBox (Preserve layer ID)
    pyautogui.moveTo(744, 686, 1.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click OK
    pyautogui.moveTo(992, 830, 1.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    print("AutoClick Script Finished")


def saveAndExit():
    # Fail-safe Force to Middle of Screen
    pyautogui.moveTo(500, 500, 1.5)
    # Left Click Exit
    pyautogui.moveTo(1883, 12, 1.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click Save
    pyautogui.moveTo(883, 568, 1.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    pyautogui.PAUSE = 4


def openDataFramePropertiesLaptop():
    print("Starting Autoclick Script - Please DO NOT touch Mouse...")
    # Fail-safe Force to Middle of Screen
    pyautogui.moveTo(681, 346, 0.5)
    # Right Click Layers (DataFrame)
    pyautogui.moveTo(56, 156, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='right')
    pyautogui.PAUSE = 0.5
    # Left Click (select) Properties
    pyautogui.moveTo(142, 591, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')


def ApplyPreserveLayerChangeLaptop():
    # Left Click General Tab (fail-safe)
    pyautogui.moveTo(480, 95, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click CheckBox (Preserve layer ID)
    pyautogui.moveTo(467, 533, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click Apply
    pyautogui.moveTo(886, 675, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click OK
    pyautogui.moveTo(718, 670, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    print("AutoClick Script Finished")


def saveAndExitLaptop():
    # Left Click Exit
    pyautogui.moveTo(1340, 10, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    # Left Click Save
    pyautogui.moveTo(623, 415, 0.5)
    pyautogui.PAUSE = 0.5
    pyautogui.click(button='left')
    pyautogui.PAUSE = 2


def AutoClick():
    printStartAutoclick()
    openDataFrameProperties()
    ApplyPreserveLayerChange()
    saveAndExit()


def arcLoop(root_path):
    print("Starting the App")
    mxdPathList = readFileDirectory(root_path)
    printFileDirectory(mxdPathList)
    pyautogui.alert('Press OK to Initialize Script')
    mxdLeft = mxdPathList
    countPassed = 0

    for singleFile in mxdPathList:
        print()
        print("=====================================================================================================")
        print("\t...Starting AutoScript() on:")
        print("\t-------------------------")
        print("\t{}".format(singleFile))
        print("\t" + '_' * len("\t | Current Number of .mxds left: {} |".format(len(mxdPathList))))
        print("\t| Current Number of .mxds left: {} |".format(len(mxdLeft) - countPassed))
        print("\t" + '_' * len("\t | Current Number of .mxds left: {} |".format(len(mxdPathList))))

        if not initialCheckForArcMap():
            print('\t\t...Opening ArcMap...')
            pyautogui.PAUSE = 1
        openFileWArcMap(singleFile)
        forceFrontWindow(r'ArcMap')
        printLoading()
        checkForArcMap()
        waitForArcMap()
        pyautogui.PAUSE = 1

        if checkIfProcessRunning(r'ArcMap'):
            forceFrontWindow(r'ArcMap')
            AutoClick()
            countPassed += 1
            print("----------------------------------------------------------")

        print("\t...Finishing on :{}".format(singleFile))
        print("----------------------------------------------------------")
        print("\tRemoved: {} from List of MXD Filepaths".format(singleFile))
        pyautogui.PAUSE = 1
        print("=====================================================================================================")
        printFileDirectory(mxdPathList[countPassed:])
        print()


def main():
    printHeader()
    arcLoop(root_folder)


if __name__ == "__main__":
    main()
