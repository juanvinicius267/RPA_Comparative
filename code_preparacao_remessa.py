#----Importing the Libs ----#

import pyautogui
import time
import webbrowser
import openpyxl #falta instalar as libs
#import requests #falta instalar as libs
#----Declaration of Global variables ----#

#---- url ways ----#
webBrowser = "internet explorer"
incresePage = './images/kdex/icon_increase_page.png'
urlKdex = "http://slainhouse.br.scania.com:7778/forms/frmservlet?config=flashweb"
urlOas = ''
urlAssist = 'https://assist.scania.com.br/site/Base/Index'
listOfUrls = [urlKdex]#, urlOas, urlAssist]
folderUrl = 'G:/20_MAINTENANCE/04 - REPOSITORIO/05 - RPA/Save File'

#---- Image Path ----#
textKdexImage = "./images/kdex/textKdexImage.png"
textMainMenuImage = "./images/kdex/mainMenu.png"
relatoriosImage = "./images/kdex/relatoriosIcon.png"
engenhariaImage = "./images/kdex/engenharia.png"
comparativosDePopId = "./images/kdex/comparativosDePopId.png"
gerarComparativo = "./images/kdex/gerarComparativo.png"
excel = "./images/kdex/excel.png"
saveAsExcel = "./images/kdex/saveAsExcel.png"
pathFolderIcon = "./images/kdex/pathFolderIcon.png"
oracleIcon = "./images/kdex/oracleIcon.png"
maximizePage = './images/kdex/maximizePage.png'
btnSave = './images/kdex/btnSave.png'
oracleDesktopIcon = './images/kdex/oracleDesktopIcon.png'
eraseInput = './images/kdex/eraseInput.png'
telaInput = './images/kdex/telaInput.png'
savePdf = './images/kdex/savePdf.png'
scaniaIconPdf = './images/kdex/scaniaIconPdf.png'

#----Class that find the image and return the coordinate on screen----#

def findIcon(filePath):
    if( (type(filePath) == str) and len(filePath) >= 5):
        print(filePath)
        start = pyautogui.locateCenterOnScreen(filePath)
        print(start)
        return start
        
    else:
         print("Caminho do arquivo n�o encontrado!")

#----Function to double click on objects
         
def goToObject(position,repeat):
    i = 0
    pyautogui.moveTo(position)
    time.sleep(2)    
    while i < repeat:
        pyautogui.click()
        i+= 1
#---- Function to type write on search input ----#
    
def typeWindowsSearch(text):
    pyautogui.press('win')
    pyautogui.write(text, interval=0.25)
    pyautogui.hotkey('enter')
    
#---- Increse the page ----#
    
def clickToIncreasePage():
    start = pyautogui.locateCenterOnScreen(incresePage)
    print(start)
    if(start != "None"):
        pyautogui.moveTo(start)
        time.sleep(2)
        pyautogui.click()
        
#---- Open link on web browser ----#
def openLink(link):
    webbrowser.open(link)



    
#----Main Program ----#
    
#Open applications links
for x in listOfUrls:
    openLink(x)
    time.sleep(2)

#try to find kdex
time.sleep(20)


goToObject(findIcon(oracleIcon),2)


goToObject(findIcon(textKdexImage),2)
time.sleep(4)

goToObject(findIcon(textMainMenuImage),2)
time.sleep(5)

goToObject(findIcon(relatoriosImage),2)

goToObject(findIcon(engenhariaImage),2)

goToObject(findIcon(comparativosDePopId),2)

# Requisar qual pop id seria necessaria realizar o comparativo
pyautogui.write("574574")
time.sleep(2)
pyautogui.hotkey('tab')
time.sleep(1)
pyautogui.hotkey('tab')
time.sleep(1)
pyautogui.write("574573",interval=0)
time.sleep(1)
pyautogui.hotkey('enter')

positionExcel = findIcon(excel)

goToObject(positionExcel,2)

goToObject(findIcon(gerarComparativo),1)

time.sleep(5)
goToObject(findIcon(saveAsExcel),1)

#goToObject(findIcon(pathFolderIcon),2)

for x in range(7):
    pyautogui.hotkey('tab')
    time.sleep(1)
    
pyautogui.hotkey('enter')

time.sleep(10)

pyautogui.write(folderUrl, interval=0)

pyautogui.hotkey('enter')

time.sleep(2)
goToObject(findIcon(btnSave),1)

time.sleep(2)
goToObject(findIcon(oracleDesktopIcon),1)

time.sleep(2)
goToObject(findIcon(eraseInput),1)

time.sleep(2)
pyautogui.write("574572",interval=0)

time.sleep(2)
goToObject(findIcon(telaInput),1)


time.sleep(1)
goToObject(findIcon(gerarComparativo),1)


time.sleep(7)
goToObject(findIcon(scaniaIconPdf),0)

time.sleep(4)
pyautogui.press(['ctrl','shift','s'])


time.sleep(2)
goToObject(findIcon(btnSave),1)

