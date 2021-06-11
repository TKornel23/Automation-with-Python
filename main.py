import pandas as pd
import pyautogui as py
py.PAUSE = 1

def getKotroczDataFrame():
    p = open(r"C:\Users\ttora\Desktop\kotrocz.xlsx", "rb")
    kotroczExcel = pd.read_excel(p, engine="openpyxl", header=0)
    kotrocz = pd.DataFrame(kotroczExcel)
    return kotrocz

def writeForPlaces(i):
    kotroczDataFrame = getKotroczDataFrame()

    py.click(827,196)
    py.click(644,588, clicks=2)
    py.click(903,228)
    py.write(str(kotroczDataFrame.iloc[i,6]))
    py.click(697,264)
    py.write(str(kotroczDataFrame.iloc[i,2]) + " MLSZ.: " + str(kotroczDataFrame.iloc[i,5]))
    py.click(1181,797)
    py.click(621,269)
    if(i > 0):
        py.click(886,581)

if __name__ == '__main__':
    dataframeSize = getKotroczDataFrame().size

    for i in range(0,int(dataframeSize)-1):
        writeForPlaces(i)
    py.alert("Vége")

