import sys
import numpy as np
import cv2
import openpyxl

sheet = None
row = 1
frame = -1
channelProperties = {} # { channelNumber: { propertyName : value } }
end = False

capture = None
fps = 30

def loadVideo(path):
    global row, sheet
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])

def nextFrame():
    global row, frame, end

    frame += 1

    nextEventTime = int(sheet['A' + str(row + 1)].value)
    if nextEventTime == frame:
        row += 1
        channel = sheet['B' + str(row)].value
        if channel == "END":
            frame -= 1
            row -= 1
            end = True
            return
        channel = int(channel)
        if channel not in channelProperties:
            channelProperties[channel] = { }
        letter = "C"
        while True:
            propertyName = sheet[letter + str(row)].value
            if propertyName == None or propertyName.strip() == "":
                break
            letter = chr(ord(letter) + 1)
            propertyValue = sheet[letter + str(row)].value
            letter = chr(ord(letter) + 1)
            channelProperties[channel][propertyName] = propertyValue
            handlePropertyChange(channel, propertyName, propertyValue)

def handlePropertyChange(channel, name, value):
    global capture, fps
    if channel == 0 and name == "file":
        if capture is not None:
            capture.release()
        capture = cv2.VideoCapture(value)
        fps = float(capture.get(cv2.CAP_PROP_FPS))
        print("fps", fps)
    elif channel == 0 and name == "frame":
        if capture is not None:
            capture.set(cv2.CAP_PROP_POS_MSEC, int(value) * (1000.0 / fps))


if __name__ == "__main__":
    loadVideo(sys.argv[1])
    while not end:
        nextFrame()
        if capture is not None:
            if not capture.isOpened():
                capture = None
            else:
                ret, img = capture.read()
                cv2.imshow('xlvidedit500', img)
                if cv2.waitKey(int(1000.0 / fps)) & 0xFF == ord('q'):
                    break
    if capture is not None:
        capture.release()
        cv2.destroyAllWindows()