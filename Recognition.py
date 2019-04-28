import time
import xlwt
import xlrd
from xlutils.copy import copy
import cv2
def draw_boundary(img, classifier, scaleFactor, minNeighbors, color, text, clf):
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    features = classifier.detectMultiScale(gray_img, scaleFactor, minNeighbors)
    coords = []
    for (x, y, w, h) in features:
        cv2.rectangle(img, (x,y), (x+w, y+h), color, 2)
        id, _ = clf.predict(gray_img[y:y+h, x:x+w])
        if id==1:
            cv2.putText(img, "Manav", (x, y-4), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
            arr[0] = 1
        elif id==2:
            cv2.putText(img, "Mani", (x, y - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
            arr[1] = 1
        elif id==3:
            cv2.putText(img, "Akshit", (x, y - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
            arr[2] = 1
        elif id==4:
            cv2.putText(img, "Shivam", (x, y - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
            arr[3] = 1
        else:
            cv2.putText(img, "Unknown", (x, y - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 1, cv2.LINE_AA)
        coords = [x, y, w, h]

    return coords

def recognize(img, clf, faceCascade):
    color = {"blue": (255, 0, 0), "red": (0, 0, 255), "green": (0, 255, 0), "white": (255, 255, 255)}
    coords = draw_boundary(img, faceCascade, 1.1, 10, color["white"], "Face", clf)
    return img

faceCascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
arr = [0,0,0,0]
clf = cv2.face.LBPHFaceRecognizer_create()
clf.read("classifier.yml")

video_capture =l cv2.VideoCapture(-1)

while True:
    _, img = video_capture.read()
    img = recognize(img, clf, faceCascade)
    cv2.imshow("face detection", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

video_capture.release()
dateCount = 1
rb = xlrd.open_workbook("Attendance.xlsx")
wb = copy(rb)
w_sheet = wb.get_sheet(0)

j = 1;

for i in range(len(arr)):
	if arr[i] == 1 :
		w_sheet.write(i+2,j,'P')
	else:
		w_sheet.write(i+2,j,'A')
str = time.strftime('%d %b %y')
w_sheet.write(1,dateCount,str)
j = j+1;
dateCount = dateCount + 1

wb.save('Attendance.xlsx')
for i in range(len(arr)):
    arr[i] = 0

cv2.destroyAllWindows()