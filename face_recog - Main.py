import face_recognition
import cv2
from openpyxl import Workbook
import datetime


video_capture = cv2.VideoCapture(0)
    
book=Workbook()
sheet=book.active
    
image_1 = face_recognition.load_image_file("1.jpg")
image_1_face_encoding = face_recognition.face_encodings(image_1)[0]
    
image_5 = face_recognition.load_image_file("5.jpg")
image_5_face_encoding = face_recognition.face_encodings(image_5)[0]
    
image_7 = face_recognition.load_image_file("7.jpg")
image_7_face_encoding = face_recognition.face_encodings(image_7)[0]
    
image_3 = face_recognition.load_image_file("3.jpg")
image_3_face_encoding = face_recognition.face_encodings(image_3)[0]
    
image_4 = face_recognition.load_image_file("4.jpg")
image_4_face_encoding = face_recognition.face_encodings(image_4)[0]
    
    
known_face_encodings = [
        
        image_1_face_encoding,
        image_5_face_encoding,
        image_7_face_encoding,
        image_3_face_encoding,
        image_4_face_encoding
        
    ]
known_face_names = [
        
        "1",
        "5",
        "7",
        "3",
        "4"
       
    ]
    
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True
    
now= datetime.datetime.now()
today=now.day
month=now.month
    
   
while True:
    ret, frame = video_capture.read()
    
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    
    rgb_small_frame = small_frame[:, :, ::-1]
    
    if process_this_frame:
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
    
    face_names = []
    for face_encoding in face_encodings:
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"
    
        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]
            if int(name) in range(1,61):
                sheet.cell(row=int(name), column=int(today)).value = "Present"
            else:
                pass
    
    face_names.append(name)
    
    process_this_frame = not process_this_frame
    
    
    for (top, right, bottom, left), name in zip(face_locations, face_names):
           top *= 4
           right *= 4
           bottom *= 4
           left *= 4
    
    cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
    
    cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
    font = cv2.FONT_HERSHEY_DUPLEX
    cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)
    
    cv2.imshow('Video', frame)
        
    book.save(str(month)+'.xlsx')
    
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
    
video_capture.release()
cv2.destroyAllWindows()
    
   
