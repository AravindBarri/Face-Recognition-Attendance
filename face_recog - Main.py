#required packages
import face_recognition
import cv2
from openpyxl import Workbook
import datetime

#it starts the video capture by opening the seperate tab window using web cam
video_capture = cv2.VideoCapture(0)
 
#create a new workbook if already existing uses that workbook (EXCEL file)
book=Workbook()
sheet=book.active # make the book as active for this code
    
image_1 = face_recognition.load_image_file("1.jpg") #loads the image from folder into variable
image_1_face_encoding = face_recognition.face_encodings(image_1)[0] # encodes the image for facial recognition in image
    
image_5 = face_recognition.load_image_file("5.jpg")
image_5_face_encoding = face_recognition.face_encodings(image_5)[0]
    
image_7 = face_recognition.load_image_file("7.jpg")
image_7_face_encoding = face_recognition.face_encodings(image_7)[0]
    
image_3 = face_recognition.load_image_file("3.jpg")
image_3_face_encoding = face_recognition.face_encodings(image_3)[0]
    
image_4 = face_recognition.load_image_file("4.jpg")
image_4_face_encoding = face_recognition.face_encodings(image_4)[0]
    
#Create a list of know faces and images
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
#gets the current date, month and time
now= datetime.datetime.now()
today=now.day
month=now.month
    
   
while True:
    ret, frame = video_capture.read() # starts reading the image from openCV (Computer Vision)
    
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25) # size of frame to show around the recognized face
    
    rgb_small_frame = small_frame[:, :, ::-1]
    
    #Start the process of recognizing the face from camera view angle
    if process_this_frame:
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
    
    face_names = []
    for face_encoding in face_encodings:
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding) # compare the openCV face with the knownfaces
        name = "Unknown"
    
        if True in matches:
            first_match_index = matches.index(True) # if match found
            name = known_face_names[first_match_index]
            if int(name) in range(1,61):
                sheet.cell(row=int(name), column=int(today)).value = "Present" #place present in the name as row and date as column
            else:
                pass
    
    face_names.append(name)
    
    process_this_frame = not process_this_frame
    
    #size of the frame to recognize
    for (top, right, bottom, left), name in zip(face_locations, face_names):
           top *= 4
           right *= 4
           bottom *= 4
           left *= 4
    
    cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2) # rectangle frame for recognition
    
    cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
    font = cv2.FONT_HERSHEY_DUPLEX
    cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1) # place text after recognition of the knowfaces with the file name
    
    cv2.imshow('Video', frame) # Combine video with frame
        
    book.save(str(month)+'.xlsx') #once face recognized it will save the excel file for not to lost of data
    
    #the openCV video capture can be closed with following key "q" this will break the continuous loop
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
    
video_capture.release() # stops video capture
cv2.destroyAllWindows() # closes all windows
    
   
