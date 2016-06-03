import Tkinter,tkFileDialog,tkMessageBox
from datetime import datetime
import cv2
import sys
import os
import  time as t
import numpy as np
from PIL import Image
import xlrd
import xlwt 
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
presentin="SRNO:"
name= str(datetime.now())
qw=str(name)+".xls"
#print qw
def atten(change):
    print change
    rb = open_workbook("attendence.xls")
    nb = copy(rb)
    s = nb.get_sheet(0)
    s.write(2,1,'P')

    wb = xlwt.Workbook("attendence.xls") 
    book= xlrd.open_workbook("attendence.xls") 
    first_sheet=book.sheet_by_index(0)

    for i in range(1,6):
       cell=first_sheet.cell(i,2)
       if int(cell.value)==change:
            s.write(i,1,'P')
    nb.save('attendence.xls')
def detect(video_path):
    
    cap = cv2.VideoCapture(video_path)
    while not cap.isOpened():
        cap = cv2.VideoCapture(video_path)
        cv2.waitKey(1000)
        print "Wait for the header"
    # Get user supplied values
    dirname='Faces'
    cascPath = "haarcascade_frontalface_default.xml"
    # Create the haar cascade
    faceCascade = cv2.CascadeClassifier(cascPath)
    i=1
    j=1
    #pos_frame = int(cap.get(cv2.cv.CV_CAP_PROP_POS_FRAMES))
    length = int(cap.get(cv2.cv.CV_CAP_PROP_FRAME_COUNT))
    pos_frame=1
    while True and pos_frame<length:
        flag, frame = cap.read()
        pos_frame = int(cap.get(cv2.cv.CV_CAP_PROP_POS_FRAMES))
        
        z=pos_frame%11
       # print z
        
        if z==0:
            if flag:
                print("\a")
                # The frame is ready and already captured
                print("Frame obtained  .....")
                print("processing frame.....")
                print("\a")
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                # Detect faces in the image
                faces = faceCascade.detectMultiScale(
                gray,
                scaleFactor=1.1,
                minNeighbors=4,
                minSize=(5, 5),
                flags = cv2.cv.CV_HAAR_SCALE_IMAGE
                )
                print("\a")
                print "Found {0} faces!".format(len(faces))

                # Draw a rectangle around the faces
                for (x, y, w, h) in faces:
                    print("\a")
                    #cv2.rectangle(image, (x, y), (x+w, y+h), (0, 255, 0), 2)
                    #cv2.imshow("Faces found" ,image)
                    cv2.imwrite(os.path.join(dirname,"subject%d.jpg" % i),frame[y:y+h,x:x+w])
                    i=i+1
                    


            else:
                # The next frame is not ready, so we try to read it again
                cap.set(cv2.cv.CV_CAP_PROP_POS_FRAMES, pos_frame-1)
                print "frame is not ready"
                # It is better to wait for a while for the next frame to be ready
                cv2.waitKey(1000)
                #   break

    print("FACES DECTED")

         
    


    
        
        
def face_algo(image_path):
    
    pred=cv2.imread(image_path)
    predict_image = cv2.cvtColor(pred, cv2.COLOR_BGR2GRAY)
    nbr_predicted, conf = recognizer.predict(predict_image)
    #nbr_actual = int(os.path.split(image_path)[1].split(".")[0].replace("subject", ""))
    present.append(nbr_predicted)
    for i in range(0,len(present)):
      for j in range(i+1,len(present)):
        if present[i]==present[j]:
            present[j]=0
    for i in range(0,len(present)):
              if(present[i]==0):
                 del present[i]

    
    
class simpleapp_tk(Tkinter.Tk): 
    def __init__(self,parent): 
        Tkinter.Tk.__init__(self,parent) 
        self.parent = parent 
        self.initialize() 
    def initialize(self): 
        self.grid() 
        self.entryVariable = Tkinter.StringVar() 
        self.entry = Tkinter.Entry(self,textvariable=self.entryVariable) 
        self.entry.grid(column=0,row=0,sticky='EW') 
        #self.entry.bind("<Return>", self.OnPressEnter) 
        self.entryVariable.set(u"Browse for video file to load") 
        button = Tkinter.Button(self,text=u"Browse",command=self.OnButtonClickBrowser) 
        button.grid(column=1,row=0) 
        button1 = Tkinter.Button(self,text=u"Process",command=self.OnButtonClickProcess) 
        button1.grid(column=1,row=1) 
        self.labelVariable = Tkinter.StringVar() 
        label = Tkinter.Label(self,textvariable=self.labelVariable,anchor="w",fg="white",bg="blue",width="50") 
        label.grid(column=0,row=2,columnspan=2,sticky='EW') 
        self.grid_columnconfigure(0,weight=1) 
        self.grid_rowconfigure(0,weight=2) 
        self.labelVariable.set("PROCESSING..")
        self.resizable(True,False) 
        self.update() 
        self.geometry(self.geometry()) 
        #self.geometry("250x100,250,100")       
        self.entry.focus_set() 
        self.entry.selection_range(0, Tkinter.END) 
    def OnButtonClickBrowser(self): 
        self.filename = tkFileDialog.askopenfilename(parent=root,defaultextension='.mp4',message='Choose a video file to process ')
        self.entryVariable.set( u""+self.filename ) 
        #self.new()
        if self.filename=="":
             tkMessageBox.showerror(title="ERROR!",message="Choose a valid video and process")
        self.entry.focus_set() 
        self.entry.selection_range(0, Tkinter.END) 
    def OnButtonClickProcess(self): 
          if self.filename=="":
             tkMessageBox.showerror(title="ERROR!",message="Choose a valid video and process")
          else:
              detect(self.filename)
              self.labelVariable.set("FACES EXRACTED..")
              
              to_reco= [os.path.join(path, f) for f in os.listdir(path)]
              to_reco[0]=to_reco[1]
              for image_path in to_reco:
                  face_algo(image_path)
              #print("PRESENT STUDENT SRNO:")
              #for i in range(0,len(present)):
                #need=",".join(str(present[i]))
               #  presentin.join(str(present[i]))
              self.labelVariable.set("DONE !")
              new=" "+str(present)
              for change in present:
                 atten(change)
            
              #print("\a")
              tkMessageBox.showinfo(title="DONE!",message="ATTENDENCE MARKED!")
              #atten(change)
          self.entry.focus_set()           
          self.entry.selection_range(0, Tkinter.END) 
    
print("LOADING DATABASE.....")
labels= np.load('labels.npy')
images=np.load('images.npy')
recognizer = cv2.createLBPHFaceRecognizer()
# Perform the tranining
recognizer.train(images, np.array(labels))
print("LOADED DATA.....")
path = './faces'
present=[]   
if __name__ == "__main__": 
    app = simpleapp_tk(None) 
    app.title('AUTO-ATTENDENCE') 
    root=app
    app.mainloop()  




