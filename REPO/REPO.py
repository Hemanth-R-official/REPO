import tkinter
import tkinter.messagebox
import cv2
import pandas as pd
import find as f
import automate as a
import os
#Initial Display
   
def REPO():
    a.run()
    
def find():    
    f.run()

def e_n_d():
    response = tkinter.messagebox.askquestion("exit","Are sure you want to exit ?")
    if response==0:
        exit()
    else:
        exit()

def display():
    address=os.getcwd()
    address+="/frontpage.jpg"
    image = cv2.imread(address) 
    cv2.imshow('REPO',image)
    cv2.waitKey(2000)
    cv2.destroyAllWindows()

    #Selection window

    window = tkinter.Tk()
    window.title("REPO")
    top_frame = tkinter.Frame(window).pack()
    bottom_frame = tkinter.Frame(window).pack(side = "bottom")


        
    btn1 = tkinter.Button(top_frame, text = "REPO", fg = "green",command=REPO).pack() 
    btn2 = tkinter.Button(top_frame, text = "FIND", fg = "red",command=find).pack()
    #'fg or foreground' is for coloring the contents (buttons)
    #'side' is used to left or right align the widgets
    btn3 = tkinter.Button(bottom_frame, text = "EXIT", fg = "red",command=e_n_d).pack(side = "left")
    window.mainloop()

def main():
    display()

if __name__=="__main__":
    main()