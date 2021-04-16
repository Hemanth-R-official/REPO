import pandas as pd
import matplotlib.pyplot as plt
import os
from tkinter import filedialog
import tkinter as tk
import tkinter.messagebox

def locate():
	filepath = filedialog.askopenfilename()
	if filepath.endswith('.xlsx'):
		return filepath 
	else:
		tkinter.messagebox.showinfo("Alert !", "Unsupported File...")
		response=tkinter.messagebox.askquestion("Unsupported File","Do you want to relocate it ? or return to main menu.")
		if response=='yes':
			locate()	
		else:
			exit()

def run():
    file_path=locate()
    data=pd.read_excel(
        os.path.join(file_path),
        engine='openpyxl',
    )

    root = tk.Tk()
    canvas1 = tk.Canvas(root, width = 400, height = 600,  relief = 'raised')
    canvas1.pack()

    #canvas2 = tk.Canvas(root, width = 400, height = 300,  relief = 'raised')
    #canvas2.pack()

    label1 = tk.Label(root, text="Find Student Data",fg="green")
    label1.config(font=('helvetica', 14))
    canvas1.create_window(200, 25, window=label1)

    label2 = tk.Label(root, text='Enter Registration Number:',fg="red")
    label2.config(font=('helvetica', 12))
    canvas1.create_window(200, 100, window=label2)

    entry1 = tk.Entry(root) 
    canvas1.create_window(200, 140, window=entry1)

    def getregno():
        regno = entry1.get()
        
        label3 = tk.Label(root, text= "fetch your data....",font=('helvetica', 10))
        canvas1.create_window(200, 230, window=label3)
        i=int(regno)
        i-=1
        Q=data.at[i,'QUANTS']
        L=data.at[i,'LOGICAL']
        V=data.at[i,'VERBAL']
        N=data.at[i,'NAME']
        C=data.at[i,'CORRECT']
        W=data.at[i,'WRONG']
        E=data.at[i,'EMAIL']
        SA=data.at[i,'STANDING ARREARS']
        x=['QUANTS','LOGICAL','VERBAL']
        y=[Q,L,V]
        
        f1=plt.figure(1)
        plt.bar(x,y,color=['red','blue','green'])
        plt.title(N+"'s Technical Skills Report")

        f2=plt.figure(2)
        y=[C,W]
        pielables=["Correct","Wrong"]
        pieexplode=[0.2,0]
        plt.pie(y,labels=pielables,explode=pieexplode,shadow=True)

        label5 = tk.Label(root, text= "NAME : "+N,font=('helvetica', 12))
        canvas1.create_window(200, 310, window=label5)
        label6 = tk.Label(root, text= "REGISTRATION NUMBER : "+regno,font=('helvetica', 12))
        canvas1.create_window(200, 340, window=label6)
        label7 = tk.Label(root, text= "STANDING ARREARS : "+str(SA) ,font=('helvetica', 12),fg="RED")
        canvas1.create_window(200, 370, window=label7)
        label8 = tk.Label(root, text= "EMAIL : "+E,font=('helvetica', 12),fg="GREEN")
        canvas1.create_window(200, 430, window=label8)

        plt.show()

    button1 = tk.Button(root, text='FIND', command=getregno, bg='white', fg='red', font=('helvetica', 9, 'bold'))
    canvas1.create_window(200, 180, window=button1)


    root.mainloop()

def main():
    run()

if __name__=="__main__":
    main()