# coding: utf-8
from tkinter import *
import tkinter.font as tkFont
import time
import random
import sys,os

from PIL import Image, ImageTk
from ChangeColor import ChangeColor
from applicationColor import *

fontApp     = "Lucida Grande"

class Application(Tk):

    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)         

        # Taille de l application
        self.container  = Frame(self, bg=OnColor )        
        scWidth    = self.container.winfo_screenwidth()    # Largeur
        scHeight   = self.container.winfo_screenheight()   # Hauteur
        screenSize = ("{}x{}".format( int(scWidth*0.5) , int(scHeight*0.8))) 

        self.container.pack(side=TOP, fill=BOTH, expand=True) 

        self.geometry(screenSize)
        self.minsize(int(scWidth*0.4) , int(scHeight*0.4))
        self.title("8 Bits Fibers")

        # Taille de la grille : 50 par 50
        for r in range(50):
            self.container.rowconfigure    ( r, weight=1)
            self.container.columnconfigure ( r, weight=1)

        # Placement des differents elements sur la grille
        topFrame     = Frame(self.container, bg=leaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )        
        downFrame    = Frame(self.container, bg=leaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
        centralFrame = Frame(self.container, bg=leaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
        leftFrame    = Frame(self.container, bg=leaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
               
        downFrame.pack  ( side = "bottom", expand=0, fill= BOTH , ipady=20  ) 
        topFrame.pack  ( side = "top"   , expand=0, fill= X    , ipady=10)                 
        leftFrame.pack  ( side = "left"  , expand=0, fill= BOTH , pady=50, padx=25, ipadx=100 )         
        centralFrame.pack( side = "left"  , expand=1, fill= BOTH , pady=50, padx=25 ) 
        #-------------------------------------------------------  
         
        helvetica = tkFont.Font(family='Arcade', size=70, weight='bold')
        Label(topFrame,text="8 Bits Fibers", fg="cyan", font=helvetica, bg=leaveColor).pack(expand=1, fill= X)     
        
        from Windows_Orange import OngletOrange
        OngletOrange( centralFrame, leftFrame, scHeight, scWidth, fontApp).orangFrame()

        # Fonction Reset --------------------------------------------------------------
        def refresh():
            os.execl(sys.executable, 'python', __file__)

        btnRefresh = Button(downFrame, text="R", bg="#ff5050", fg="white", height=0, border=0, command=refresh)
        btnRefresh.place(relx = 1, rely = 1, anchor = SE, relwidth=0.1, relheight =1)
        ChangeColor(btnRefresh, "#ff5050", "#ff7c80" )
        
        # Fonction Quitter  -----------------------------------------------------------         
        btnQuitter = Button(downFrame, text="Quitter", bg=leaveColor, fg="white", height=0, border=0, command=self.destroy)
        btnQuitter.place(relx = 0, rely = 1, anchor = SW, relwidth=0.9, relheight =1)
        ChangeColor(btnQuitter,leaveColor , OnColor )                 

if __name__ == "__main__":
    app = Application()
    app.attributes('-alpha', 0.99)
    app.mainloop()