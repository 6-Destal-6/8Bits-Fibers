# coding: utf-8
from tkinter import *
import tkinter.font as tkFont
import time
import random
import sys,os

from PIL import Image, ImageTk
from ChangeColor import ChangeColor

LeaveColor  = "gray9"  # gris foncé
OnColor     = "gray6"  # gris moyen
colorFont   = "white"
fontApp     = "Lucida Grande"

class Application(Tk):

    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)         

        # Taille de l application
        container  = Frame(self, bg=OnColor )        
        scWidth    = container.winfo_screenwidth()    # Largeur
        scHeight   = container.winfo_screenheight()   # Hauteur
        screenSize = ("{}x{}".format( int(scWidth*0.5) , int(scHeight*0.8))) 

        container.pack(side=TOP, fill=BOTH, expand=True) 

        self.geometry(screenSize)
        self.minsize(int(scWidth*0.4) , int(scHeight*0.4))
        self.title("8 Bits Fibers")

        # Taille de la grille : 50 par 50
        for r in range(50):
            container.rowconfigure    ( r, weight=1)
            container.columnconfigure ( r, weight=1)

        # Placement des differents elements sur la grille
        topp_Frame   = Frame(container, bg=LeaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )        
        down_frame   = Frame(container, bg=LeaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
        centralFrame = Frame(container, bg=LeaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
        left_frame   = Frame(container, bg=LeaveColor   , highlightthickness=0, bd=1, relief=SUNKEN )
               
        down_frame.pack  ( side = "bottom", expand=0, fill= BOTH , ipady=20  ) 
        topp_Frame.pack  ( side = "top"   , expand=0, fill= X    , ipady=10)                 
        left_frame.pack  ( side = "left"  , expand=0, fill= BOTH , pady=50, padx=25, ipadx=100 )         
        centralFrame.pack( side = "left"  , expand=1, fill= BOTH , pady=50, padx=25 ) 
        #-------------------------------------------------------  
         
        helvetica = tkFont.Font(family='Arcade', size=70, weight='bold')
        Label(topp_Frame,text="8 Bits Fibers", fg="cyan", font=helvetica, bg=LeaveColor).pack(expand=1, fill= X)     
        
        from Windows_Orange import OngletOrang
        OngletOrang(centralFrame, left_frame, scHeight, scWidth, fontApp, LeaveColor, OnColor).OrangFrame()

        # Fonction Reset --------------------------------------------------------------
        def refresh():
            os.execl(sys.executable, 'python', __file__)

        btnRefresh = Button(down_frame, text="R", bg="#ff5050", fg="white", height=0, border=0, command=refresh)
        btnRefresh.place(relx = 1, rely = 1, anchor = SE, relwidth=0.1, relheight =1)
        ChangeColor(btnRefresh, "#ff5050", "#ff7c80" )
        
        # Fonction Quitter  -----------------------------------------------------------         
        btnQuitter = Button(down_frame, text="Quitter", bg=LeaveColor, fg="white", height=0, border=0, command=self.destroy)
        btnQuitter.place(relx = 0, rely = 1, anchor = SW, relwidth=0.9, relheight =1)
        ChangeColor(btnQuitter,LeaveColor , OnColor )                 

if __name__ == "__main__":
    app = Application()
    app.attributes('-alpha', 0.99)
    app.mainloop()