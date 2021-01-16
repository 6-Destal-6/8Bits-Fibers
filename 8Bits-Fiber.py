# coding: utf-8
import tkinter as tk
import tkinter.font as tkFont
import time
import random
import sys,os

from PIL import Image, ImageTk
from ChangeColor import ChangeColor

LeaveColor  = "gray6"  # gris foncé
OnColor     = "gray8"  # gris moyen
bgColor     = "gray10" # gris clair

colorFont   = "white"
fontApp     = "Lucida Grande"


class Application(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs) 

        container = tk.Frame(self, background=bgColor )
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)     

           
        scWidth    = container.winfo_screenwidth()    # Largeur
        scHeight   = container.winfo_screenheight()   # Hauteur
        screenSize = ("{}x{}".format( int(scWidth*0.5) , int(scHeight*0.8)))

        # Taille de l application
        self.geometry(screenSize)
        self.minsize(int(scWidth*0.4) , int(scHeight*0.4))
        self.title("8-Bit:Fiber")

        # Taille de la grille : 50 par 50
        for r in range(50):
            container.rowconfigure    ( r, weight=1)
            container.columnconfigure ( r, weight=1)

        # Placement des differents elements sur la grille
        topp_Frame   = tk.Frame(container, background=OnColor  , highlightthickness=0, bd=0, relief=tk.SUNKEN )        
        down_frame   = tk.Frame(container, background=OnColor  , highlightthickness=0, bd=1, relief=tk.SUNKEN )
        centralFrame = tk.Frame(container, background=OnColor  , highlightthickness=0, bd=1, relief=tk.SUNKEN )
        left_frame   = tk.Frame(container, background=LeaveColor  , highlightthickness=0, bd=1, relief=tk.SUNKEN )
        #-------------------------------------------------------
        
        #-------------------------------------------------------  
        topp_Frame.pack  ( side = "top"   , expand=0, fill= "both", ipady=50 ) 
        down_frame.pack  ( side = "bottom", expand=0, fill= "both", ipady=20  )  
        left_frame.pack  ( side = "left"  , expand=0, fill= "both", pady=50, padx=50, ipadx=100 ) 
        centralFrame.pack( side = "left"  , expand=1, fill= "both", pady=50, ipadx=35 ) 

        # definit le chemin jusqu au program
        dirname = os.path.dirname(os.path.abspath(__file__))

        self.image1 = tk.PhotoImage(file = os.path.join(dirname, str('logo/rectangle.png'   ) ))
        self.image2 = tk.PhotoImage(file = os.path.join(dirname, str('logo/info.png'        ) ))

        canvas = tk.Canvas( topp_Frame , width = 200, height = 5, bg = LeaveColor, bd=0, highlightthickness=0 , relief=tk.SUNKEN)
        canvas.pack( side = "left", expand=1, fill= "both" )    

        def move(texte):

            canvas.delete("all")

            #print("4.0 - Fonction Move - import Image 1")  
            self.imageFinal1 = canvas.create_image( int(scWidth*0.69)  , 55, image = self.image1)

            #print("4.1 - Fonction Move - import Image 1")  
            self.imageFinal2 = canvas.create_image( int(scWidth*0.483) , 55, image = self.image2)    

            #print("4.2 - Fonction Move - Texte")  
            self.imageTexte  = canvas.create_text ( int(scWidth*0.56) + int(len(texte) ) , 55 , fill="white",font="Arial 15 bold", text=texte)

            nbr = 28

            #print("4.3 - Fonction Move - Aller") 
            for x in range(nbr)  :        

                canvas.move(self.imageFinal1, -x, 0)
                canvas.move(self.imageFinal2, -x, 0)
                canvas.move(self.imageTexte,  -x, 0)
                canvas.update()
                time.sleep(0.01)

                if x == nbr-1 :

                    time.sleep(1)

                    # print("4.4 - Fonction Move - Retour") 
                    for x in range(nbr)  :

                        canvas.move(self.imageFinal1, +x, 0)  
                        canvas.move(self.imageFinal2, +x, 0)  
                        canvas.move(self.imageTexte,  +x, 0)
                        canvas.update()
                        time.sleep(0.01)
        
        
        from Windows_Orange import OngletOrang
        OngletOrang(centralFrame, left_frame, scHeight, scWidth, fontApp, LeaveColor, OnColor, bgColor, move ).OrangFrame()

        # Fonction Reset --------------------------------------------------------------
        def refresh():
            os.execl(sys.executable, 'python', __file__)

        btnRefresh = tk.Button(down_frame, text="R", bg="#ff5050", fg="white", height=0, border=0, command=refresh)
        btnRefresh.place(relx = 1, rely = 1, anchor = tk.SE, relwidth=0.1, relheight =1)
        ChangeColor(btnRefresh, "#ff5050", "#ff7c80" )
        
        # Fonction Quitter  -----------------------------------------------------------         
        btnQuitter = tk.Button(down_frame, text="Quitter", bg=LeaveColor, fg="white", height=0, border=0, command=self.destroy)
        btnQuitter.place(relx = 0, rely = 1, anchor = tk.SW, relwidth=0.9, relheight =1)
        ChangeColor(btnQuitter,LeaveColor , OnColor )
        # -----------------------------------------------------------------------------      
       
    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()    

if __name__ == "__main__":
    app = Application()
    app.attributes('-alpha', 0.99)
    app.mainloop()