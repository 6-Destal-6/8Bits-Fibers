# coding: utf-8

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from os import walk
from os import path
from os import rename


from dbfread import DBF
from PIL import Image, ImageTk

import pandas as pd
import xlrd3 as xlrd
import time
import tkinter.font as tkfont


import tkinter, win32api, win32con, pywintypes

# import les processing windows
import win32com.client as win32

from functools import partial

class OngletOrange :    

    def __init__(self, centralFrame, left_frame, screenHeight, screenWidth, fontApp, LeaveColor, OnColor, bgColor, move ):

        self.centralFrame = centralFrame
        self.left_frame   = left_frame
        self.screenHeight = screenHeight
        self.screenWidth  = screenWidth
        self.fontApp      = fontApp
        self.LeaveColor   = LeaveColor  # gris foncé
        self.OnColor      = OnColor     # gris clair
        self.bgColor      = bgColor
        self.move         = move

    def OrangeFrame(self):

        colorRed    = "salmon"
        colorGreen  = "SeaGreen3"
        colorBlue   = "cornflower blue"
        colorOrange = "tan1"

        global incAerien
        incAerien = 0

        def Actived(Bouton):
            if Bouton["state"] == DISABLED:
                Bouton["state"] = NORMAL

        def Disabled(Bouton) :
            if Bouton["state"] == NORMAL:
                Bouton["state"] = DISABLED

        def RecupExcelPatch(directory, File, OFile , sheetRange): 

            if "C3B" in File or "C6" in File  :
                c3bFile  = (directory+"/"+File)
                book     = xlrd.open_workbook(c3bFile)
                sh       = book.sheet_by_index(3)
                return sh

            elif "C3A" in File or "C7" in File :  
                c3bFile  = (directory+"/"+File)
                book     = xlrd.open_workbook(c3bFile)
                sh       = book.sheet_by_index(1)
                return sh            

        def JusteOuFaux(faute) :

            if faute == 0 :
                titre = LABELBORDEREAU( scrollable_frame , "black" , u" ☺ Cool tout est juste ☺ " )
                titre['bg']   = colorGreen
                titre['fg']   = "black"

            else :   
                titre = LABELBORDEREAU( scrollable_frame , "black" , u" Merci de Corriger" )
                titre['bg']   = colorRed
                titre['fg']   = "black"  

        def LABELTITRE( texte ):
            labelTitre = Label( scrollable_frame, text=texte )
            labelTitre.pack(side = TOP, expand=0, fill=X)
            labelTitre.configure(font=("Helvetica", 11, "normal"), fg="bisque", bg=self.LeaveColor ) 

        def LABELBORDEREAU( Parent , color , texte ):
            Titre = Label(Parent, bg="SeaGreen1",fg="white",text= texte , highlightthickness=0, relief=FLAT, activebackground="brown2" )
            Titre.pack(expand=0, fill="x")
            Titre.configure(font=( "Helvetica", 10, "bold" ), fg="white" , bg=color)   
            return Titre

        def LABELRESULTAT( taille, couleur, epais, texte ):
            labetdirectory = Label( scrollable_frame, text=texte, anchor="w" )
            labetdirectory.pack(side = TOP, expand=0, fill=X)
            labetdirectory.configure(font=("Helvetica", taille, epais), fg=couleur, bg=self.LeaveColor )   

        def FuncUpFile():            
 
            try : 
                try :
                    global Racine                                      
                    Racine = filedialog.askdirectory(initialdir=r"C:\Users\Arnaud_2018\Desktop\DOE-Orage",title='Choisissez un repertoire')
                except:            
                    Path = path.dirname(path.abspath(__file__))                                       
                    Racine = filedialog.askdirectory(initialdir=Path,title='Choisissez un repertoire')

            except :
                print ("FuncUpFile : Annulée")

            if Racine != "" :

                fauxGCB_1 , fauxSiren , fauxNumFCI = 0 , 0 , 0

                for directory, dirnames, filenames  in walk(Racine, topdown=False):                     

                    Name = path.basename(directory)                    

                    def ChoiVraiFaux( val1eur, val2eur , text, increment ):
                        increment   = 0 
                        if Name[ val1eur : val2eur ] == text :   
                            pass                           
                        else :
                            increment +=1

                    # exemple : GCB_1_830959771_F98092200220_FII-88-025-287-DIS_DFT_V2

                    # Verifie que le code commence par GCB_1
                    if path.basename(directory) != "Appui Aérien" and path.basename(directory) != "Relevé de chambre" and len(path.basename(directory)) != len(str("F66074040620_88516") ):
                                                
                        ChoiVraiFaux( 0 , 6  ,"GCB_1_"      , fauxGCB_1  )   # Verifie que le code commence par GCB_1                        
                        ChoiVraiFaux( 6 , 15 ,"830959771"   , fauxSiren  )   # Verifie le numéro de Siren
                        ChoiVraiFaux( 15, 16 ,"_"           , fauxNumFCI )   # Verifie le numéro FCI
                        ChoiVraiFaux( 16, 17 ,"F"           , fauxNumFCI )   # Verifie le numéro FCI
                        ChoiVraiFaux( 28, 29 ,"_"           , fauxNumFCI )   # Verifie le numéro FCI

                        # Verifie si Command d'accès ou dossier de fin de travaux
                        if Name[len(Name)-7:len(Name)-1] == "_DFT_V":
                            Titre = LABELBORDEREAU( canvas, colorGreen , Name) 
                            Titre.configure(font=( "Helvetica", 10, "bold" ), fg="black")   

                        elif Name[len(Name)-6:len(Name)-1] == "_CA_V" or Name[len(Name)-3:len(Name)] == "_CA":
                            LABELBORDEREAU( canvas,"RoyalBlue1" , Name) 

                        elif "Nro" in Name :
                            LABELBORDEREAU( canvas,"maroon2" , Name)

                        else:
                            LABELBORDEREAU( canvas, colorOrange , Name) 

                # Active les Boutons -----------------------  
                Actived( ButtC3B   )
                LabelBoutonValide(FrameC3B , 0.25 )

                listButton = [ShapeCableButton, ShapeSupportButton ,ShapeBpeButton ]
                for increment in range(1 ,len(listButton)+1):
                    LabelBoutonValide(FrameShape , increment/4 )
                    Actived( listButton[increment-1] )
                # Active les Boutons -----------------------  
                
                self.move(u"Un dossier vient d'être Monté")  

        def FunCheckC3B():

            AlveolAccept = []

            for n in range(1, 21) :
                AlveolAccept.append( "A{}".format( n )  )
                AlveolAccept.append( "B{}".format( n )  )
                AlveolAccept.append( "C{}".format( n )  )
                AlveolAccept.append( "D{}".format( n )  )

            # print ( AlveolAccept ) 
            LABELRESULTAT( 9, colorGreen , "normal", "\n" )          

            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # Sort les fichies de la liste de fichier                
                for File in sorted(filenames ) : 

                    if "C3B" in File or "C3A" in File :

                        # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx
                        try :     
                            sh = RecupExcelPatch( directory , File, "C3A", 1 )
                        except:
                            sh = RecupExcelPatch( directory , File, "C3B", 3 )

                        # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                        depart = 11    # Debute appartir de la ligne 11 
                        for rx in range( depart , sh.nrows ) :
                            
                            Alveol = str( sh.cell(rowx=rx, colx=1).value)
                            TypeA  = str( sh.cell(rowx=rx, colx=2).value)
                            NumA   = str( sh.cell(rowx=rx, colx=3).value)
                            if NumA == "" :
                                NumA = u"<vide>"

                            TypeB  = str( sh.cell(rowx=rx, colx=4).value)                            
                            NumB   = str( sh.cell(rowx=rx, colx=5).value)
                            if NumB == "" :
                                NumB = u"<vide>"

                            if len(NumA) < 9 :
                                Value = "►\t" + "Ligne : " + str(rx+1).zfill(3) + "\t\t" + "Alveole: "+Alveol + "\t" + "Type A = "+TypeA + "\t" + NumA + "\t\t" + "Type B = "+ TypeB  + "\t" + NumB
                            
                            else :
                                Value = "►\t" + "Ligne : " + str(rx+1).zfill(3) + "\t\t" + "Alveole: "+Alveol + "\t" + "Type A = "+TypeA + "\t" + NumA + "\t" + "Type B = "+ TypeB  + "\t" + NumB

                            if TypeA == "C"  :

                                if str(Alveol) not in AlveolAccept :
                                    val = str( "►\t" + "Ligne : " + str(rx).zfill(3) + "\t\tAlveol Fausse: " + str( Alveol ) )
                                    LABELRESULTAT( 9, colorRed , "normal", val )

                                if   TypeB == "C" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "IMB" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "F" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "P" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )                        
                            
                                elif TypeB == "PT" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "A" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "AT" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                else :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )

                            elif TypeA == "CT"  :
                                if   TypeB == "P" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )
                                
                                elif   TypeB == "A" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                else :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )

                            elif TypeA == "A"  :
                                if   TypeB == "IMB" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "A" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "AT" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "F" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "P" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "PT" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                else :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )

                            elif TypeA == "AT"  :
                                if   TypeB == "P" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif TypeB == "A" :
                                    LABELRESULTAT( 9, colorOrange , "normal", "\n" + Value + u"\n  Toléré, cependant : Type A = A  et  Type B = AT serait plus juste ► inverser les 2 supports si C3A impossible en C3B\n")

                                else :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )

                            elif TypeA == "P"  :
                                if   TypeB == "F" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "P" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                elif   TypeB == "PT" :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                                else :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )
                            
                            else :
                                LABELRESULTAT( 9, colorGreen , "normal",  Value )

            self.move(u"Check C3A ou C3B à été réalisé")         
        
        # Fonction qui analyse le dossier Appui Aérien ----------------------------------------------------------------------------------------
        def CheckAppuiAerien(): 

            # espace
            LABELRESULTAT( 10, "white", "normal", u"\n" )   

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , colorBlue , "Analyse du dossier Appui Aérien"  )

            global C3Blist 
            C6List          = []
            PoteauxFauxList = []
            faute           = 0
            increment       = 0  
            inc             = 0 

            # Retrouve tous les chemins , les dossiers et les fichiers de Racine
            for directory, dirnames, filenames  in walk(Racine):                 

                if path.basename(directory) == u"Appui Aérien" :

                    # Sort les fichies de la liste de fichier                
                    for File in sorted(filenames ) :
                        
                        if "C6" in File :  
                            sh = RecupExcelPatch( directory , File, "C6", 3 )

                            # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                            depart = 8
                            for rx in range( 8 ,sh.nrows) :    
                                Support     = str( sh.cell( rowx=rx, colx=0  ).value)
                                PoteauxFaux = str( sh.cell( rowx=rx, colx=29 ).value)

                                if str(Support) != "" :
                                    C6List.append( str(Support) )

                                    if PoteauxFaux != "" :
                                        PoteauxFauxList.append( str(Support) )

                            # Affiche le nombre de calbes present dans la C3Bs
                            texte = "Il y a "+ str(len(C6List )) + " Poteaux dans "+ File
                            LABELBORDEREAU( scrollable_frame, "black", texte )

                            listGespot = []  
                            inc+=1

            for Element in sorted(C3Blist) : 

                firstSepared  = Element.split(' ')[0]
                SecondSepared = Element.split(' ')[2]

                if firstSepared == "Poteau" :
                    # print (  "C6List : " + str(C6List)   )
                    resultat = SecondSepared.split("/")
                    # print ( "resultat : " + str(resultat[1] ) )

                    if    resultat[1] not in   C6List : 
                        # print ( firstSepared + " " +  SecondSepared )
                        LABELRESULTAT( 9, colorRed , "bold", u"►\t" + str(firstSepared) + " " + str(SecondSepared) + u"\test présent dans la C3B mais absent de la C6" )

                    else :
                        # print ( firstSepared + " " +  SecondSepared )
                        LABELRESULTAT( 9, colorGreen, "bold", u"►\t" + str(firstSepared) + " " + str(SecondSepared) + u"\t Ok" )

                elif "GESPOT" in File :
                    print ("GESPOT : " + str(File[7:-5] ) )             
                    listGespot.append( File[7:-5] )

            # Affiche le nombre de calbes present dans la C3Bs
            LABELBORDEREAU( scrollable_frame, "black", "Il y a besoin de " + str(len(PoteauxFauxList) ) + " fichiers Gespot dans le dossier Appui Aérien" )

            for Element in PoteauxFauxList  :       

                if Element  not in listGespot :
                    faute += 1
                    LABELRESULTAT( 9, colorRed , "bold", u"►\t" + Element + u"\test absent ► Ce poteau à besoin d'un Gespot" )
                else :
                    LABELRESULTAT( 9, colorGreen, "bold", u"►\t" + Element + u"\t Ok" )            

            if inc == 0 :
                Titre = LABELBORDEREAU( scrollable_frame , colorRed , "Le dossier < Appui Aérien > est vide" )
                Titre['fg'] = "black"
                LABELRESULTAT( 9, colorGreen, "normal", "" )

            # Retrouve tous les chemins , les dossiers et les fichiers de Racine
            for directory, dirnames, filenames  in walk(Racine):
                for File in filenames :
                    if str(len(PoteauxFauxList) ) != 0 :   
                        if "C7.xlsx" in File :
                            LABELBORDEREAU( scrollable_frame , "black" , File )

                            # print (File )   
                            sh = RecupExcelPatch( directory , File, "C7", 1 )

                            # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                            depart = 17
                            for rx in range(depart,sh.nrows):  

                                SUPPORT     = str( sh.cell( rowx=rx, colx=0  ).value)
                                
                                if  SUPPORT not in PoteauxFauxList :
                                    LABELRESULTAT( 9, colorRed, "bold", u"►\t" + SUPPORT + u"\t ◄ Poteaux Inconnu" )
                                    faute +=1  
                                else :
                                    LABELRESULTAT( 9, colorGreen, "bold", u"►\t" + SUPPORT ) 

                            increment = 1
                        elif "C7.pdf" in File :
                            LABELBORDEREAU( scrollable_frame , "black" , File + " Impossible d'analyser le Pdf")
                            increment = 1

            if increment == 0 :
                faute +=1 

            JusteOuFaux(faute)

            self.move(u"Check du Dossier Appui Aérien") 

        # Fonction qui analyse le dossier Appui Aérien ----------------------------------------------------------------------------------------


        # Fonction qui analyse le dossier Relevé de Chambre -----------------------------------------------------------------------------------
        def CheckReleveChambre() :

            global mylistFOA
            global C3Blist

            mylistFOA           = []
            addNexList          = []
            addC3BChambreList   = []
            NewElementFOA       = []

            Actived(ButNameC16)            
            LabelBoutonValide(FrameChambre , 0.55 )

            for element in  C3Blist:
                addNexList.append( element.split(" ")[2] ) # Renvoie > 88516/152

            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # S'arrete sur le dossier "Relevé de chambre"
                if path.basename(directory) == u"Relevé de chambre" :

                    # Renvoie le Titre
                    LABELTITRE( str("\n\nAnalyse du dossier : " + path.basename(directory) )  )

                    # Affiche le nombre de fichier FOA dans ce dossier
                    LABELBORDEREAU( scrollable_frame , "black" , " ( il y a "+ str(len(filenames )) + " fichiers FOA )"  )

                    # Boucle sur la liste du DBF pour sortir chaque Element #
                    for File in sorted(filenames ) :

                        Newfilenames  = File[:len(File)-9].replace("-", "/")
                        #print ( "filenames  = " + Newfilenames  )

                        if File[len(File)-4:] == ".xls": 
                            #print ( "xls : " + File[len(File)-8:-4] )
                            mylistFOA.append( File[:len(File)-8] )

                            if File[len(File)-8:-4] == "_C16" :
                                LABELRESULTAT( 9, "white", "normal", File ) 
                            else:
                                LABELRESULTAT( 9, "#ff5050", "bold", File + str(" ◄ Manque _C16 à la fin") )

                        elif File[len(File)-4:] == "xlsx":

                            if Newfilenames  in addNexList :
                                mylistFOA.append( File[:len(File)-9] ) 

                                if File[len(File)-9:-5] == "_C16" :
                                    if len(File) < 18 :
                                        LABELRESULTAT( 9, colorGreen, "bold", File + str("\t\t◄ Nécessaire donc OK") ) 
                                    else:
                                        LABELRESULTAT( 9, colorGreen, "bold", File + str("\t◄ Nécessaire donc OK") ) 

                                else:
                                    LABELRESULTAT( 9, "#ff5050", "bold", File + str("\t◄ Nécessaire : Manque _C16 à la fin") )
                                    LabelBoutonValide(FrameChambre , 0.55 ) 
                                    Actived(ButNameC16) 

                            else :
                                mylistFOA.append( File[:len(File)-9] ) 
                                
                                if File[len(File)-9:-5] == "_C16" :
                                    if len(File) < 9 :
                                        LABELRESULTAT( 9, "gray63", "normal", File + str("\t◄ Pas Nécessaire à Retirer du dossier \"Relevé de Chambre\"") )
                                    else:
                                        LABELRESULTAT( 9, "gray63", "normal", File + str("\t\t◄ Pas Nécessaire à Retirer du dossier \"Relevé de Chambre\"") )

                                else:
                                    LABELRESULTAT( 9, "#ff5050", "bold", File + str("\t◄ Manque _C16 à la fin") )
                                    LabelBoutonValide(FrameChambre , 0.55 ) 
                                    Actived(ButNameC16) 

            for element in ChambreC3Blist :
                NewElement = element.split(" ")[2]
                if element.split(" ")[0] == "Chambre" :
                    addC3BChambreList.append(NewElement)

            for element in mylistFOA :
                NewElement = element.replace("-" , "/")
                NewElementFOA.append( NewElement )

            # retire les doublons
            NewElementFOA     = list(dict.fromkeys(NewElementFOA))       
            addC3BChambreList = list(dict.fromkeys(addC3BChambreList))

            # print ("NewElementFOA : " + str(NewElementFOA) )
            # print ("addC3BChambreList : " + str(addC3BChambreList) )
            i=0

            dirname = path.dirname( path.abspath(__file__))
            file    = path.join( dirname, 'RapportFoaManquant.txt' )
            my_file = open( file ,'w+' )

            for element in sorted(addC3BChambreList)   :

                if element not in NewElementFOA :
                    i+=1
                    # print ( "manquant : " + str(element) )
                    LABELRESULTAT( 9, colorRed, "bold", "Le fichier FOA ( " + element + str(" )\t est manquant dans le dossier \"Relevé de Chambre\"") )                    
                    my_file.write( element + "\r")

            my_file.close()  

            self.move(u"Check du Dossier Relevé de Chambre") 
        # Fonction qui analyse le dossier Relevé de Chambre -----------------------------------------------------------------------------


        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        def CheBPE() :  

            # Initialisation des Listes
            C3Blist          = []            
            supportListShape = []
            NewSupportC3Blist= []
            supportC3Blist   = []
            
            inc              = 0
            nombreDeBpeShape = 0 
            faute            = 0

            def supportAvecBoitier( interger, Type ):

                nombreDeBPE      = 0
                supportC3Blist   = []

                if Boitier[0] == Type :

                    Cable   = str( sh.cell(rowx=rx, colx=15).value)
                    support = str( sh.cell(rowx=rx, colx=interger).value)                    

                    if Cable != u"Câble non posé" :
                        supportC3Blist.append(  support + " \t " + Boitier )
                        C3Blist.append( support + " \t " + Boitier )
                        nombreDeBPE += 1

                return nombreDeBPE , C3Blist , supportC3Blist

            # espace
            LABELRESULTAT( 10, "white", "normal", u"\n" ) 

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , colorBlue , "Analyse du nombres de boitiers"  )
                      
            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # Sort les fichies de la liste de fichier                
                for File in sorted(filenames ) : 

                    # S'arrete sur le fichier Fxxxxxxxxxxx_C3B.xlsx
                    if "C3B" in File :

                        Fxxxxxxxxxxx_C3Bxlsx = File

                        # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx      
                        sh = RecupExcelPatch( directory , File, "C3B" , 3 )

                        # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                        depart = 11    # Debute appartir de la ligne 11 
                        for rx in range(depart,sh.nrows):                           
                            
                            Boitier = str( sh.cell(rowx=rx, colx=14).value)                       
                            
                            if Boitier != "" :

                                nombreDeBPE, C3Blist, supportC3Blist = supportAvecBoitier( 3, "A" )
                                nombreDeBPE, C3Blist, supportC3Blist = supportAvecBoitier( 5, "B" )

                    # S'arrete sur le fichier support.dbf
                    elif "bpe.dbf" in File :
                        nombreDeBpeShape = 0
                        dbfFile  = (directory+"/"+File) 
                        for record in DBF(dbfFile):
                            nombreDeBpeShape +=1                            

                        # Affiche le nombre de cables present dans la C3Bs
                        if nombreDeBpeShape == 0:
                            pass
                        else:
                            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeBpeShape) + " supports ayant un boitier dans le shape : " + File + "\t" + path.basename(directory)  )

                            for record in DBF(dbfFile):
                                Value = str(record['id_support'])
                                if Value != '' :
                                    inc +=1                                    
                                    LABELRESULTAT( 10, "white", "normal", str( inc ).zfill(3) + " \t " + str( Value ) ) 

                                    # Tous les supports se trouvant dans le Shape
                                    supportListShape.append(Value)            

            # Affiche le Titre de la Fonction
            LABELTITRE( "\n\nAnalyse comparative entre le fichier " + Fxxxxxxxxxxx_C3Bxlsx + " et le Shape bpe.dbf"  )

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeBPE) + " supports ayant un boitier dans la C3B"  )

            for Element in sorted(C3Blist) :
                if Element != "" :                    
                    firstSepared  = Element.split(' ')[0]
                    FourSepared   = Element.split(' ')[3]
                    # print (  "Element : " + firstSepared + "Element : " + FourSepared      ) 

                    if len(firstSepared) < 9 :
                        LABELRESULTAT( 10, colorGreen, "bold", "Present dans la C3B : \t\t" + firstSepared + "\t\t" + FourSepared  ) 
                    else:
                        LABELRESULTAT( 10, colorGreen, "bold", "Present dans la C3B : \t\t" + firstSepared + "\t" + FourSepared ) 

            for Element in sorted(supportC3Blist) :            
                if Element not in C3Blist :
                    firstSepared  = Element.split(' ')[0]
                    SecondSepared = Element.split(' ')[1]
                    # print (  "Element : " + firstSepared + "Element : " + SecondSepared      ) 

                    if len(Element) < 8 :
                        LABELRESULTAT( 10, colorRed, "bold", "Absent de la C3B : \t\t"  + firstSepared + "\t\t" + SecondSepared ) 
                    else:
                        LABELRESULTAT( 10, colorRed, "bold", "Absent de la C3B : \t\t" + firstSepared + "\t" + SecondSepared ) 
            
            if inc == 0 :
                Titre = LABELBORDEREAU( scrollable_frame , colorRed , "Le fichier bpe.shp est manquant ou vide" )
                Titre['fg'] = "black"

            LABELRESULTAT( 10, colorGreen, "bold", "" ) 
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(len(supportListShape) ) + " supports ayant un boitier dans le Shape"  )

            for Element in supportC3Blist :

                firstSepared  = Element.split(' ')[0]
                NewSupportC3Blist.append(firstSepared)

                if firstSepared in supportListShape :
                    #print ("element trouvée : " + firstSepared)
                    LABELRESULTAT( 10, colorGreen, "bold", "Present dans bpe.shape : \t\t" + firstSepared )
           
                else :
                    faute += 1
                    #print ("element non trouvée : " + firstSepared)
                    LABELRESULTAT( 10, colorRed, "bold", "Absent dans bpe.shape : \t\t" + firstSepared ) 

            for Element in supportListShape :
                # print ("supportListShape" + str(Element) )
                if Element not in NewSupportC3Blist and  Element != '' :

                    #print ("element trouvée : " + firstSepared)
                    LABELRESULTAT( 10, colorOrange, "bold", "En trop dans bpe.shape : \t\t" + Element )

            JusteOuFaux(faute) 

            self.move(u"Check des BPE")
        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        

        # Fonction qui analyse et compare le shape support par rapport à la C3B ---------------------------------------------------------
        def CheckSupport() :
            global C3Blist     
            global ChambreC3Blist
            global incAerien     

            # Active le Bouton Relevé de Chambre
            Actived(ButtReleve)
            LabelBoutonValide(FrameChambre , 0.25 )
            LabelBoutonValide(FrameChambre , 0.50 )           

            def TypeSupport( ColonneType , ColonneSupport ):

                Type = str( sh.cell( rowx=rx, colx=ColonneType ).value)

                if Type == "C" :
                    Type = "Chambre"
                if Type == "A" :
                    Type = "Poteau"

                CABLE   = str( sh.cell(rowx=rx, colx=15).value)

                if CABLE != u"Câble non posé" :                    
                            
                    if Type != "F" and Type != "AT" :                                

                        SUP1ORT = str( sh.cell(rowx=rx, colx=ColonneSupport).value)
                        C3Blist.append( Type + " - " + SUP1ORT )

                        if Type != "A" and Type != "A" :

                            ChambreC3Blist.append( Type + " - " + SUP1ORT )

                return ChambreC3Blist

            # Initialisation des listes
            supportListC3B      = []
            supportListShape    = []
            C3Blist             = []            
            ChambreC3Blist      = []
            InseeListC3B        = []          
            nombreDeSupport     = 0
            increment           = 0
            nombreDeSupportShape= 0
            faute               = 0

            try :

                # espace
                LABELRESULTAT( 10, "white", "normal", u"\n" )   

                # Affiche le nombre de calbes present dans la C3B
                LABELBORDEREAU( scrollable_frame , colorBlue , u"Analyse des Supports"  )

                # Retrouve tous les chemins , les dossiers et les fichiers de Racine
                for directory, dirnames, filenames  in walk(Racine):

                    # Sort les fichies de la liste de fichier
                    for File in sorted(filenames ) :
                        nombreDeSupportShape= 0                      

                        # S'arrete sur le fichier Fxxxxxxxxxxx_C3B.xlsx
                        if "C3B" in File or "C3A" in File:
                            Fxxxxxxxxxxx_C3Bxlsx = File                        

                            # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx
                            try :     
                                sh = RecupExcelPatch( directory , File, "C3A", 1 )
                            except:
                                sh = RecupExcelPatch( directory , File, "C3B", 3 )                                 

                            for rx in range(11,sh.nrows):
                                TypeSupport( 2 , 3 )
                                TypeSupport( 4 , 5 )        

                        # S'arrete sur le fichier support.dbf
                        elif "support.dbf" in File :                        

                            # Affiche le Titre de la Fonction
                            LABELTITRE( "\n\nAnalyse comparative entre le fichier " + Fxxxxxxxxxxx_C3Bxlsx + " et le Shape support.dbf"  )                        
                            dbfFile  = (directory+"/"+File) 

                            for record in DBF(dbfFile):
                                nombreDeSupportShape +=1                            

                            # Affiche le nombre de cables present dans la C3Bs
                            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeSupportShape) + " support présent dans le shape" )

                            for record in DBF(dbfFile):
                                increment +=1
                                Value = str(record['id_support'])
                                LABELRESULTAT( 10, "white", "normal", str( increment ).zfill(3) + " \t " + str( Value ) ) 

                                # Tous les supports se trouvant dans le Shape
                                supportListShape.append(Value)                
                        
                # Affiche le Titre de la Fonction
                LABELTITRE( "\n\nAnalyse comparative entre le fichier " + Fxxxxxxxxxxx_C3Bxlsx + " et support.shp"  )

                C3Blist= list(dict.fromkeys(C3Blist)) # retire les doublons 
                for element in sorted(C3Blist) :
                    if element != "" :
                        nombreDeSupport += 1

                # Affiche le nombre de calbes present dans la C3Bs
                LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeSupport) + " supports présents dans la C3B" )

                for element in sorted(C3Blist) :

                    Newelement = element.split(" ")[2]
                    INSEE      = Newelement.split("/")[0]

                    InseeListC3B  .append(INSEE)
                    supportListC3B.append(Newelement)

                    if element != "" :
                        if "Chambre" in element :
                            LABELRESULTAT( 10, colorGreen, "bold", element )
                        else:
                            Actived( ButtAppui  )
                            LabelBoutonValide(FrameChambre , 0.75)
                            LABELRESULTAT( 10, colorBlue , "bold", element )

                if len(supportListShape) == 0 :
                    Titre = LABELBORDEREAU( scrollable_frame , colorRed , "Le fichier support.shp est manquant, vide ou zippé " )
                    Titre['fg'] = "black"
                    LABELRESULTAT( 9, colorGreen, "normal", "" )
                    faute += 1

                else :
                    # Affiche le nombre de calbes present dans la C3Bs
                    LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(len(supportListShape) ) + " supports présents dans le support.shp" )

                    for element in sorted(supportListC3B) :
                        # print ( "supportListC3B : " + str(element) )
                        if element not in supportListShape :
                            faute += 1
                            if len(element) < 8 :
                                LABELRESULTAT( 10, colorRed, "bold", "Support : " + element + u"\t\test présent dans la C3B mais absent de Support.shp" ) 

                            else :
                                LABELRESULTAT( 10, colorRed, "bold", "Support : " + element + u"\test présent dans la C3B mais absent de Support.shp" )         

                    for element in sorted(supportListShape) :
                        # print ( "supportListShape : " + str(element) )
                        if element not in supportListC3B :
                            faute += 1
                            if element == "" :
                                LABELRESULTAT( 10, colorOrange, "bold", u"Support : <vide> à supprimer du Shape" )

                            else :
                                LABELRESULTAT( 10, colorOrange, "bold", u"Support : " + element + " à supprimer du Shape" )  
                # ------------------------------------------------------------------------------------------------------------------------------------
                # ------------------------------------------------------------------------------------------------------------------------------------
                InseeListC3B= list(dict.fromkeys(InseeListC3B)) # retire les doublons 

                DictCommune = { "SAINT-BASLEMONT"   : 88411 , "THUILLIERES"      : 88472 , "SAINT-BASLEMONT" : 88411 , 
                                "VALLEROY LE SEC"   : 88490 , "VITTEL"           : 88516 , "SANDAUCOURT"     : 88440 ,
                                "DOMBROT-SUR-VAIR"  : 88141 , "BELMONT-SUR-VAIR" : 88051 , "NORROY"          : 88332 , 
                                "MONTHUREUX-LE-SEC" : 88309 , "HAREVILLE"        : 88231
                                }
               
                val_list = list(DictCommune.values())

                for commune in InseeListC3B :
                    if int(commune) in val_list :
                        print( u"valeur Trouvée : " + str(commune) )
                        LABELRESULTAT( 10, "pink1", "bold", u"Commune : " + str(commune) + " \tTrouvée")  

                    else :
                        print( u"valeur Inconnue : " + str(commune) )
                        LABELRESULTAT( 10, "orchid1", "bold", u"Commune : " + str(commune) + " \tInconnue")
                # ------------------------------------------------------------------------------------------------------------------------------------
                # ------------------------------------------------------------------------------------------------------------------------------------
                JusteOuFaux(faute)              
                      
            except Exception as e:
                # print("Il manque le Champ " + str( e )  + " dans le Shape Support ")
                LABELRESULTAT( 10, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )

            self.move(u"Check des Supports")           
        # Fonction qui analyse et compare le shape support par rapport à la C3B ---------------------------------------------------------

        # Fonction qui analyse et compare le shape cable par rapport à la C3B -----------------------------------------------------------
        def CheckCable(): 

            def ElementSeparee(Element):

                # Compare les Elements avec la liste DBFlistsansID
                firstSepared  = Element.split('-')[0]
                SecondSepared = Element.split('-')[1]
                ThirtSepared  = Element.split('-')[2]      
                # print ( "SecondSepared : " + str( firstSepared ) + str( len( firstSepared ) ) )
           
                return firstSepared, SecondSepared, ThirtSepared

            def ReconnaissanceSupport( interger ):
                Support = str( sh.cell(rowx=rx, colx=interger).value)

                if Support == "" :
                    Support = "SupportTiers"

                return Support
            
            def RecupereSupport(id, type ):

                idValue   = str(record[id])     # Récupere les Valeurs du champ id
                typeValue = str(record[type])   # Récupere les Valeurs du champ type

                # Si la Case est égale à AT remplacer la valeur par "SupportTiers"
                if typeValue == "AT" or typeValue == "F" :
                    idValue = "SupportTiers"

                if typeValue == "IMB" :
                    idValue = "Immeuble"

                idDiam = str(record['diam_cbl'])                                 

                return idValue, idDiam

            # Initialisation des Listes
            C3Blist         = []
            DBFlist         = []
            NroDBFlist      = []
            nombreDeCable   = 0
            inc             = 0

            # espace
            LABELRESULTAT( 10, "white", "normal", u"\n" )   

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , colorBlue , u"Analyse des Câbles"  )
           
            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # Sort les fichies de la liste de fichier
                for File in sorted(filenames ) :

                    if path.basename(directory) == u"Check" :  

                        if "cable.dbf" in File :
                            inc += 1                                                   
                            dbfFile = (directory+"/"+File) # Créer le Chemin vers le Shape

                            # Boucle sur chaque ligne du fichier
                            for record in DBF(dbfFile):

                                IDAValue, IDiamValue = RecupereSupport('id_a' , 'type_a' )
                                IDBValue, IDiamValue = RecupereSupport('id_b' , 'type_b' )

                                # Réunie Chaques Valeurs
                                Value = IDAValue + "-" + IDBValue + "-" + IDiamValue
                                NroDBFlist.append( Value )

                    else :

                        try :
                            # S'arrete sur le fichier Fxxxxxxxxxxx_C3B.xlsx
                            if "C3B" in File :

                                Fxxxxxxxxxxx_C3Bxlsx = File

                                # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx
                                sh = RecupExcelPatch( directory , File , "C3B", 3)  

                                # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                                depart = 11    # Debute appartir de la ligne 11                    
                                for rx in range( depart , sh.nrows):

                                    # S'arrete sur la Colonne P : Câble
                                    CABLE   = str( sh.cell(rowx=rx, colx=15).value)

                                    # Continue le Programme Uniquement si le Câble est Posé
                                    if CABLE != u"Câble non posé" :

                                        # Nombre de Câble
                                        nombreDeCable += 1  

                                        SUP1ORT = ReconnaissanceSupport(3)
                                        SUP2ORT = ReconnaissanceSupport(5)

                                        # S'arrete sur la Colonne K : Diamètre du Câble à Poser
                                        DIAM    = str( sh.cell(rowx=rx, colx=10).value)

                                        # Réunie Chaques Valeurs
                                        Value = SUP1ORT +"-"+ SUP2ORT +"-"+ DIAM
                                        C3Blist.append( Value )  
                        
                            # S'arrete sur le shape cable.dbf
                            elif "cable.dbf" in File :

                                inc += 1

                                # Créer le Chemin vers le Shape
                                dbfFile  = (directory+"/"+File) 

                                # Boucle sur chaque ligne du fichier
                                for record in DBF(dbfFile):

                                    IDAValue, IDiamValue = RecupereSupport('id_a' , 'type_a' )
                                    IDBValue, IDiamValue = RecupereSupport('id_b' , 'type_b' )

                                    # Réunie Chaques Valeurs
                                    Value = IDAValue + "-" + IDBValue + "-" + IDiamValue
                                    DBFlist.append( Value )

                        except Exception as e:
                            LABELRESULTAT( 10, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )   

            # Affiche le Titre de la Fonction
            LABELTITRE(  "\n\nAnalyse comparative entre le fichier " + File + " et cable.shp"  )
            
            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(len(NroDBFlist) ) + " câbles posées dans la C3B" )  

            j=0
            # Boucle sur la liste Résultat des cables à Garder
            for Element in sorted(NroDBFlist) :
                j+=1
                firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                if len( firstSepared ) < 9 :                        
                    if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                        LABELRESULTAT( 10, "pink" , "bold", str(j) + u"\tCâble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                    else :
                        LABELRESULTAT( 10, "pink", "bold", str(j) + u"\tCâble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )

                else :
                    if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                        LABELRESULTAT( 10, "pink", "bold", str(j) + u"\tCâble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                    else :
                        LABELRESULTAT( 10, "pink", "bold", str(j) + u"\tCâble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )
         
            # Affiche le Titre de la Fonction
            LABELTITRE(  "\n\nAnalyse comparative entre le fichier " + Fxxxxxxxxxxx_C3Bxlsx + " et cable.shp"  )
            
            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeCable) + " câbles posées dans la C3B" )            

            # Entete
            LABELRESULTAT( 10, "white", "bold", u"\t\t\t id_a \t\t\t   id_b\t\tdiam_cbl" )         

            # Boucle sur la liste Résultat des cables à Garder
            for Element in sorted(DBFlist) :                
                                
                # Compare les Elements avec la liste DBFlistsansID
                if Element not in C3Blist :
                    firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                    if len( firstSepared ) < 9 :                        
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorOrange, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorOrange, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )

                    else :
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorOrange, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorOrange, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )
                
                else :  
                    firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                    if len( firstSepared ) < 9 :
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorGreen, "bold", u"\tCâble à Garder : \t\t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorGreen, "bold", u"\tCâble à Garder : \t\t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )

                    else :
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorGreen, "bold", u"\tCâble à Garder : \t\t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorGreen, "bold", u"\tCâble à Garder : \t\t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )

            # Boucle sur la liste du DBF pour sortir chaque Element #
            for Element in sorted(C3Blist) :

                # Compare les Elements avec la liste DBFlistsansID
                if Element not in DBFlist :                    
                    firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                    if len(firstSepared) <9 :
                        if len( SecondSepared ) < 9 :
                            LABELRESULTAT( 10, colorRed, "bold", u"Le câble partant de : \t" + firstSepared + "\t\tvers \t" + SecondSepared + u"\t\tØ " + ThirtSepared + "\test présent dans la C3B mais absent de cable.shp" )

                        else :
                            LABELRESULTAT( 10, colorRed, "bold", u"Le câble partant de : \t" + firstSepared + "\t\tvers \t" + SecondSepared + u"\tØ "+ ThirtSepared + "\test présent dans la C3B mais absent de cable.shp" )

                    else :
                        if len( SecondSepared ) < 9 :
                            LABELRESULTAT( 10, colorRed, "bold", u"Le câble partant de : \t" + firstSepared + "\tvers \t" + SecondSepared + u"\t\tØ " + ThirtSepared + "\test présent dans la C3B mais absent de cable.shp" )

                        else :
                            LABELRESULTAT( 10, colorRed, "bold", u"Le câble partant de : \t" + firstSepared + "\tvers \t" + SecondSepared + u"\tØ "+ ThirtSepared + "\test présent dans la C3B mais absent de cable.shp" )

            if len(DBFlist) == 0 :
                Titre = LABELBORDEREAU( scrollable_frame , colorRed , "Le fichier cable.shp est manquant ou vide" )
                Titre['fg'] = "black"

            self.move(u"Check des Câbles") 
        # Fonction qui analyse et compare le shape cable par rapport à la C3B ---------------------------------------------------------------------


        # Renomme les fichiers 88516/1431.xlsx en 88516/1431_C16.xlsx -----------------------------------------------------------------------------  
        def Rename():

            def DectectionFinFile(interger):

                RenameCount = 0

                # Analyse la fin du fichier ► 88516/1431.xlsx
                if interger == 4 :
                    finFile = File [len(File)-8:len (File) - int(interger) ]   # _C16
                elif interger == 5 :
                    finFile = File [len(File)-9:len (File) - int(interger) ]   # _C16
                #print ( "finFile : ",finFile )
                
                # S'arrete sur le fichier si cette fin n'e"st pas égale à "_C16"
                if finFile != "_C16" :
                    try:
                        # Renomme la fin par _C16
                        olddirectory = directory + "/" + File

                        if interger == 4 :
                            newdirectory = directory + "/" + File[:-int(interger)] + "_C16.xls"
                        elif interger == 5 :
                            newdirectory = directory + "/" + File[:-int(interger)] + "_C16.xlsx"                            
                        else:
                            pass
                        # print ( "newdirectory5 : " + newdirectory )

                        rename(olddirectory,newdirectory)
                        RenameCount += 1

                    except:
                        print(u"Fichier déjà existant : " + str(File) )

                return RenameCount 

            # Retrouve tous les chemins , les dossiers et les fichiers de Racine
            for directory, dirnames, filenames  in walk(Racine):

                # Sort les fichies de la liste de fichier
                for File in sorted(filenames ) :

                    # S'arrete sur le Dossier Relevé de Chambre
                    if path.basename(directory) == u"Relevé de chambre" :     

                        if File[len(File)-4:] == ".xls": 
                            RenameCount = DectectionFinFile(4)

                        elif File[len(File)-5:] == ".xlsx":
                            RenameCount = DectectionFinFile(5)

            # comptabilise le nombre de fichiers renommés
            if RenameCount == 0 :
                LABELRESULTAT( 10, "white", "normal", "Tous les Fichiers de Relevé de chambre sont bien nommés") 
                Disabled( ButNameC16 )
            else :
                LABELRESULTAT( 10, "white", "normal", str(RenameCount) + " Fichiers ont été Renommés") 

            self.move(u"Renommage des fichiers Relevés FOA") 
        
        # Renomme les fichiers 88516/1431.xlsx en 88516/1431_C16.xlsx ----------------------------------------------------------------------------- 


        # Creation des Boutons --------------------------------------------------------------------------------------------------------------------              
        def Picture (Path) :            
            Size     = int(self.screenHeight*0.05)          # Definit la taille du logo par rapport a l'ecran            
            dirname  = path.dirname(path.abspath(__file__)) # definit le chemin jusqu au program            
            picture  = path.join(dirname, Path)             # Definit le Chemin du PNG
            original = Image.open(picture)                  # Ouvre le PNG               
            resized  = original.resize((Size, Size))        # le redimensionne    
            self.img = ImageTk.PhotoImage(resized)          # import l'image redimensionner             
            return self.img                                 # Retourne l'image    
        
        def CreaButton(parent, realY, image, numPage):

            # Creer un bouton
            Tool = Button(parent, image=image, border=0, bg=self.LeaveColor, highlightthickness=0, bd=1, relief=FLAT
                    , activebackground=self.LeaveColor, command= FuncUpFile, cursor="hand2" )            
            Tool.pack( expand=True, fill=BOTH ) # Affiche le bouton

            return Tool

        def LabelBoutonValide(Parent, Yposition):
            CheckLabelBoutton = Label(Parent, text=u"✔" )
            CheckLabelBoutton.place(relx = 0.8, rely = Yposition, anchor = W, relwidth=0.2, relheight =0.2)
            CheckLabelBoutton.configure(font=("Helvetica", 15, "bold"), fg="bisque", bg=colorGreen )              

        def Bouton(Parent, Text, Yposition, command ):
            helvetica = tkfont.Font(family='Helvetica', size=10, weight='bold')
            Butt = Button(Parent, bg=self.bgColor ,fg="bisque" ,text=Text ,highlightthickness=0, relief=FLAT, cursor="plus", pady=20  
                , activebackground=self.LeaveColor, command=command, state=DISABLED, font=helvetica )
            Butt.place(relx = 0, rely = Yposition, anchor = W, relwidth=1, relheight =0.2)    

            def LabelBoutonDisabled(Parent):
                CheckLabelBoutton = Label( Parent, text=u"✘" )
                CheckLabelBoutton.place(relx = 0.8, rely = Yposition, anchor = W, relwidth=0.2, relheight =0.2)
                CheckLabelBoutton.configure(font=("Helvetica", 15, "bold"), fg="bisque", bg="#ff5050" )  

            LabelBoutonDisabled( Parent) 

            return Butt  

        self.imageBalais    = Picture("logo/zelda.png")

        # Creer un bouton
        Tool = Button(self.centralFrame, image=self.imageBalais, border=0, bg=self.LeaveColor, highlightthickness=0, bd=1, relief=FLAT
                , activebackground=self.LeaveColor, command= '#', cursor="hand2" )        
        Tool.pack( expand=False, fill=Y, side=RIGHT ) # Affiche le bouton

        # Fonction Balais ------------------------------------------------------------------------------------------------------------------------- 
        container = Frame( self.centralFrame )   # Creation du container     

        # Creation de l'espace dessin
        canvas = Canvas(container, bg= self.LeaveColor , highlightthickness=0, bd=1, relief=SUNKEN ) 
        
        # Creation du scrollbar
        scrollbar        = Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame( canvas )
        scrollable_frame.bind( "<Configure>", lambda e: canvas.configure( scrollregion=canvas.bbox("all") ) )

        # Creation d'une fenetre dans l'espace dessin lié au scrollbar
        canvas.create_window((0, 0), window=scrollable_frame, anchor="w")
        canvas.configure(yscrollcommand=scrollbar.set)

        container.pack(               fill="both" , expand=True )
        canvas   .pack( side="left" , fill="both" , expand=True )          
        scrollbar.pack( side="right", fill="y"                  )    

        # Frame Monter ---------------------------------------------------------------------------------------------------------------------------
        FrameMonter = Frame(self.left_frame, bg=self.bgColor)
        FrameMonter.place(relx = 0, rely = 3/20, anchor = SW, relwidth=1, relheight =3/20)

        LABELBORDEREAU( FrameMonter , self.LeaveColor , "Monter Dossier" )
        
        ButtMonter = Bouton(FrameMonter , "#" , 0.5, FuncUpFile )
        ButtMonter['state'] = NORMAL

        image1    = Picture("logo/monter_6.png")
        CreaButton(FrameMonter, 0.1, image1, FunCheckC3B)

        # Bouton C3B + Shape ---------------------------------------------------------------------------------------------------------------------
        FrameC3B = Frame(self.left_frame, bg=self.LeaveColor )
        FrameC3B.place(relx = 0, rely = 9/20, anchor = SW, relwidth=1, relheight =2/10 )

        LABELBORDEREAU( FrameC3B , self.OnColor , "Check" )
        ButtC3B     = Bouton(FrameC3B, u"C3B ou C3A", 0.25 , FunCheckC3B ) # Analyse le Dossier Relevée de Chambre

        # Bouton C3B + Shape ---------------------------------------------------------------------------------------------------------------------
        FrameShape = Frame(self.left_frame, bg=self.LeaveColor )
        FrameShape.place(relx = 0, rely = 13/20, anchor = SW, relwidth=1, relheight = 2/10 )

        LABELBORDEREAU( FrameShape , self.OnColor , "C3B + Shape" )
        ShapeSupportButton = Bouton(FrameShape, u"Support\t" , 0.25, CheckSupport ) # comparaison entre C3B et le DBF du Shape Support
        ShapeCableButton   = Bouton(FrameShape, u"Câble\t"   , 0.50, CheckCable   ) # comparaison entre C3B et le DBF du Shape Cable
        ShapeBpeButton     = Bouton(FrameShape, u"BPE\t"     , 0.75, CheBPE       ) # comparaison entre C3B et le DBF du Shape BPE
        
        #  Bouton Sous-Dossier -------------------------------------------------------------------------------------------------------------------         
        FrameChambre = Frame(self.left_frame, bg=self.LeaveColor )
        FrameChambre.place(relx = 0, rely = 19/20 , anchor = SW, relwidth=1, relheight = 2/10 )

        LABELBORDEREAU( FrameChambre , self.OnColor , "Sous-Dossier" )
        ButtReleve  = Bouton(FrameChambre, u"Relevé de \nChambre" , 0.25, CheckReleveChambre )     
        ButNameC16  = Bouton(FrameChambre, u"Rename\n_C16\t"      , 0.50, Rename             )
        ButtAppui   = Bouton(FrameChambre, u"Appui\t\nAérien\t"   , 0.75, CheckAppuiAerien   )

        # Creation du Bouton ---------------------------------------------------------------------------------------------------------------------- 