# coding: utf-8

from tkinter import *
from tkinter import filedialog
from os import walk
from os import path
from os import rename

from iteration_utilities import duplicates
from iteration_utilities import unique_everseen

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

class OngletOrang :    

    def __init__(self, centralFrame, left_frame, screenHeight, screenWidth, fontApp, LeaveColor, OnColor):

        self.centralFrame = centralFrame
        self.left_frame   = left_frame
        self.screenHeight = screenHeight
        self.screenWidth  = screenWidth
        self.fontApp      = fontApp
        self.LeaveColor   = LeaveColor  # gris foncé
        self.OnColor      = OnColor     # gris clair

    def OrangFrame(self):

        colorRed   = "salmon"
        colorGreen = "SeaGreen3"
        colorBlue  = "cornflower blue"
        colorOrang = "tan1"

        def Actived(Bouton):
            if Bouton["state"] == DISABLED:
                Bouton["state"] = NORMAL

        def Disabled(Bouton) :
            if Bouton["state"] == NORMAL:
                Bouton["state"] = DISABLED

        def RecupExcelPatch(directory, File, OFile ):

            c3bFile  = (directory+"/"+File)
            book     = xlrd.open_workbook(c3bFile)

            if "C3B" in File or "C6" in File  :  
                sh       = book.sheet_by_index(3)
            elif "C3A" in File or "C7" in File :  
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

        def LABELBORDEREAU( Parent , color , texte ):
            Titre = Label(Parent, bg="SeaGreen1",fg="white",text= texte , highlightthickness=0, relief=FLAT, activebackground="brown2" )
            Titre.pack(fill=X)
            Titre.configure(font=( "Helvetica", 10, "bold" ), fg="white" , bg=color)   
            return Titre

        def LABELRESULTAT( taille, couleur, epais, texte ):
            labetdirectory = Label( scrollable_frame, text=texte, anchor="w" )
            labetdirectory.pack(side = TOP, expand=1, fill=X)
            labetdirectory.configure(font=("Helvetica", taille, epais), fg=couleur, bg=self.LeaveColor )   

        def FuncUpFile():  
            
            try :
                global Racine                                      
                Racine = filedialog.askdirectory(initialdir=r"C:\Users\Arnaud_2018\Desktop\DOE-Orage",title='Choisissez un repertoire')
            except:            
                Path = path.dirname(path.abspath(__file__))                                       
                Racine = filedialog.askdirectory(initialdir=Path,title='Choisissez un repertoire')

            if Racine != "" :
                
                Actived( ButtonAnalyseC3B       )
                Actived( ButtonSupportShp       )
                Actived( ButtonCableShp         )
                Actived( ButtonBpeShp           )                 
                
                fauxGCB_1 , fauxSiren , fauxNumFCI = 0 , 0 , 0

                for directory, dirnames, filenames  in walk(Racine, topdown=False):                    

                    Name = path.basename(directory)                    

                    def ChoiVraiFaux( val1eur, val2eur , text, increment ):
                        increment   = 0 
                        if Name[ val1eur : val2eur ] == text :   
                            pass                           
                        else :
                            increment +=1

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

                        else:
                            LABELBORDEREAU( canvas, colorOrang , Name) 

        def CheckC3bXlsx():

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
                            sh = RecupExcelPatch( directory , File, "C3A" )
                        except:
                            sh = RecupExcelPatch( directory , File, "C3B" )

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

                            """ teste
                            listC  = ["C","IMB","F","P","PT","A","AT","" ]
                            listCT = ["P","A" ]
                            lsitA  = ["IMB","A", "AT","F","P","PT" ]
                            listAT = ["P","A" ]
                            listP  = ["F","P","PT" ]
                            """

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
                                    LABELRESULTAT( 9, colorOrang , "normal", "\n" + Value + u"\n  Toléré, cependant : Type A = A  et  Type B = AT serait plus juste ► inverser les 2 supports si C3A impossible en C3B\n")

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
                            sh = RecupExcelPatch( directory , File, "C6" )

                            # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                            depart = 8
                            for rx in range( depart ,sh.nrows) :    
                                Support     = str( sh.cell( rowx=rx, colx=0  ).value)
                                PoteauxFaux = str( sh.cell( rowx=rx, colx=29 ).value)

                                if str(Support) != "" :
                                    C6List.append( str(Support) )

                                    if PoteauxFaux != "" :
                                        PoteauxFauxList.append( str(Support) )

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame, "black", "Il y a "+ str(len(C6List)) + " Poteaux dans "+ File )

            listGespot = []  
            inc+=1

            for Element in sorted(C3Blist) : 

                firstSepared  = Element.split(' ')[0]
                SecondSepared = Element.split(' ')[2]

                if firstSepared == "Poteau" :
                    resultat = SecondSepared.split("/")

                    if resultat[1] not in C6List : 
                        LABELRESULTAT( 9, colorRed , "bold", u"►\t" + str(firstSepared) + " " + str(SecondSepared) + u"\test présent dans la C3B mais absent de la C6" )

                    else :
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
                            sh = RecupExcelPatch( directory , File, "C7" )

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
        # Fonction qui analyse le dossier Appui Aérien ----------------------------------------------------------------------------------------


        # Fonction qui analyse le dossier Relevé de Chambre -----------------------------------------------------------------------------------
        def CheckReleveChambre() :

            global mylistFOA
            global C3Blist

            mylistFOA           = []
            addNexList          = []
            addC3BChambreList   = []
            NewElementFOA       = []        

            for element in  C3Blist:
                addNexList.append( element.split(" ")[2] ) # Renvoie > 88516/152

            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # S'arrete sur le dossier "Relevé de chambre"
                if path.basename(directory) == u"Relevé de chambre" :

                    # Affiche le nombre de calbes present dans la C3B
                    LABELBORDEREAU( scrollable_frame , self.OnColor , u""  )
                    LABELBORDEREAU( scrollable_frame , colorBlue , u"Analyse du dossier Relevé de Chambres"  )

                    # Affiche le nombre de fichier FOA dans ce dossier
                    LABELBORDEREAU( scrollable_frame , "black" , " ( il y a "+ str(len(filenames )) + " fichiers FOA )"  )

                    # Boucle sur la liste du DBF pour sortir chaque Element #
                    for File in sorted(filenames ) :

                        Newfilenames  = File[:len(File)-9].replace("-", "/")

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
                                    Actived(ButtonRenameC16) 

                            else :
                                mylistFOA.append( File[:len(File)-9] ) 
                                
                                if File[len(File)-9:-5] == "_C16" :
                                    if len(File) < 9 :
                                        LABELRESULTAT( 9, "gray63", "normal", File + str("\t◄ Pas Nécessaire à Retirer du dossier \"Relevé de Chambre\"") )
                                    else:
                                        LABELRESULTAT( 9, "gray63", "normal", File + str("\t\t◄ Pas Nécessaire à Retirer du dossier \"Relevé de Chambre\"") )

                                else:
                                    LABELRESULTAT( 9, "#ff5050", "bold", File + str("\t◄ Manque _C16 à la fin") )
                                    Actived(ButtonRenameC16) 

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
        # Fonction qui analyse le dossier Relevé de Chambre -----------------------------------------------------------------------------


        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        def CheckBpeShp() :  

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

                        # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx      
                        sh = RecupExcelPatch( directory , File, "C3B" )

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
                        if nombreDeBpeShape != 0:

                            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeBpeShape) + " supports ayant un boitier dans le shape : " + File + "\t" + path.basename(directory)  )

                            for record in DBF(dbfFile):
                                Value = str(record['id_support'])
                                if Value != '' :
                                    inc +=1                                    
                                    LABELRESULTAT( 10, "white", "normal", str( inc ).zfill(3) + " \t " + str( Value ) ) 

                                    # Tous les supports se trouvant dans le Shape
                                    supportListShape.append(Value)            

            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeBPE) + " supports ayant un boitier dans la C3B"  )

            for Element in sorted(C3Blist) :
                if Element != "" :                    
                    firstSepared  = Element.split(' ')[0]
                    SecondSepared = Element.split(' ')[3]
                    # print (  "Element : " + firstSepared + "Element : " + SecondSepared      ) 

                    if len(firstSepared) < 9 :
                        LABELRESULTAT( 10, colorGreen, "bold", "Present dans la C3B : \t\t" + firstSepared + "\t\t" + SecondSepared  ) 
                    else:
                        LABELRESULTAT( 10, colorGreen, "bold", "Present dans la C3B : \t\t" + firstSepared + "\t" + SecondSepared ) 

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

            i,j=0,0
            for Element in supportC3Blist :

                firstSepared  = Element.split(' ')[0]
                NewSupportC3Blist.append(firstSepared)

                if firstSepared in supportListShape :
                    i+=1
                    print ("element trouvée : " + firstSepared)
                    LABELRESULTAT( 10, colorGreen, "bold", str( i ).zfill(3) +" - Present dans bpe.shape : \t\t" + firstSepared )
           
                else :
                    j+=1
                    faute += 1
                    print ("element non trouvée : " + firstSepared)
                    LABELRESULTAT( 10, colorRed, "bold", str( j ).zfill(3) +" - Absent dans bpe.shape : \t\t" + firstSepared ) 

            i=0
            for Element in supportListShape :
                # print ("supportListShape" + str(Element) )
                if Element not in NewSupportC3Blist and  Element != '' :
                    i+=1
                    print ("element trouvée en trop : " + firstSepared)
                    LABELRESULTAT( 10, colorOrang, "bold", str( i ).zfill(3) + " - En trop dans bpe.shape : \t\t" + Element )

            JusteOuFaux(faute) 

        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        

        # Fonction qui analyse et compare le shape support par rapport à la C3B ---------------------------------------------------------
        def CheckSupportShp() :
            global C3Blist     
            global ChambreC3Blist

            Actived(ButtonReleveDeChambre) 
                   

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

                    if path.basename(directory) == "Appui Aérien" :
                        Actived(ButtonAppuiAerien)

                    # Sort les fichies de la liste de fichier
                    for File in sorted(filenames ) :
                        nombreDeSupportShape= 0                      

                        # S'arrete sur le fichier Fxxxxxxxxxxx_C3B.xlsx
                        if "C3B" in File or "C3A" in File:                    

                            # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx
                            try :     
                                sh = RecupExcelPatch( directory , File, "C3A" )
                            except:
                                sh = RecupExcelPatch( directory , File, "C3B" )                                 

                            for rx in range(11,sh.nrows):
                                TypeSupport( 2 , 3 )
                                TypeSupport( 4 , 5 )        

                        # S'arrete sur le fichier support.dbf
                        elif "support.dbf" in File :                        

                            # Affiche le Titre de la Fonction
                            dbfFile  = (directory+"/"+File) 

                            for record in DBF(dbfFile):
                                nombreDeSupportShape +=1                            

                            # Affiche le nombre de cables present dans la C3Bs
                            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeSupportShape) + " supports présent dans le shape" )

                            for record in DBF(dbfFile):
                                increment +=1
                                Value = str(record['id_support'])
                                LABELRESULTAT( 10, "white", "normal", str( increment ).zfill(3) + " \t " + str( Value ) ) 

                                # Tous les supports se trouvant dans le Shape
                                supportListShape.append(Value)                
                        
                # Affiche le Titre de la Fonction
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
                            Actived( ButtonAppuiAerien  )
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
                                LABELRESULTAT( 10, colorOrang, "bold", u"Support : <vide> à supprimer du Shape" )
                            else :
                                LABELRESULTAT( 10, colorOrang, "bold", u"Support : " + element + " à supprimer du Shape" )  
                # ------------------------------------------------------------------------------------------------------------------------------------
                # ------------------------------------------------------------------------------------------------------------------------------------
                InseeListC3B= list(dict.fromkeys(InseeListC3B)) # retire les doublons 

                DictCommune = { "SAINT-BASLEMONT"   : 88411 , "THUILLIERES"      : 88472 , "SAINT-BASLEMONT" : 88411 , 
                                "VALLEROY LE SEC"   : 88490 , "VITTEL"           : 88516 , "SANDAUCOURT"     : 88440 ,
                                "DOMBROT-SUR-VAIR"  : 88141 , "BELMONT-SUR-VAIR" : 88051 , "NORROY"          : 88332 , 
                                "MONTHUREUX-LE-SEC" : 88309 , "HAREVILLE"        : 88231 , "THEY-SOUS-MONTFORT" : 88466,
                                }
               
                val_list = list(DictCommune.values())

                for commune in InseeListC3B :
                    if int(commune) in val_list :
                        #print( u"valeur Trouvée : " + str(commune) )
                        LABELRESULTAT( 10, "pink1", "bold", u"Commune : " + str(commune) + " \tTrouvée")  

                    else :
                        #print( u"valeur Inconnue : " + str(commune) )
                        LABELRESULTAT( 10, "orchid1", "bold", u"Commune : " + str(commune) + " \tInconnue")
                # ------------------------------------------------------------------------------------------------------------------------------------
                # ------------------------------------------------------------------------------------------------------------------------------------
                JusteOuFaux(faute)              
                      
            except Exception as e:
                # print("Il manque le Champ " + str( e )  + " dans le Shape Support ")
                LABELRESULTAT( 10, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )
     
        # Fonction qui analyse et compare le shape support par rapport à la C3B ---------------------------------------------------------

        # Fonction qui analyse et compare le shape cable par rapport à la C3B -----------------------------------------------------------
        def CheckCableShp(): 

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

                                # Monte le fichier Fxxxxxxxxxxx_C3B.xlsx
                                sh = RecupExcelPatch( directory , File , "C3B")  

                                # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                                depart = 11    # Debute appartir de la ligne 11                    
                                for rx in range( depart , sh.nrows):

                                    # S'arrete sur la Colonne P : Câble
                                    CABLE   = str( sh.cell(rowx=rx, colx=15).value)

                                    # Continue le Programme Uniquement si le Câble est Posé
                                    if CABLE != u"Câble non posé" :                                        

                                        SUP1ORT = ReconnaissanceSupport(3)
                                        SUP2ORT = ReconnaissanceSupport(5)                                        

                                        # S'arrete sur la Colonne K : Diamètre du Câble à Poser
                                        DIAM    = str( sh.cell(rowx=rx, colx=10).value)

                                        # Réunie Chaques Valeurs
                                        Value = SUP1ORT +"-"+ SUP2ORT +"-"+ DIAM
                                        C3Blist.append( Value )  

                                        # Nombre de Câble                                        
                                        nombreDeCable += 1  
                        
                            # S'arrete sur le shape cable.dbf
                            elif "cable.dbf" in File :

                                inc += 1                                
                                dbfFile  = (directory+"/"+File) # Créer le Chemin vers le Shape

                                # Boucle sur chaque ligne du fichier
                                for record in DBF(dbfFile):

                                    IDAValue, IDiamValue = RecupereSupport('id_a' , 'type_a' )
                                    IDBValue, IDiamValue = RecupereSupport('id_b' , 'type_b' )

                                    # Réunie Chaques Valeurs
                                    Value = IDAValue + "-" + IDBValue + "-" + IDiamValue
                                    DBFlist.append( Value )

                        except Exception as e:
                            LABELRESULTAT( 10, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )   
           
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
         
            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeCable) + " câbles posées dans la C3B" )            

            # Entete
            LABELRESULTAT( 10, "white", "bold", u"\t\t\t id_a \t\t\t   id_b\t\tdiam_cbl" ) 
            
            i=0 
            for Element in sorted(DBFlist) :
                firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                Count = (DBFlist.count(Element) )                
                if Count != 1 :
                    i+=1
                    LABELRESULTAT( 10, "pink", "bold", str(i).zfill(3) + u" - Vérifier le Doublon : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
            
            # Boucle sur la liste Résultat des cables à Garder
            for Element in sorted(DBFlist) :                
                                
                # Compare les Elements avec la liste DBFlistsansID
                if Element not in C3Blist :
                    firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                    if len( firstSepared ) < 9 :                        
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorOrang, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorOrang, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )

                    else :
                        if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                            LABELRESULTAT( 10, colorOrang, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t\t" +  str(ThirtSepared) )
                        else :
                            LABELRESULTAT( 10, colorOrang, "bold", u"Câble à supprimer du Shape : \t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )
                
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
        # Fonction qui analyse et compare le shape cable par rapport à la C3B ---------------------------------------------------------------------


        # Renomme les fichiers 88516/1431.xlsx en 88516/1431_C16.xlsx -----------------------------------------------------------------------------  
        def CheckRenameC16():

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
                Disabled( ButtonRenameC16 )
            else :
                LABELRESULTAT( 10, "white", "normal", str(RenameCount) + " Fichiers ont été Renommés") 

        # Creation des Boutons --------------------------------------------------------------------------------------------------------------------              
        def Picture (Path) :            
            Size     = int(self.screenHeight*0.12)          # Definit la taille du logo par rapport a l'ecran            
            dirname  = path.dirname(path.abspath(__file__)) # definit le chemin jusqu au program                 
            picture  = path.join(dirname, Path)             # Definit le Chemin du PNG 
            original = Image.open(picture)                  # Ouvre le PNG               
            resized  = original.resize((Size, Size))        # le redimensionne    
            self.img = ImageTk.PhotoImage(resized)          # import l'image redimensionner             
            return self.img                                 # Retourne l'image          

        # Creation du container 
        container = Frame( self.centralFrame )          

        # Creation de l'espace dessin
        canvas = Canvas(container, bg= self.LeaveColor , highlightthickness=0, bd=1, relief=SUNKEN )

        # Définition des la Police d'écriture  
        helvetica = tkfont.Font(family='Arcade', size=17, weight='bold')

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

        def CreationDeBouton(): 

            FrameShape   = Frame( self.left_frame, bg=self.LeaveColor )
            FrameShape.place(   relx=0, rely=1, anchor=SW, relwidth=1, relheight=0.75 )

            listButtonTitre  = [ u"C3B C3A"     , u"Support\n.shp"   , u"Cable\n.shp"   , u"Bpe.shp"   , u"Releve de\nChambre"   , u"Rename\nC16"  , u"Appui\nAerien" ]
            listCommandShape = [ CheckC3bXlsx        , CheckSupportShp     , CheckCableShp     , CheckBpeShp     , CheckReleveChambre       , CheckRenameC16     , CheckAppuiAerien     ]
            listButtonName   = [ u'ButtonAnalyseC3B' , u'ButtonSupportShp' , u'ButtonCableShp' , u'ButtonBpeShp' , u'ButtonReleveDeChambre' , u'ButtonRenameC16' , u'ButtonAppuiAerien' ]

            def LabelBoutonDisabled(photo, Yposition):                    
                Label( FrameShape, image=photo, bg=self.LeaveColor ).place(relx = 0.05, rely = Yposition, anchor = W)

            icons = ["./logo/Button4K-MenuC3B.png","./logo/Button4K-MenuSupport.png","./logo/Button4K-MenuCable.png"
                        ,"./logo/Button4K-MenuBoitier.png","./logo/Button4K-MenuReleve.png"
                        ,"./logo/Button4K-MenuRename.png","./logo/Button4K-MenuSupport.png" ] 

            decalage = .05

            image      = Image.open( icons[0] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconC3b = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconC3b, 0/7+decalage ) 

            image      = Image.open( icons[1] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconSupport  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconSupport, 1/7+decalage) 

            image      = Image.open( icons[2] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconCable  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconCable, 2/7+decalage ) 

            image      = Image.open( icons[3] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconBoitier  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconBoitier, 3/7+decalage ) 

            image      = Image.open( icons[4] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconReleve  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconReleve, 4/7+decalage ) 

            image      = Image.open( icons[5] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconRename  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconRename, 5/7+decalage ) 

            image      = Image.open( icons[6] )
            image      = image.resize((60,60), Image.ANTIALIAS)
            self.iconAppui  = ImageTk.PhotoImage(image)
            LabelBoutonDisabled( self.iconAppui, 6/7+decalage)            

            for i in range(0 ,7) :    
                listButtonName[i] = Button(FrameShape, bg=self.LeaveColor , highlightthickness=0, cursor="hand2", fg="cyan"
                    , relief=FLAT, activebackground=self.LeaveColor, command=listCommandShape[i], state=DISABLED, font=helvetica,text=listButtonTitre[i] )
                listButtonName[i].place(relx = 16/40, rely = i/7+.05, anchor = W, relwidth=0.6, relheight =0.10)  

            return listButtonName  
    
        ButtonAnalyseC3B, ButtonSupportShp, ButtonCableShp, ButtonBpeShp, ButtonReleveDeChambre, ButtonRenameC16 ,ButtonAppuiAerien = CreationDeBouton( ) 
         
        #creation du Bouton Monter Fichier
        Button(self.left_frame, image=Picture("logo/ButtonMonter.png"), border=0, bg=self.LeaveColor, highlightthickness=0
            , bd=1, relief=FLAT, activebackground=self.LeaveColor, command= FuncUpFile, cursor="hand2" 
                ).place(  relx=0, rely=9/40, anchor=SW, relwidth=1, relheight=0.2)