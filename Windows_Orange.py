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
import time
import progressbar 

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
        colorGray  = "gray63"

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

        def TitreFonction(texte):
            LABELRESULTAT( 9, "white", "normal", u"\n" )
            LABELBORDEREAU( scrollable_frame , colorBlue , texte )  

        def FuncUpFile():  
            
            try :
                global Racine                                      
                Racine = filedialog.askdirectory(initialdir=r"C:\Users\Arnaud_2018\Desktop\DOE-Orage",title='Choisissez un repertoire')
            except:            
                Path = path.dirname(path.abspath(__file__))                                       
                Racine = filedialog.askdirectory(initialdir=Path,title='Choisissez un repertoire')

            if Racine != "" :

                Label( scrollable_frame, bg=colorOrang, text="",width=int(self.screenWidth*0.054 ) ).pack( )
                
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

            TitreFonction( "Analyse de la Command dacces"  ) 

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

                            listSupportA = [ "C","CT","A","AT","P" ]

                            listC  = ["C","IMB","F","P","PT","A","AT","" ]
                            listCT = ["P","A" ]
                            lsitA  = ["IMB","A", "AT","F","P","PT" ]
                            listAT = ["P","A" ]
                            listP  = ["F","P","PT" ]

                            if len(NumA) < 9 :
                                Value = "►\t" + "Ligne : " + str(rx+1).zfill(3) + "\t\t" + "Alveole: "+Alveol + "\t" + "Type A = "+TypeA + "\t" + NumA + "\t\t" + "Type B = "+ TypeB  + "\t" + NumB
                            else :
                                Value = "►\t" + "Ligne : " + str(rx+1).zfill(3) + "\t\t" + "Alveole: "+Alveol + "\t" + "Type A = "+TypeA + "\t" + NumA + "\t" + "Type B = "+ TypeB  + "\t" + NumB
              
                            if TypeA == "C"  :
                                if str(Alveol) not in AlveolAccept :
                                    val = str( "►\t" + "Ligne : " + str(rx).zfill(3) + "\t\tAlveol Fausse: " + str( Alveol ) )
                                    LABELRESULTAT( 9, colorRed , "normal", val )

                                if TypeB not in listC :
                                    LABELRESULTAT( 9, colorRed , "normal", Value )
                                else :
                                    LABELRESULTAT( 9, colorGreen , "normal",  Value )

                            elif TypeA == "CT"  :
                                if TypeB not in  listCT :
                                    LABELRESULTAT( 9, colorRed , "normal",  Value )
                                else :
                                    LABELRESULTAT( 9, colorGreen , "normal", Value )

                            elif TypeA == "A"  :
                                if TypeB not in  lsitA :
                                    LABELRESULTAT( 9, colorRed , "normal",  Value )
                                else :
                                    LABELRESULTAT( 9, colorGreen , "normal", Value )

                            elif TypeA == "AT"  :
                                if TypeB not in  listAT :
                                    LABELRESULTAT( 9, colorRed , "normal",  Value )
                                else :
                                    LABELRESULTAT( 9, colorGreen , "normal", Value )

                            elif TypeA == "P"  :
                                if TypeB not in  listP :
                                    LABELRESULTAT( 9, colorRed , "normal",  Value )
                                else :
                                    LABELRESULTAT( 9, colorGreen , "normal", Value )
                            
                            else :
                                LABELRESULTAT( 9, colorOrang , "normal",  Value )
       
        # Fonction qui analyse le dossier Appui Aérien ----------------------------------------------------------------------------------------
        def CheckAppuiAerien():

            TitreFonction( "Analyse du dossier Appui Aérien"  )  

            global C3Blist 
            C6List , PoteauxFauxList = [] , []
            faute, increment ,   inc = 0 , 0 , 0

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
            LABELBORDEREAU( scrollable_frame, "black", "Il y a "+ str(len(C6List)) + " Poteaux dans la C6 " )

            listGespot = []  
            inc+=1

            for Element in sorted(C3Blist) : 

                firstSepared  = Element.split(' ')[0]
                SecondSepared = Element.split(' ')[2]

                if firstSepared == "Poteau" :
                    resultat = SecondSepared.split("/")

                    if resultat[1] not in C6List : 
                        LABELRESULTAT( 9, colorRed , "normal", u"►\t" + str(firstSepared) + " " + str(SecondSepared) + u"\test présent dans la C3B mais absent de la C6" )
                    else :
                        LABELRESULTAT( 9, colorBlue , "normal", u"►\t" + str(firstSepared) + " " + str(SecondSepared) + u"\t Ok" )

                elif "GESPOT" in File :         
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
            mylistFOA , addNexList , addC3BChambreList , NewElementFOA = [] , [] , [] , []

            # ----------------------------------------------------------------------------------------------------
            def ResultatValidation( color , varText , texte ) :

                mylistFOA.append( File[:len(File)-9] ) 

                if File[len(File)-9:-5] == "_C16" :
                    if len(File) < 18 :
                        LABELRESULTAT( 9, color, varText, File + texte ) 
                    else:
                        LABELRESULTAT( 9, color, varText, File + texte ) 
                else:
                    LABELRESULTAT( 9, colorRed, varText , File + str("\t◄ Manque _C16 à la fin") )
                    Actived(ButtonRenameC16)
            # ----------------------------------------------------------------------------------------------------
           
            for element in  C3Blist:
                addNexList.append( element.split(" ")[2] ) # Renvoie > 88516/152

            # retrouve tous les chemin dossier et fichiers 
            for directory, dirnames, filenames  in walk(Racine):

                # S'arrete sur le dossier "Relevé de chambre"
                if path.basename(directory) == u"Relevé de chambre" :

                    TitreFonction( u"Analyse du dossier Relevé de Chambres"   )  

                    LABELBORDEREAU( scrollable_frame , "black" , " ( il y a "+ str(len(filenames )) + " fichiers FOA )"  )

                    # Boucle sur la liste du DBF pour sortir chaque Element #
                    for File in sorted(filenames ) :

                        Newfilenames  = File[:len(File)-9].replace("-", "/")

                        if File[len(File)-4:] == ".xls":            
                            if File[len(File)-8:-4] == "_C16" :
                                LABELRESULTAT( 9, "white", "normal", File ) 
                            else:
                                LABELRESULTAT( 9, colorRed, "bold", File + str(" ◄ Manque _C16 à la fin") )
                            mylistFOA.append( File[:len(File)-8] )

                        elif File[len(File)-4:] == "xlsx":
                            if Newfilenames in addNexList :
                                ResultatValidation( colorGreen , "bold" , str("\t\t◄ Nécessaire donc OK") )
                            else :
                                ResultatValidation( colorGray , "normal" ,  str("\t\t◄ Pas Nécessaire à Retirer du dossier \"Relevé de Chambre\"") ) 

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

            dirname = path.dirname( path.abspath(__file__))
            file    = path.join( dirname, 'RapportFoaManquant.txt' )
            my_file = open( file ,'w+' )

            for element in sorted(addC3BChambreList) :
                if element not in NewElementFOA :
                    # print ( "manquant : " + str(element) )
                    LABELRESULTAT( 9, colorRed, "bold", "Le fichier FOA ( " + element + str(" )\t est manquant dans le dossier \"Relevé de Chambre\"") )                    
                    my_file.write( element + "\r")

            my_file.close()  
        # Fonction qui analyse le dossier Relevé de Chambre -----------------------------------------------------------------------------


        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        def CheckBpeShp() :  

            # Initialisation des Listes
            C3Blist , supportListShape , supportListShape , NewSupportC3Blist , supportC3Blist = [] , [] , [] , [] , []        
            nombreDeBpeShape = 0 

            def supportAvecBoitier( interger, Type ):

                if Boitier[0] == Type :

                    Cable   = str( sh.cell(rowx=rx, colx=15).value)
                    support = str( sh.cell(rowx=rx, colx=interger).value)                    

                    if Cable != u"Câble non posé" :
                        supportC3Blist.append(  support + " \t " + Boitier )
                        C3Blist.append( support + " \t " + Boitier )

                return C3Blist , supportC3Blist

            TitreFonction( u"Analyse du nombres de boitiers"   ) 
                      
            for directory, dirnames, filenames  in walk(Racine):          
                for File in sorted(filenames ) : 
                    if "C3B" in File :
                        sh = RecupExcelPatch( directory , File, "C3B" )

                        # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                        depart = 11    # Debute appartir de la ligne 11 
                        for rx in range(depart,sh.nrows):                           
                            
                            Boitier = str( sh.cell(rowx=rx, colx=14).value)                       
                            
                            if Boitier != "" :

                                C3Blist, supportC3Blist = supportAvecBoitier( 3, "A" )
                                C3Blist, supportC3Blist = supportAvecBoitier( 5, "B" )

                    # S'arrete sur le fichier support.dbf
                    elif "bpe.dbf" in File :
                        nombreDeBpeShape = 0
                        dbfFile  = (directory+"/"+File) 
                        for record in DBF(dbfFile):
                            nombreDeBpeShape +=1                            

                        # Affiche le nombre de cables present dans la C3Bs
                        if nombreDeBpeShape != 0:
                            for record in DBF(dbfFile):
                                Value = str(record['id_support'])
                                if Value != '' :
                                    # Tous les supports se trouvant dans le Shape
                                    supportListShape.append(Value)            
            nombreDeBPE = 0
            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeBPE) + " supports ayant un boitier dans la C3B"  )
            
            def verifPresent( listA , listB , information , firstValue , SecondValue ):

                def RetourC3B( SecondNbr, LongText, texte):
                    firstSepared  = Element.split(' ')[0]
                    SecondSepared = Element.split(' ')[SecondNbr]
                    NewSupportC3Blist.append(firstSepared)

                    if len(firstSepared) < LongText :
                        LABELRESULTAT( 9, colorGreen, "bold", str( i ).zfill(3) + texte +  "\t\t" + firstSepared + "\t\t" + SecondSepared  ) 
                    else:
                        LABELRESULTAT( 9, colorGreen, "bold", str( i ).zfill(3) + texte +  "\t\t" + firstSepared + "\t" + SecondSepared )

                i=1
                for Element in sorted(listA) :                    
                    if Element not in listB :
                        RetourC3B( firstValue, SecondValue , information )
                        i+=1

            emptyList = [""]
            verifPresent( C3Blist , emptyList , " - Present dans la C3B :" , 3 , 9 )
            verifPresent( supportC3Blist , C3Blist , " - Absent de la C3B :" , 1 , 8 )

            LABELRESULTAT( 9, colorGreen, "bold", "" ) 
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(len(supportListShape) ) + " supports ayant un boitier dans le Shape"  )

            i,j=1,1
            for Element in sorted(supportC3Blist) :
                firstSepared  = Element.split(' ')[0]
                NewSupportC3Blist.append(firstSepared)
                if firstSepared in supportListShape :                    
                    LABELRESULTAT( 9, colorGreen, "bold", str( i ).zfill(3) + " - Present dans bpe.shape : \t" + firstSepared ) 
                    i+=1          
                else :                    
                    LABELRESULTAT( 9, colorRed, "bold", str( j ).zfill(3) +" - Absent dans bpe.shape : \t" + firstSepared )
                    j+=1 

            i=1
            for Element in sorted(supportListShape) :
                if Element not in NewSupportC3Blist and  Element != '' :
                    LABELRESULTAT( 9, colorOrang, "bold", str( i ).zfill(3) + " - En trop dans bpe.shape : \t" + Element )
                    i+=1
            JusteOuFaux(j-1) 

        # Fonction qui analyse et compare le shape BPE par rapport à la C3B -------------------------------------------------------------
        

        # Fonction qui analyse et compare le shape support par rapport à la C3B ---------------------------------------------------------
        def CheckSupportShp() :
            global C3Blist     
            global ChambreC3Blist

            Actived(ButtonReleveDeChambre)                    

            def TypeSupport( ColonneType , ColonneSupport ):

                Type  = str( sh.cell( rowx=rx, colx=ColonneType ).value)  
                Cable = str( sh.cell(rowx=rx, colx=15).value)                        
                if Cable != u"Câble non posé" and Type != "F" and Type != "AT" : 

                    Support = str( sh.cell(rowx=rx, colx=ColonneSupport).value)

                    if Type == "C" :
                        Type = "Chambre"
                        ChambreC3Blist.append( Type + " - " + Support )
                    if Type == "A" :
                        Type = "Poteau"    
                    
                    C3Blist.append( Type + " - " + Support )

                return ChambreC3Blist

            # Initialisation des listes
            supportListC3B,supportListShape,C3Blist,ChambreC3Blist,InseeListC3B = [] , [] , [], [] ,[]  
            nombreDeSupport , increment , nbrDeSupportDsShape , faute = 0 , 1 , 0 , 0

            try :

                TitreFonction( u"Analyse des Supports" ) 

                # Retrouve tous les chemins , les dossiers et les fichiers de Racine
                for directory, dirnames, filenames  in walk(Racine):

                    if path.basename(directory) == "Appui Aérien" :
                        Actived(ButtonAppuiAerien)

                    # Sort les fichies de la liste de fichier
                    for File in sorted(filenames ) :
                        nbrDeSupportDsShape= 0                      

                        # S'arrete sur le fichier Fxxxxxxxxxxx_C3B.xlsx
                        if "C3B" in File:   

                            sh = RecupExcelPatch( directory , File, "C3B" )                                 

                            for rx in range(11,sh.nrows):
                                TypeSupport( 2 , 3 )
                                TypeSupport( 4 , 5 )        

                        # S'arrete sur le fichier support.dbf
                        elif "support.dbf" in File :                        

                            dbfFile  = (directory+"/"+File) 

                            for record in DBF(dbfFile):
                                nbrDeSupportDsShape +=1                            

                            # Affiche le nombre de cables present dans la C3Bs
                            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nbrDeSupportDsShape) + " supports présent dans le shape" )

                            for record in DBF(dbfFile):
                                
                                Value = str(record['id_support'])
                                LABELRESULTAT( 9, "white", "normal", str( increment ).zfill(3) + " \t " + str( Value ) ) 

                                # Tous les supports se trouvant dans le Shape
                                supportListShape.append(Value)    

                                increment +=1            
                        
                # Affiche le Titre de la Fonction
                C3Blist= list(dict.fromkeys(C3Blist)) # retire les doublons 
                for element in sorted(C3Blist) :
                    if element != "" :
                        nombreDeSupport += 1

                # Affiche le nombre de calbes present dans la C3Bs
                LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeSupport) + " supports présents dans la C3B" )
                i = 1
                dirname = path.dirname( path.abspath(__file__))
                file    = path.join( dirname, 'RapportFoaManquant.txt' )
                print ( file )
                my_file = open( file ,'w+' )
                for element in sorted(C3Blist) :

                    Newelement = element.split(" ")[2]
                    INSEE      = Newelement.split("/")[0]

                    InseeListC3B  .append(INSEE)
                    supportListC3B.append(Newelement)

                    if element != "" :
                        
                        if "Chambre" in element :
                            LABELRESULTAT( 9, colorGreen, "normal", str(i).zfill(3) + " - " + element )
                            Newfilenames  = Newelement.replace("/", "-")
                            my_file.write( Newfilenames + "\r")
                            i+=1
                        else:
                            Actived( ButtonAppuiAerien  )
                            LABELRESULTAT( 9, colorBlue , "normal", str(i).zfill(3) + " - " + element )
                            i+=1

                my_file.close() 

                if len(supportListShape) == 0 :
                    Titre = LABELBORDEREAU( scrollable_frame , colorRed , "Le fichier support.shp est manquant, vide ou zippé " )
                    Titre['fg'] = "black"
                    LABELRESULTAT( 9, colorGreen, "normal", "" )
                    faute += 1

                else :                    
                    def BoucleSurListe( color, liste, inverseListe, number, elementVar , information ) :
                        for element in sorted(liste) :
                            if element not in inverseListe :
                                if len(element) < number :
                                    LABELRESULTAT( 9, color, "bold", "Support : " + elementVar + information ) 
                                else :
                                    LABELRESULTAT( 9, color, "bold", "Support : " + element + information ) 

                    BoucleSurListe( colorRed    , supportListC3B    , supportListShape  , 8 , element , u"\test présent dans la C3B mais absent de Support.shp"  )
                    BoucleSurListe( colorOrang  , supportListShape  , supportListC3B    , 1 , "vide" , "\tà supprimer du Shape" )
                    faute += 1

                InseeListC3B= list(dict.fromkeys(InseeListC3B)) # retire les doublons 

                DictCommune = { "SAINT-BASLEMONT"   : 88411 , "THUILLIERES"      : 88472 , "SAINT-BASLEMONT"    : 88411 , 
                                "VALLEROY LE SEC"   : 88490 , "VITTEL"           : 88516 , "SANDAUCOURT"        : 88440 ,
                                "DOMBROT-SUR-VAIR"  : 88141 , "BELMONT-SUR-VAIR" : 88051 , "NORROY"             : 88332 , 
                                "MONTHUREUX-LE-SEC" : 88309 , "HAREVILLE"        : 88231 , "THEY-SOUS-MONTFORT" : 88466 
                                }
               
                val_list = list(DictCommune.values())

                for commune in InseeListC3B :
                    if int(commune) in val_list :
                        LABELRESULTAT( 9, "pink1", "bold", u"Commune : " + str(commune) + " \tTrouvée") 
                    else :
                        LABELRESULTAT( 9, "orchid1", "bold", u"Commune : " + str(commune) + " \tInconnue")

                JusteOuFaux(faute)              
                      
            except Exception as e:
                # print("Il manque le Champ " + str( e )  + " dans le Shape Support ")
                LABELRESULTAT( 9, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )
     
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

                # Si la Case est égale à AT ou F remplacer la valeur par "SupportTiers"
                if typeValue == "AT" or typeValue == "F" :
                    idValue = "SupportTiers"

                if typeValue == "IMB" :
                    idValue = "Immeuble"

                idDiam = str(record['diam_cbl'])                                 

                return idValue, idDiam

            # Initialisation des Listes
            C3Blist ,DBFlist = [] , []
            nombreDeCable , inc = 0 , 0

            TitreFonction( u"Analyse des Câbles"  ) 
           
            for directory, dirnames, filenames  in walk(Racine):
                for File in sorted(filenames ) :
                    try :
                        if "C3B" in File :
                            sh = RecupExcelPatch( directory , File , "C3B")  

                            # Boucle sur toutes les lignes du fichier pour récupérer les informations                        
                            depart = 11                  
                            for rx in range( depart , sh.nrows):

                                # S'arrete sur la Colonne P : Câble
                                CABLE   = str( sh.cell(rowx=rx, colx=15).value)

                                # Continue le Programme Uniquement si le Câble est Posé
                                if CABLE != u"Câble non posé" :                                        

                                    SUP1ORT = ReconnaissanceSupport(3)
                                    SUP2ORT = ReconnaissanceSupport(5)                                        

                                    # S'arrete sur la Colonne K : Diamètre du Câble à Poser
                                    DiametreCable    = str( sh.cell(rowx=rx, colx=10).value)

                                    # Réunie Chaques Valeurs
                                    Value = SUP1ORT +"-"+ SUP2ORT +"-"+ DiametreCable
                                    C3Blist.append( Value )  
                                     
                                    nombreDeCable += 1  
                    
                        # S'arrete sur le shape cable.dbf
                        elif "cable.dbf" in File :
                            inc += 1                                
                            dbfFile  = (directory+"/"+File) # Créer le Chemin vers le Shape

                            # Boucle sur chaque ligne du fichier
                            for record in DBF(dbfFile) :                                
                                IDAValue, IDiamValue = RecupereSupport('id_a' , 'type_a' )
                                IDBValue, IDiamValue = RecupereSupport('id_b' , 'type_b' )
                                Value = IDAValue + "-" + IDBValue + "-" + IDiamValue
                                DBFlist.append( Value )

                    except Exception as e:
                        LABELRESULTAT( 10, colorRed, "bold", " \n\n" + str( e )  + " \nfermé le fichier Excel avant analyse" )   
        
            def TextJustify( color, texte, tab1, tab2, tab3, tab4, tab5 , tab6, tab7, tab8, tab9, tab10, tab11 ):
                firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                if len( firstSepared ) < 9 :
                    if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                        LABELRESULTAT( 9, color, "bold", texte + tab1 + str(firstSepared) + tab2 + str(SecondSepared) + tab3 +  str(ThirtSepared) )
                    else :
                        LABELRESULTAT( 9, color, "bold", texte + tab4 + str(firstSepared) + tab5 + str(SecondSepared) + tab6 +  str(ThirtSepared) )
                else :
                    if len( SecondSepared ) < 9 and SecondSepared != "Immeuble" :
                        LABELRESULTAT( 9, color, "bold", texte + tab7 + str(firstSepared) + tab8 + str(SecondSepared) + tab9 +  str(ThirtSepared) )
                    else :
                        LABELRESULTAT( 9, color, "bold", texte + tab10 + str(firstSepared) + tab11 + str(SecondSepared) + "\t" +  str(ThirtSepared) )
            
            # Affiche le nombre de calbes present dans la C3B
            LABELBORDEREAU( scrollable_frame , "black" , "Il y a " + str(nombreDeCable) + " câbles posées dans la C3B" )            
            LABELRESULTAT( 10, "white", "bold", u"\t\t\t\t id_a \t\t   id_b\t\tdiam_cbl" ) 
            
            i=1 
            for Element in sorted(DBFlist) :
                firstSepared, SecondSepared, ThirtSepared = ElementSeparee(Element)
                Count = (DBFlist.count(Element) )                
                if Count != 1 :                    
                    LABELRESULTAT( 10, "pink", "bold", str(i).zfill(3) + u" - Vérifier le Doublon : \t\t" + str(firstSepared) + "\t" + str(SecondSepared) + "\t" +  str(ThirtSepared) )
                    i+=1

            # Début de l'analyse du Shape
            for Element in sorted(DBFlist) :  
                if Element not in C3Blist :
                    TextJustify(colorOrang , u"Câble à supprimer du Shape : ", 
                    "\t" , "\t\t", "\t" , "\t" , "\t\t", "\t\t", "\t" , "\t", "\t\t" , "\t" , "\t"  ) 
                else : 
                    TextJustify(colorGreen , u"\tCâble à Garder : ",
                    "\t\t" , "\t\t", "\t\t" , "\t\t" , "\t\t", "\t", "\t\t" , "\t", "\t\t" , "\t\t" , "\t" ) 


            for Element in sorted(C3Blist) :
                if Element not in DBFlist : 
                    TextJustify(colorRed , u"Câble à ajouter sur le Shape : ", 
                    "\t" , "\t\t", "\t" , "\t" , "\t\t", "\t\t", "\t" , "\t", "\t\t" , "\t" , "\t" )
                else:
                    pass

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
                LABELRESULTAT( 9, "white", "normal", "Tous les Fichiers de Relevé de chambre sont bien nommés") 
                Disabled( ButtonRenameC16 )
            else :
                LABELRESULTAT( 9, "white", "normal", str(RenameCount) + " Fichiers ont été Renommés") 

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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120) ), "units")          

        # Creation de l'espace dessin
        canvas = Canvas(container, bg= self.LeaveColor , highlightthickness=0, bd=1, relief=SUNKEN )
        canvas.bind_all("<MouseWheel>", _on_mousewheel)        

        # Définition des la Police d'écriture  
        helvetica = tkfont.Font(family='Verdana', size=10, weight='bold')

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

        """        
        widgets = ['Loading: ', progressbar.AnimatedMarker()] 
        bar = progressbar.ProgressBar(widgets=widgets).start() 
        
        for i in range(50): 
            time.sleep(0.1) 
            bar.update(i) 
        """


        def CreationDeBouton(): 

            FrameShape   = Frame( self.left_frame, bg=self.LeaveColor )
            FrameShape.place(   relx=0, rely=1, anchor=SW, relwidth=1, relheight=0.75 )

            listButtonTitre  = [ u"\tC3B C3A"     , u"\tSupport.shp"   , u"\tCable.shp"   , u"\tBpe.shp"   , u"\tReleve de\n\tChambre"   , u"\tRename C16"  , u"\tAppui Aerien" ]
            listCommandShape = [ CheckC3bXlsx        , CheckSupportShp     , CheckCableShp     , CheckBpeShp     , CheckReleveChambre       , CheckRenameC16     , CheckAppuiAerien     ]
            listButtonName   = [ u'ButtonAnalyseC3B' , u'ButtonSupportShp' , u'ButtonCableShp' , u'ButtonBpeShp' , u'ButtonReleveDeChambre' , u'ButtonRenameC16' , u'ButtonAppuiAerien' ]

            icons = ["./logo/Button4K-MenuC3B.png","./logo/Button4K-MenuSupport.png","./logo/Button4K-MenuCable.png"
                        ,"./logo/Button4K-MenuBoitier.png","./logo/Button4K-MenuReleve.png"
                        ,"./logo/Button4K-MenuRename.png","./logo/Button4K-MenuSupport.png" ] 

            decalage   = .05
            self.icons = []

            i=0
            for pathIcon in icons :

                image = Image.open( pathIcon )
                image = image.resize((45,45), Image.ANTIALIAS)
                icon  = ImageTk.PhotoImage( image )    

                listButtonName[i] = Button(FrameShape,bd=1, bg=self.LeaveColor , highlightthickness=4, cursor="hand2", fg="cyan"
                    , relief=SOLID , activebackground="cyan" , command=listCommandShape[i], state=DISABLED
                    , font=helvetica,text=listButtonTitre[i] , highlightcolor="pink", highlightbackground="#37d3ff" )
                listButtonName[i].place(relx = 0, rely = i/11+decalage, anchor = W, relwidth= 1, relheight = 0.10 ) 

                Label( FrameShape, image=icon, bg=self.LeaveColor ).place(relx = 0.05, rely = i/11+decalage, anchor = W)
                self.icons.append( icon )  

                i+=1 

            return listButtonName  
    
        ButtonAnalyseC3B, ButtonSupportShp, ButtonCableShp, ButtonBpeShp, ButtonReleveDeChambre, ButtonRenameC16 ,ButtonAppuiAerien = CreationDeBouton( ) 
         
        #creation du Bouton Monter Fichier
        Button(self.left_frame, image=Picture("logo/ButtonMonter.png"), border=0, bg=self.LeaveColor, highlightthickness=0
            , bd=1, relief=FLAT, activebackground=self.LeaveColor, command= FuncUpFile, cursor="hand2" 
                ).place(  relx=0, rely=9/40, anchor=SW, relwidth=1, relheight=0.2)