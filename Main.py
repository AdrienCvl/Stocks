# coding: utf-8

import os
from os import system, name
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

Terminal_Name = "ORTHO 38"
Nom = input("Nom du fichier (format : 210618.xlsx) : ")

Valeurs = np.zeros((53, 4))
Produits = np.array((["Attelle de poignet taille 1"], ["Attelle de poignet taille 2"], ["Attelle de pouce T1"], ["Attelle de pouce T2"], ["Attelle de pgt/pce gauche T0"], ["Attelle de pgt/pce gauche T1"], ["Attelle de pgt/pce gauche T2"], ["Attelle de pgt/pce droit T0"], ["Attelle de pgt/pce droit T1"], ["Attelle de pgt/pce droit T2"], ["Attelle metacarpiens gauche T1"], ["Attelle metacarpiens gauche T2"], ["Attelle metacarpiens  droit T1"], ["Attelle metacarpiens droit T2"], ["Gilet d'épaule"], ["Attelle cheville enfant"], ["Attelle cheville adulte"], ["Attelle de cheville A2T T1"], ["Attelle de cheville A2T T2"], ["Attelle de cheville A2T T3"], ["Chaussure marche basse T1"], ["Chaussure marche basse T2"], ["Chaussure marche basse T3"], ["Chaussure de décharge XS"], ["Chaussure de décharge S"], ["Chaussure de décharge M"], ["Chaussure de décharge L"], ["Chaussure de décharge XL"], ["Collier cervical petit T1"], ["Collier cervical petit T2"], ["Collier cervical petit T3"], ["Collier cervical petit T4"], ["Collier cervical grand T1"], ["Collier cervical grand T2"], ["Collier cervical grand T3"], ["Collier cervical grand T4"], ["Ceinture de soutien lombaire T1"], ["Ceinture de soutien lombaire T2"], ["Ceinture de soutien lombaire T3"], ["Attelle de genou haut 40cm"], ["Attelle de genou haut 60cmT1"], ["Attelle de genou haut 60cm T2"], ["Attelle de genou haut 50cm T1"], ["Attelle de genou haut 50cm T2"], ["Attelle de genou haut 70cm"], ["Botte de marche haute S T1"], ["Botte de marche haute M T2"], ["Botte de marche haute L T3"], ["Botte de marche basse S T1"], ["Botte de marche basse M T2"], ["Botte de marche basse L T3"], ["Bequille enfant"], ["Bequille adulte"]))

df_name = pd.DataFrame(np.array(Produits), columns=["Produit"])

df2 = pd.DataFrame(np.array(Valeurs), columns=["Stock", 'Sorties', 'Verification 1', 'Verification 2'])

def Main_Menu():
    clear()
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:
        print ("Bienvenue chez ORTHO 38,\nQue voulez-vous faire ?\n-----------------------\n1 : Relever les stocks\n2 : relever les sorties\n3 : Vérrifier les sorties\n--------------\nX : Quitter\n  ")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "1" :
            clear()
            Stock_Menu()
            break

        elif User_Command == "2" :
            clear()
            Sortie_Menu()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                clear()
                Terminal_Launched = False
            if Choix == "B":
                clear()
                Main_Menu()
                break

        else :
            print ("Entrez une des valeur correspondante !")

def Stock_Menu():
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:

        print ("RELEVE STOCKS\n-----------------------\n1 : Entrer Stock du jour\n-----------------------\nB : Retour\nX : Quitter\n  ")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "1" :
            Stock()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                Terminal_Launched = False
                clear()
            elif Choix == "B":
                clear()
                Stock_Menu()
                break

        else :
            print ("Entrez une des valeur correspondante !")

def Stock():

    print ("")
    print ("")
    print ("Entrée du stock (chiffres/nombres) :")
    print ("_________________________________________")
    print ("")
    print ("")

    o = input("Gilet d'épaules : ")
    try:
        o = int(o)
    except ValueError:
        print("Vous devez entrer un nombre")
        Lettre = True
        while Lettre:
            o = input("Gilet d'épaules : ")
            try:
                o = int(o)
            except ValueError:
                print("Vous devez entrer un nombre")
                Lettre = True
            else:
                Lettre = False 
    df2.at[14,"Stock"] = o

    clear()

    x = input("Chaussure de décharge XS :")
    try:
        x = int(x)
    except ValueError:
        print("Vous devez entrer un nombre")
        Lettre = True
        while Lettre:
            x = input("Chaussure de décharge XS :")
            try:
                x = int(x)
            except ValueError:
                print("Vous devez entrer un nombre")
                Lettre = True
            else:
                Lettre = False 
    df2.at[23,"Stock"] = x

    y = input("Chaussure de décharge S :")
    try:
        y = int(y)
    except ValueError:
        print("Vous devez entrer un nombre")
        Lettre = True
        while Lettre:
            y = input("Chaussure de décharge S :")
            try:
                y = int(y)
            except ValueError:
                print("Vous devez entrer un nombre")
                Lettre = True
            else:
                Lettre = False 
    df2.at[24,"Stock"] = y

    writer = ExcelWriter(Nom)
    df_name.to_excel(writer, Nom,index=False, startcol=0)
    df2.to_excel(writer,Nom,index=False, startcol=1)
    writer.save()
    Main_Menu()

def Sortie_Menu():
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:

        print ("RELEVE SORTIES\n-----------------------\n1 : Relever sorties\n2 : Vérification 1\n3 : Vérification 2\n-----------------------\nB : back\nX : Quitter\n")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "1" :
            clear()
            Sortie()
            break

        elif User_Command == "2" :
            clear()
            Verif_1()
            break

        elif User_Command == "3" :
            clear()
            Verif_2()
            break

        elif User_Command == "B" :
            Main_Menu()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                Terminal_Launched = False
                clear()
            elif Choix == "B":
                clear()
                Main_Menu()
                break

        else :
            print ("Entrez une des valeur correspondante !")

def Sortie():
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:

        print ("SORTIES\n-----------------------\nEntrez la variable correspondante\n-----------------------\nB : Retour\nX : Quitter\n  ")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "o" :
            df2.iloc[14]['Sorties'] = df2.iloc[14]['Sorties'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "x" :
            df2.iloc[23]['Sorties'] = df2.iloc[23]['Sorties'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "y" :
            df2.iloc[24]['Sorties'] = df2.iloc[24]['Sorties'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "B" :
            clear()
            Physics_Menu()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                Terminal_Launched = False
                clear()
            elif Choix == "B":
                clear()
                Sortie()
                break

        else :
            print ("Entrez une des valeur correspondante !")


def Verif_1():
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:

        print ("VERIFICATION 1\n-----------------------\nEntrez la variable correspondante\n-----------------------\nB : Retour\nX : Quitter\n  ")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "o" :
            df2.iloc[14]['Verification 1'] = df2.iloc[14]['Verification 1'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "x" :
            df2.iloc[23]['Verification 1'] = df2.iloc[23]['Verification 1'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "y" :
            df2.iloc[24]['Verification 1'] = df2.iloc[24]['Verification 1'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "B" :
            clear()
            Physics_Menu()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                Terminal_Launched = False
                clear()
            elif Choix == "B":
                clear()
                Sortie()
                break

        else :
            print ("Entrez une des valeur correspondante !")

def Verif_2():
    Terminal_Launched = True
    User_Command = ""

    while Terminal_Launched:

        print ("VERIFICATION 2\n-----------------------\nEntrez la variable correspondante\n-----------------------\nB : Retour\nX : Quitter\n  ")
        User_Command = input("[Ecrivez ici]>")

        if User_Command == "o" :
            df2.iloc[14]['Verification 2'] = df2.iloc[14]['Verification 2'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "x" :
            df2.iloc[23]['Verification 2'] = df2.iloc[23]['Verification 2'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "y" :
            df2.iloc[24]['Verification 2'] = df2.iloc[24]['Verification 2'] + 1
            writer = ExcelWriter(Nom)
            df_name.to_excel(writer, Nom,index=False, startcol=0)
            df2.to_excel(writer,Nom,index=False, startcol=1)
            writer.save()
            clear()
            Sortie()
            break

        elif User_Command == "B" :
            clear()
            Physics_Menu()
            break

        elif User_Command == "X" :
            clear()
            print ("Êtes vous sûr ?\nX : Oui, Quitter\nB : Non, Retour")
            Choix = input ("Choix :")
            if Choix == "X":
                Terminal_Launched = False
                clear()
            elif Choix == "B":
                clear()
                Sortie()
                break

        else :
            print ("Entrez une des valeur correspondante !")

def clear():
    if name == 'nt':
        _ = system('cls')

    else:
        _ = system('clear')

Main_Menu()
