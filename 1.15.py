"""
This app allows user to mailmerge infos from the user interface into a word docx.

Composed of two classes (one Data // one Tkinter) and functions to pre-load the data into fields, based on user's choice

The Tkinter class is about the design, interface, entry fields and buttons. Few variables are declared inside it.
The Data class is about getting the data from the fields into the mailmerge function.


coded by kiki @VCH

"""

from mailmerge import *
import datetime
from tkinter import *
from tkinter import messagebox
import tkinter as tk
import os
import win32com.client as win32
from pathlib import Path
import shutil


# time variable for the date inside the field date of the program when it starts. 
now=datetime.datetime.now().strftime('%d.%m.%Y')
date_now=datetime.datetime.now().strftime("%Y%m%d")

# few fonctions that auto-reads data from txt files into the fields, in order to win some extra time
# they need to be launch before the program itself. Check if settings folder exists, as well as the files (folder = to 0 or containing file/s)
def id():
    cwd = os.getcwd()
    path = cwd + r"\settings"
    if os.path.isdir(path) and os.path.exists(path):
        if len(os.listdir(path)) == 0:
            print("No settings ID.txt file found, proceed with value none")
        else:
            print("loading ID.txt")
            with open("settings\ID.txt", "r", encoding="utf-8") as f:
                for word in f:
                    return(word)        
    else:
        print("Settings folder not here...from ID")

def place():
    cwd = os.getcwd()
    path = cwd + r"\settings"
    if os.path.isdir(path) and os.path.exists(path):
        if len(os.listdir(path)) == 0:
            print("No settings place.txt file found, proceed with value none")
        else:
            print("loading place.txt")
            with open("settings\place.txt", "r", encoding="utf-8") as f:
                for word in f:
                    return(word)             
    else:
        print("Settings folder not here...from place")


 
# class that contains the button/function. Each function is equivalent to a button. Except for get_data that does what it says, 
# as well as get_value that gets the initials (3 letters) so it gets different data, but it has the same purpose.
class data():

    #button output
    def open():
        os.system("explorer output")

    # we collect the data from the fields from this function and check how many ,/coma we have in the Identity/Identité/ID field
    def get_data():
        global ID
        global DT
        global INF
        global Loc
        global Agent
        global off
        global noff
        global name_opj
        global mail_opj
        ID=id_text.get()
        # 2 from theses variables we can mix and match new variables, allowing the pre load of infos but still let the user input be the infos that will
        # be passed to the mailmerge function. Point 1 and 2 are therefore connected and allow class data and class tkinter to communicate (because of global)
        # below are the variables that every function assigned to a button will return
        DT=value.get()
        INF=infr.get()
        Loc=loca.get()
        Agent = cop.get()
        off=var.get() 
        noff = off
        if noff == "OPJ":
            name_opj = "OPJ"
            mail_opj = "OPJ"
        if noff == 'AUBERSON C.':
            name_opj = "Lt. C. AUBERSON"
            mail_opj = "christophe.auberson@ne.ch"
        elif noff == 'BAECHLER S.':
            name_opj = "Com. pr. S. BAECHLER"
            mail_opj = "simon.baechler@ne.ch"
        elif noff == 'CHEVALIER F.':
            name_opj = "Lt. F. CHEVALIER"
            mail_opj = "frederic.chevalier@ne.ch"
        elif noff == 'EIGENHEER T.':
            name_opj = "Com. T. EIGENHEER"
            mail_opj = "tristan.eigenheer@ne.ch"
        elif noff == 'HUMAIR S.':
            name_opj = "Com. S. HUMAIR"
            mail_opj = "sebastien.humair@ne.ch"
        elif noff == 'BOURQUIN J.-D.':
            name_opj = "Plt. J.-D. BOURQUIN"
            mail_opj = "jean-daniel.bourquin@ne.ch"
        elif noff == 'CHUAT T.':
            name_opj = "Com. T. CHUAT"
            mail_opj = "thierry.chuat@ne.ch"
        elif noff == 'GALLET O.':
            name_opj = "Cap. O. GALLET"
            mail_opj = "olivier.gallet@ne.ch"
        elif noff == 'GARCIA M.':
            name_opj = "Com. M. GARCIA"
            mail_opj = "manuel.garcia@ne.ch"
        elif noff == 'GEISER T.':
            name_opj = "Plt. T. GEISER"
            mail_opj = "thierry.geiser@ne.ch"
        elif noff == 'GUIGNARD G.':
            name_opj = "Com. G. GUIGNARD"
            mail_opj = "gilles.guignard@ne.ch"
        elif noff == 'HAFSI S.':
            name_opj = "Com. div. S. HAFSI"
            mail_opj = "sami.hafsi@ne.ch"
        elif noff == 'JALLARD R.':
            name_opj = "Cap. R. JALLARD"
            mail_opj = "raphael.jallard@ne.ch"
        elif noff == 'KULLMANN A.':
            name_opj = "Plt. A. KULLMANN"
            mail_opj = "anthony.kullmann@ne.ch"
        elif noff == 'KRAMER P.':
            name_opj = "Cap. P. KRAMER"
            mail_opj = "philippe.kramer@ne.ch"
        elif noff == 'MOLLIER B.':
            name_opj = "Cap. B. MOLLIER"
            mail_opj = "bertrand.mollier@ne.ch"
        elif noff == 'ROCHAIX P.-L.':
            name_opj = "Com. pr. P.-L. ROCHAIX"
            mail_opj = "pierre-louis.rochaix@ne.ch"
        elif noff == 'SAUDAN A.':
            name_opj = "Cap. A. SAUDAN"
            mail_opj = "alain.saudan@ne.ch"
        elif noff == 'SESTER N.':
            name_opj = "Com. N. SESTER"
            mail_opj = "nathalie.sester@ne.ch"
        else:
            name_opj = "à compléter"

        # below are the variables that every function assigned to a button will return
        global N
        global N2
        global P
        global Date_naissance
        global Origin
        global adresse
        global add_rue
        global add_ville
        global tel
        global gender
        global gender_b
        global gender_c

        # from now on we get the ,/coma number to create a distribution pattern of the id
        # print statement based on the result of the ,/coma analysis

        # if we get 6 * , then this is ID partielle
        if ID.count(",") == 6:
            print("6 virgules // c'est une id partielle")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[2].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[5]
            add_rue=adresse[10:].rsplit("à")[0]
            add_ville=adresse[10:].rsplit("à")[1]
            tel=ID.split(",")[6].replace("."," ").replace(","," ").replace(":", " ").replace("Tél"," ").replace("Tél.", " ").replace("tél", " ").replace("tél.", " ").replace("Téléphone", " ").replace("téléphone", " ").replace("éphone", " ").replace("éphon", " ").replace("épho", " ").replace("éph", " ").replace("ép", " ").replace("natel", " ").replace("Natel", " ").replace("mobile", " ")
            gender = ID.split(",")
            gender = gender[1]
            gender = gender.strip()
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "il"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c) 

        # if we get 7 * , and a . at the end, then this is ID complète, without phone number.
        elif ID.count(",") == 7:
            print("7 virgules // c'est une id complète, sans téléphone")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[3].split("à")[0].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[7]
            add_rue=adresse[7].rsplit("à")[0]
            add_ville=adresse[7].rsplit("à")
            tel="pas de téléphone."
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "il"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c)   

        # if we get 8 * , then this is ID complète
        elif ID.count(",") == 8:
            print("8 virgules // c'est une id complète")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[3].split("à")[0].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[7]
            add_rue=adresse[10:].rsplit("à")[0]
            add_ville=adresse[10:].rsplit("à")[1]
            tel=ID.split(",")[8].replace("."," ").replace(","," ").replace(":", " ").replace("Tél"," ").replace("Tél.", " ").replace("tél", " ").replace("tél.", " ").replace("Téléphone", " ").replace("téléphone", " ").replace("éphone", " ").replace("éphon", " ").replace("épho", " ").replace("éph", " ").replace("ép", " ").replace("natel", " ").replace("Natel", " ").replace("mobile", " ")
            gender = ID.split(",")[3].strip()
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "il"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c)   

        # if we get 9 * , then this is ID infopol
        elif ID.count(",") == 9:
            print("9 virgules // c'est une id infopol")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[3].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[7]
            add_rue=ID.split(",")[7].replace("domicilié", " ").replace("domiciliée", " ")
            add_ville=ID.split(",")[8]
            tel=ID.split(",")[9].replace("."," ").replace(","," ").replace(":", " ").replace("Tél"," ").replace("Tél.", " ").replace("tél", " ").replace("tél.", " ").replace("Téléphone", " ").replace("téléphone", " ").replace("éphone", " ").replace("éphon", " ").replace("épho", " ").replace("éph", " ").replace("ép", " ").replace("natel", " ").replace("Natel", " ").replace("mobile", " ")
            gender = ID.split(",")[3].strip()
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "elle"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c) 

        # if we get 10 * , then this is ID infopol with region (né le 16.08.2020 à El Khroub, Constantine/Algérie)
        elif ID.count(",") == 10:
            print("10 virgules // c'est une id infopol avec région AND séj illégal +1 ,")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[3].split("à")[0].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[8]
            add_rue=ID.split(",")[8].replace("domicilié", " ").replace("domiciliée", " ")
            add_ville=ID.split(",")[9]
            tele=ID.split(",")[10].replace("."," ").replace(","," ").replace(":", " ").replace("Tél"," ").replace("Tél.", " ").replace("tél", " ").replace("tél.", " ").replace("Téléphone", " ").replace("téléphone", " ").replace("éphone", " ").replace("éphon", " ").replace("épho", " ").replace("éph", " ").replace("ép", " ").replace("natel", " ").replace("Natel", " ").replace("mobile", " ")
            tel=tele.split(".")[0]
            gender = ID.split(",")[3].strip()
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "elle"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c)     

        # if we get 11 * , then this is ID infopol with region (né le 16.08.2020 à El Khroub, Constantine/Algérie) AND sans permis x: séjour illégal, sous inter.... = extra , after sej illégal
        elif ID.count(",") == 11:
            print("11 virgules // c'est une id infopol avec région AND séj illégal +1 ,")
            N=ID.split(",")[0]
            All=ID.split(",")[0]
            All2=All.split(" ")
            def upper_only(All2):
                    NOM = []
                    for s in All2:
                        if s.isupper():
                            NOM.append(s)
                    return NOM
            #variable with just upercase word in the first element of the list, that is ID
            NOM = upper_only(All2)
            #way of gettint what's left before the first coma = UPERCASE - Lowercase = lowercase is what's left
            Prenom = [value for value in All2 if value not in NOM]
            #other variable, always depending on ID variable, but it transforms the none type object into a list
            N2=" ".join(NOM)
            P=" ".join(Prenom)
            Date_naissance=ID.split(",")[3].split("à")[0].replace("né le", " ").replace("née le", " ")
            Origin=ID.split(",")[4].replace("originaire de", " ")
            adresse=ID.split(",")[8]
            add_rue=ID.split(",")[8].replace("domicilié", " ").replace("domiciliée", " ")
            add_ville=ID.split(",")[9]
            tele=ID.split(",")[10]
            tel=tele.split(".")[0].replace("."," ").replace(","," ").replace(":", " ").replace("Tél"," ").replace("Tél.", " ").replace("tél", " ").replace("tél.", " ").replace("Téléphone", " ").replace("téléphone", " ").replace("éphone", " ").replace("éphon", " ").replace("épho", " ").replace("éph", " ").replace("ép", " ").replace("natel", " ").replace("Natel", " ").replace("mobile", " ")
            gender = ID.split(",")[3].strip()
            gender = gender.split(" ")
            if len(gender[0]) == 3:
                gender = "Mme"
                gender_b = "e"
                gender_c = "elle"
            else:
                gender = "M."
                gender_b = " "
                gender_c = "elle"
            return(ID, NOM, Prenom, Date_naissance, Origin, adresse, add_rue, add_ville, tel, N, N2, P, gender, gender_b, gender_c)     

        else:
            print("c'est un autre type d'ID")
            messagebox.showinfo("Attention", "L'identité n'est pas dans un format reconnu par l'application.")

    # button Prévenu
    def distribute_prevenu():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                data.get_data()
                print("génération des fichiers prévenu")

                def questions_prevenus():
                    cwd = os.getcwd()
                    path = cwd + r"\settings"
                    if os.path.isdir(path) and os.path.exists(path):
                        if len(os.listdir(path)) == 0:
                            print("No settings questions_prévenus.txt file found, proceed with value none")
                        else:
                            print("loading questions_prévenus.txt")
                         
                        with open("settings\questions_prévenus.txt", "r", encoding="utf-8") as f:
                            line = f.read()
                            for l in line:
                                return line                 
                    else:
                        print("Settings folder not here...from questions_prévenus")
                
                word = questions_prevenus()
                word = str(word)
                word = word.replace("[", "").replace("]", "").replace("\\n", "").replace('"', '').replace(",", "").replace("'", "")
                print(word)

                template_1 = "temp/01 Prevenu/auteur.docx"
                document_1 = MailMerge(template_1)
                document_1.merge(question_prevW = word, ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_1.write('output/auteur_output.docx')
                os.rename('output/auteur_output.docx', "output/%s_%s_PVA_prevenu.docx" %(date_now,N))

                template_2 = "temp/01 Prevenu/droits.docx"
                document_2 = MailMerge(template_2)
                document_2.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_2.write('output/droits_output.docx')
                os.rename('output/droits_output.docx', "output/%s_%s_droits.docx" %(date_now,N))

                template_3 = "temp/01 Prevenu/mandat.docx"
                document_3 = MailMerge(template_3)
                document_3.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_3.write('output/mandat_output.docx')
                os.rename('output/mandat_output.docx', "output/%s_%s_mandat.docx" %(date_now,N))

                template_4 = "temp/01 Prevenu/decla_pat.docx"
                document_4 = MailMerge(template_4)
                document_4.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_4.write('output/decla_pat_output.docx')
                os.rename('output/decla_pat_output.docx', "output/%s_%s_decla_pat.docx" %(date_now,N))

                template_5 = "temp/01 Prevenu/cession.docx"
                document_5 = MailMerge(template_5)
                document_5.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_5.write('output/cession_output.docx')
                os.rename('output/cession_output.docx', "output/%s_%s_cession.docx" %(date_now,N))

                template_6 = "temp/01 Prevenu/election.docx"
                document_6 = MailMerge(template_6)
                document_6.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_6.write('output/election_output.docx')
                os.rename('output/election_output.docx', "output/%s_%s_election.docx" %(date_now,N))

                template_7 = "temp/01 Prevenu/plainte.docx"
                document_7 = MailMerge(template_7)
                document_7.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_7.write('output/plainte_output.docx')
                os.rename('output/plainte_output.docx', "output/%s_%s_plainte.docx" %(date_now,N))

                template_8 = "temp/01 Prevenu/avis.docx"
                document_8 = MailMerge(template_8)
                document_8.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_8.write('output/avis.docx')
                os.rename('output/avis.docx', "output/%s_%s_avis.docx" %(date_now,N))

                template_9 = "temp/01 Prevenu/ordrecellule.docx"
                document_9 = MailMerge(template_9)
                document_9.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_9.write('output/ordrecellule.docx')
                os.rename('output/ordrecellule.docx', "output/%s_%s_ordrecellule.docx" %(date_now,N))

                template_10 = "temp/01 Prevenu/auto_perquis.docx"
                document_10 = MailMerge(template_10)
                document_10.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_10.write('output/auto_perquis.docx')
                os.rename('output/auto_perquis.docx', "output/%s_%s_auto_perquis.docx" %(date_now,N))

                template_11 = "temp/01 Prevenu/mandat_perquis.docx"
                document_11 = MailMerge(template_11)
                document_11.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_11.write('output/mandat_perquis.docx')
                os.rename('output/mandat_perquis.docx', "output/%s_%s_mandat_perquis.docx" %(date_now,N))

                template_12 = "temp/01 Prevenu/pv_perquis.docx"
                document_12 = MailMerge(template_12)
                document_12.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_12.write('output/pv_perquis.docx')
                os.rename('output/pv_perquis.docx', "output/%s_%s_pv_perquis.docx" %(date_now,N))

                template_13 = "temp/01 Prevenu/compl_toxicol.docx"
                document_13 = MailMerge(template_13)
                document_13.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_13.write('output/compl_toxicol.docx')
                os.rename('output/compl_toxicol.docx', "output/%s_%s_compl_toxicol.docx" %(date_now,N))

                template_14 = "temp/01 Prevenu/transport.docx"
                document_14 = MailMerge(template_14)
                document_14.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_14.write('output/transport.docx')
                os.rename('output/transport.docx', "output/%s_%s_transport.docx" %(date_now,N))

                template_15 = "temp/01 Prevenu/analyse_labo.docx"
                document_15 = MailMerge(template_15)
                document_15.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_15.write('output/analyse_labo.docx')
                os.rename('output/analyse_labo.docx', "output/%s_%s_analyse_labo.docx" %(date_now,N))

                template_16 = "temp/01 Prevenu/pv_saisie.docx"
                document_16 = MailMerge(template_16)
                document_16.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_16.write('output/pv_saisie.docx')
                os.rename('output/pv_saisie.docx', "output/%s_%s_pv_saisie.docx" %(date_now,N))

                template_17 = "temp/01 Prevenu/IN.docx"
                document_17 = MailMerge(template_17)
                document_17.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_17.write('output/IN.docx')
                os.rename('output/IN.docx', "output/%s_%s_IN.docx" %(date_now,N))

                template_18 = "temp/01 Prevenu/rapport_vierge.docx"
                document_18 = MailMerge(template_18)
                document_18.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_18.write('output/rapport_vierge.docx')
                os.rename('output/rapport_vierge.docx', "output/%s_%s_rapport.docx" %(date_now,N))

                template_19 = "temp/01 Prevenu/rapport_rubriques.docx"
                document_19 = MailMerge(template_19)
                document_19.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_19.write('output/rapport_rubriques.docx')
                os.rename('output/rapport_rubriques.docx', "output/%s_%s_rapport_rubriques.docx" %(date_now,N))
            else:
                print ("Repertoire output pas vide / from prévenu")
                messagebox.showinfo("Attention", "Le dossier output contient encore des fichiers. Le vider avant d'en générer d'autres.")
        else:
            print("Repertoire output n'existe pas / from prévenu")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")

    #Button PADR
    def distribute_padr():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                data.get_data()
                print("génération des fichiers padr")
                template_1 = "temp/02 Padr/padr.docx"
                document_1 = MailMerge(template_1)
                document_1.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_1.write('output/padr_output.docx')
                os.rename('output/padr_output.docx', "output/%s_%s_PVA_PADR.docx" %(date_now,N))

                template_2 = "temp/02 Padr/plainte.docx"
                document_2 = MailMerge(template_2)
                document_2.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_2.write('output/plainte_padr_output.docx')
                os.rename('output/plainte_padr_output.docx', "output/%s_%s_plainte_padr.docx" %(date_now,N))

                template_3 = "temp/02 Padr/padr_LAVI.docx"
                document_3 = MailMerge(template_3)
                document_3.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_3.write('output/padr_LAVI.docx')
                os.rename('output/padr_LAVI.docx', "output/%s_%s_PVA_PADR_LAVI.docx" %(date_now,N))

                template_4 = "temp/02 Padr/padr_mandat.docx"
                document_4 = MailMerge(template_4)
                document_4.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_4.write('output/padr_mandat.docx')
                os.rename('output/padr_mandat.docx', "output/%s_%s_PVA_PADR_mandat.docx" %(date_now,N))

                template_5 = "temp/02 Padr/plaignant_mandat.docx"
                document_5 = MailMerge(template_5)
                document_5.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent, genreW=gender, gender_bW=gender_b, gender_cW=gender_c)
                document_5.write('output/plaignant_mandat.docx')
                os.rename('output/plaignant_mandat.docx', "output/%s_%s_PVA_plaignant_mandat.docx" %(date_now,N))
            else:
                print("Repertoire output pas vide / from padr")
                messagebox.showinfo("Attention", "Le dossier output contient encore des fichiers. Le vider avant d'en générer d'autres.")
        else:
            print("Repertoire output n'existe pas / from padr")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")
  
    #Button Victime
    def distribute_victime():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                data.get_data()
                print("génération des fichiers victime")

                template_1 = "temp/03 Victime/lavi_padr.docx"
                document_1 = MailMerge(template_1)
                document_1.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent)
                document_1.write('output/lavi_output.docx')
                os.rename('output/lavi_output.docx', "output/%s_%s_PVA_LAVI.docx" %(date_now,N))
               
                template_2 = "temp/03 Victime/SAVI_CHX.docx"
                document_2 = MailMerge(template_2)
                document_2.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent)
                document_2.write('output/SAVI_CHX_output.docx')
                os.rename('output/SAVI_CHX_output.docx', "output/%s_%s_SAVI_CHX.docx" %(date_now,N))

                template_3 = "temp/03 Victime/SAVI_NE.docx"
                document_3 = MailMerge(template_3)
                document_3.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent)
                document_3.write('output/SAVI_NE_output.docx')
                os.rename('output/SAVI_NE_output.docx', "output/%s_%s_SAVI_NE.docx" %(date_now,N))

                template_4 = "temp/03 Victime/plainte.docx"
                document_4 = MailMerge(template_4)
                document_4.merge(ID_full=ID, date=DT, officier=name_opj, Infractions=INF, NOM_Prenom=ID.split(",")[0], lieu=Loc, NOMW=N2, PrenomW=P, Naissance=Date_naissance, Tel=tel, rue=add_rue, npa=add_ville, origine=Origin, agentW=Agent)
                document_4.write('output/plainte_lavi_output.docx')
                os.rename('output/plainte_lavi_output.docx', "output/%s_%s_plainte_lavi.docx" %(date_now,N))
            else:
                print("Repertoire output pas vide / from victime")
                messagebox.showinfo("Attention", "Le dossier output contient encore des fichiers. Le vider avant d'en générer d'autres.")
        else:
            print("Repertoire output n'existe pas /from victime")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")

    #Button Mandat mail
    def mail_avis_off():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                print ("génération du mail Mandat sans doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Mandat de données signalétiques de %s pour approbation' % (N)
                mail.Body = " Bonjour %s, \n\n Voici le mandat de données signalétiques de %s pour approbation. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                mail.Display(False)
            else:
                print ("génération du mail Mandat AVEC doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Mandat de données signalétiques de %s pour approbation' % (N)
                mail.Body = " Bonjour %s, \n\n Voici le mandat de données signalétiques de %s pour approbation. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                # need to cd to right folder, in order to attach the file
                cwd = os.getcwd()
                path = cwd + r"\output"
                attachment = path + "\%s_%s_mandat.docx" %(date_now,N)
                mail.Attachments.Add(attachment)
                mail.Display(False)
        else:
            print("Repertoire n'existe pas / from mandat mail")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")

    #Button Avis cellule mail
    def mail_cellule():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                print ("génération du mail Avis cellule sans doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Avis de mise en cellule de %s' % (N)
                mail.Body = " Bonjour %s, \n\n Voici la mise en cellule de %s. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                mail.Display(False)
            else:
                print ("génération du mail Avis cellule AVEC doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Avis de mise en cellule de %s' % (N)
                mail.Body = " Bonjour %s, \n\n Voici la mise en cellule de %s. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                # need to cd to right folder, in order to attach the file
                cwd = os.getcwd()
                path = cwd + r"\output"
                attachment = path + "\%s_%s_ordrecellule.docx" %(date_now,N)
                mail.Attachments.Add(attachment)
                mail.Display(False)
        else:
            print("Repertoire n'existe pas / from mail cellule")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")

    #Button Avis arrestation
    def mail_avis():
        cwd = os.getcwd()
        path = cwd + r"\output"
        if os.path.isdir(path) and os.path.exists(path):
            if len(os.listdir(path)) == 0:
                print ("génération du mail Avis arrestation sans doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Avis d\'arrestation de %s' % (N)
                mail.Body = " Bonjour %s, \n\n Bonjour M/Mme le/la procureur/e, \n\n Voici l\'avis d\'arrestation de %s. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                mail.Display(False)
            else:
                print ("génération du mail Avis arrestation AVEC doc")
                data.get_data()
                date_now=datetime.datetime.now().strftime("%Y%m%d")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = '%s' % (mail_opj)
                mail.Subject = 'Avis d\'arrestation de %s' % (N)
                mail.Body = " Bonjour %s, \n\n Bonjour M/Mme le/la procureur/e, \n\n Voici l\'avis d\'arrestation de %s. \n\n Meilleures salutations.\n\n %s " % (name_opj, N, Agent)
                # need to cd to right folder, in order to attach the file
                cwd = os.getcwd()
                path = cwd + r"\output"
                attachment = path + "\%s_%s_avis.docx" %(date_now,N)
                mail.Attachments.Add(attachment)
                mail.Display(False)
        else:
            print("Repertoire n'existe pas / from mail avis")
            messagebox.showinfo("Attention", "Le dossier output n'est plus présent. En créer un manuellement ou re-installer l'application.")

    # function that collects 3 letters from the entry field
    def get_value():
        cop = entry_XXX.get()
        cop = cop.upper()
        if len(cop) < 3 :
            messagebox.showinfo("Attention", "Il n'y a pas assez de caractères, il en faut 3.")
            pass
        elif len(cop) > 3:
            messagebox.showinfo("Attention", "Il y a trop de caractères, il en faut 3.")
            pass
        else:
            print("3/3 caractères.")
        return(cop)
 
    # function that use the pevious function (3 letters) to get all files starting with the 3 letters, and ending with .doc from datainfopol. Copying it into the output directory of this app.
    def txt():
        name = data.get_value()
        print(name)
        input=Path(r"\\nefp1\datainfopol")         
        txt_files = input.glob("%s*.doc" % name)
        output=Path("output")
        for file in txt_files:
            shutil.copy(str(file), str(output))
            print(file)

# class for the front office aka the visual app 
class app(tk.Frame):
    def __init__(self):
        self.tk.Frame.__init()
   
if __name__ == "__main__":
    window=tk.Tk()
    window.title("ID Deployer V 1.15")
    window.geometry("500x280")
    window.resizable(width=True, height=True)
    # 1 these global variables are how we can have a value in the field, and then still use the user input to be what's going to be mailmerged
    global value
    value = tk.StringVar(window)
    value.set(now)
    global cop
    cop = tk.StringVar(window)
    cop.set(ID_cop())
    global loca
    loca = tk.StringVar(window)
    loca.set(Lieu())
    global infr
    infr = tk.StringVar(window)
    infr.set(Infr())
    off = StringVar(window)
    var = StringVar(window)
    var.set("OPJ")

    #Labels
    l1=Label(window, text= "Identité")
    l1.grid(row=0,column=0, pady=2)
    l2=Label(window, text= "Date")
    l2.grid(row=2,column=0, pady=2)
    l3=Label(window, text= "OPJ")
    l3.grid(row=4,column=0, pady=2)
    l4=Label(window, text= "Infraction/s")
    l4.grid(row=6,column=0, pady=2)
    l5=Label(window, text= "Lieu de l'administratif")
    l5.grid(row=8,column=0, pady=4)
    l6=Label(window, text= "NOM / Prénom / Grade")
    l6.grid(row=10,column=0, pady=4)

    global id_text
    global date_text
    global opj_text
    global infraction_text
    global lieu_text
    global agent_text

    #Entry fields
    id_text=StringVar()
    e1=Entry(window, width=40, bd=1, textvariable=id_text)
    e1.grid(row=0, column=1, pady=2)
    date_text=StringVar()
    e2=Entry(window, bd=1, textvariable=value, width=10)
    e2.grid(row=2, column=1, pady=2, sticky=W)
    opj_text=StringVar()
    infraction_text=StringVar()
    e4=Entry(window, bd=1, textvariable=infr, width=40)
    e4.grid(row=6, column=1, pady=2)
    lieu_text=StringVar()
    e5=Entry(window, bd=1, textvariable=loca, width=25)
    e5.grid(row=8, column=1, pady=2, sticky=W)
    agent_text=StringVar()
    e6=Entry(window, bd=1, textvariable=cop, width=35)
    e6.grid(row=10, column=1, pady=2, sticky=W)
    entry_XXX=StringVar()
    e7=Entry(window, bd=1, textvariable=entry_XXX, width=6)
    e7.grid(row=14, column=1, pady=2, sticky=W)

    option = OptionMenu(window, var,
            'AUBERSON C.',
            'BAECHLER S.',
            'CHEVALIER F.',
            'EIGENHEER T.',
            'HUMAIR S.',
            'BOURQUIN J.-D.',
            'CHUAT T.',
            'GALLET O.',
            'GARCIA M.',
            'GEISER T.',
            'GUIGNARD G.',
            'HAFSI S.',
            'JALLARD R.',
            'KRAMER P.',
            'KULLMANN A.',
            'MOLLIER B.',
            'ROCHAIX P.-L.',
            'SAUDAN A.',
            'SESTER N.')
    option.grid(row=4, column=1, pady=2, sticky=W)

    #Buttons
    b1=Button(window, text="Prévenu", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.distribute_prevenu)
    b1.grid(row=11, column=0, pady=2)
    b2=Button(window,text="PADR/Plaignant", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.distribute_padr)
    b2.grid(row=12, column=0, pady=2)
    b3=Button(window,text="Victime", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.distribute_victime)
    b3.grid(row=13, column=0, pady=2)
    b4=Button(window,text="Output", width=16, borderwidth=1, relief="raised", fg="blue", activebackground="green", overrelief="sunken", command=data.open)
    b4.grid(row=11, column=1, pady=2)
    b5=Button(window,text="Fermer", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=window.quit)
    b5.grid(row=12, column=1, pady=2)
    b6=Button(window,text="Avis arrestation mail", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.mail_avis)
    b6.grid(row=13, column=2, pady=2)
    b7=Button(window,text="Avis cellule mail", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.mail_cellule)
    b7.grid(row=12, column=2, pady=2)
    b8=Button(window,text="Mandat mail", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.mail_avis_off)
    b8.grid(row=11, column=2, pady=2)
    b9=Button(window,text="Récupérer", width=16, borderwidth=1, relief="raised", activebackground="green", overrelief="sunken", command=data.txt)
    b9.grid(row=14, column=0, pady=2)

    window.mainloop()  