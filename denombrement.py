from pathlib import Path
from tkinter import *
import tkinter.font as tkfont
import email
import smtplib
from tkinter import messagebox as msg
from math import *
import json
import os, sys
#import win32com.client as win32


class box1(Frame):


    def __init__(self, boss, nbdilution=6, font= "Times New Roman", command = '', width = 870, padx=5, pady=5):
        Frame.__init__(self, boss)
        
        self.nbdilution = nbdilution
        #Innitialisation de la liste à double dimmention qui vas contenire 
        # les valeurs de tous les champs
        self.varfields = [NONE] * nbdilution
        for i in range(nbdilution):
            self.varfields[i] = [NONE] * 2
            for j in range(2):
                self.varfields[i][j] = StringVar()  

        self.command = command
        self.isvalid = False
        self.Errors = list() #la variable contenant la liste des erreures
        self.fields_values = dict()   #Contient les données (Nbr de colonies) des boites lues
        self.sommeBoitesRetenues = 0    #la somme de tous les boites retenues
        self.nbBoitesretenues = {}     #il sera une liste de tuple qui contiendra les dilution retenues et le nombre deboite pour chaque
        self.volume = DoubleVar(value = 1.0)   #volume d'ensemencement

        """ création et gestion de l'interface du logiciel """

        #les champs d'entrée
        self.field = [NONE] * nbdilution
        self.framefield = list()
        self.font_field = tkfont.Font(family = "Times New Roman", size = 12, weight = 'normal') 
        for i in range(nbdilution):
            self.field[i] = list()
            self.framefield.append(Frame(self, relief = "groove", bd = 2))
            Label(self.framefield[i], text = f'Dilution {i+1}').grid(row = 0, columnspan = 2)
            for j in range(2):
                self.field[i].append(Entry(self.framefield[i],font = self.font_field, textvariable = self.varfields[i][j], width = 10))
                self.field[i][j].grid(row = 1, column = j, padx = 2.5, pady = 5)
            self.framefield[i].grid(row = 1, column = i+1, padx = 7, pady = 10)

        # Mes boutons
        self.font_btn = tkfont.Font(family = "Times New Roman", size = 12, weight = 'normal')   
        self.btn = list()
        nomboutons = ('Calculer', 'Réinitialiser')
        for i in range(len(nomboutons)):
            self.btn.append(Button(self, text = nomboutons[i], font =self.font_btn, fg = "white", width = 15 , command=lambda arg = nomboutons[i]: self.action(arg)))
            self.btn[i].grid(row = 2, columnspan = 3, column = i+2, pady = 5)
            if nomboutons[i] == "Calculer":
                self.btn[i].config(bg = "green")
            else:
                self.btn[i].config(bg = "red")

        # Frame d'affichage du résultats
        self.font_resultat = tkfont.Font(family = "Times New Roman", size = 14, weight = 'bold')
        self.font_result = tkfont.Font(family = "Times New Roman", size = 20, weight = 'bold' )
        self.cadre_resultat = LabelFrame(self,text = 'RESULTAT (ufc/g)', font = self.font_resultat, labelanchor = 'n', relief = 'ridge', bd = '5', bg = '#fe5a5a')
        self.result = Label(self.cadre_resultat, text = '...',font = self.font_result, anchor = 'center' , bg = "#fe5a5a", fg = 'white', width = 25, height = 3)
        self.result.pack()
        self.cadre_resultat.grid(row = 3, columnspan = 4, column = 2, pady = 25)

        # designed by ibrahima gaye
        self.font_info = tkfont.Font(family = "Harlow Solid Italic", size = 14, weight = 'normal' )
        self.info = Message(self,text = "Designed by IbsomTech", width = 250,  font = self.font_info)
        self.info.grid(row = 4, column = 5, columnspan = 2, pady = 1)

        
    """ Gestion des boutons de validation, rénitialisation et de calcul """

    def action(self, choice):
        if choice == 'Valider':
            self.validate(self.varfields)

        if choice == 'Réinitialiser':
            self.erasefields(self.field)
        
        if choice == 'Calculer':
            self.Calculate()
                
    def validate(self, champs):
        #Vérification champs par champs si la valeur entrée est correcte
        if self.fields_values:
            self.fields_values.clear()
        else:
            for i in range(len(champs)):
                values = list()
                for j in range(len(champs[i])):
                    value = champs[i][j].get()
                    if value.upper() == 'NC':
                        values.append(value.upper())
                    elif value.isnumeric():
                        values.append(int(value))
                    elif value == '':
                        values.append('vide')
                    else:
                        self.Errors.append(f"La valeur dans la Case {i*2+j+1} n'est pas correcte")
                        self.field[i][j].config(bg = 'red')
                self.fields_values[i] = values

        #Vérification si tous les données entrées sont valides          
        if self.Errors:
            self.isvalid = False
        else:

            self.isvalid = True

    def erasefields(self, champs):
        for i in range(len(champs)): #réeinitialiser tous les chammps à vide
            for j in range(len(champs[i])):
                champs[i][j].delete(0, END)
                champs[i][j].configure(bg = 'white')
        self.fields_values = {}
        self.sommeBoitesRetenues = 0
        self.nbBoitesretenues = {}
        self.result.configure(text = "...") # vider l'affichage du résultat
        self.Errors.clear()     #Effacer la liste des erreurs recenser
        



    def Calculate(self):

        self.validate(self.varfields)
        if self.isvalid:
            resultat = self.resultat()
            try:
                self.result.configure(text = "N = {:.2e}   ufc/g".format(resultat))
            except:
                pass
        else:
            pass

        self.printError(self.Errors)


    def printError(self, Errors):
        
        text = ''
        #Formatage de la list des erreurs et l'enregistrer dans un variable de chaine
        for i in range(len(Errors)):    
            text += f"{i+1}: {Errors[i]}\n"
        if text:
            msg.showerror("ERREUR", text)

    """ getion de la calcul """

    def tauxDilution(self):
        d = 0
        for key in self.fields_values.keys():
            for boite in self.fields_values[key]:
                try:
                    if 30 < boite < 300:
                        d = 10**(-key-1)
                        break
                except TypeError:
                    continue
            if d:
                break
        return d

    def dictDilRetenues(self):
        for key in self.fields_values:
            i = 0
            for boite in self.fields_values[key]:
                try:
                    if 30 <= boite <= 300:
                        i += 1
                except TypeError:
                    continue
            if(i >= 1) and len(self.nbBoitesretenues) < 3:
                self.nbBoitesretenues[key] = i
                # SUM += sum(self.dictBoites[key])
                self.sommeBoitesRetenues += sum([j for j in self.fields_values[key] if isinstance(j, int) and 30 <= j <= 300 ])

        if len(self.nbBoitesretenues) == 3:
            n1, n2, n3 = tuple(self.nbBoitesretenues.values())
        elif len(self.nbBoitesretenues) == 2:
            n1, n2 = tuple(self.nbBoitesretenues.values())
            n3 = 0
        elif len(self.nbBoitesretenues) == 1:
            n1, = tuple(self.nbBoitesretenues.values())
            n2, n3 = 0, 0
        else:
            n1, n2, n3 = 0,0,0
        return n1, n2, n3

    def resultat(self):
        n = self.dictDilRetenues()
        print(type(n[0]))
        try:
            print(self.volume.get())
            return self.sommeBoitesRetenues / (self.tauxDilution() * self.volume.get() * (n[0] + (n[1] * 0.1) + (n[2]*0.01)))
        except ZeroDivisionError :
            self.Errors.append("Vous n'avez enregistrer aucune boite")


""" Gestion de la barre des menus """
def save():
    msg.showinfo('INFO', "Cette Fonctionnalité n'est pas encore disponible")

def save_as():
    msg.showinfo('INFO', "Cette Fonctionnalité n'est pas encore disponible")

def quit():
    to_quit = msg.askquestion(title = "Quiteer l'application", message = "Etes vous sûr de vouloir quitter l'application ?", icon = 'warning')
    if to_quit == 'yes':
        root.destroy()
    else:
        pass

def about():
    info = Toplevel(root)
    info.geometry("300x250+500+300")
    info.title('A propos de Denombrement')
    info.maxsize(width = 300, height =250)
    frame_about = Frame(info)
    font_about = tkfont.Font(family = "Times New Roman", size = 12, weight = 'normal' )
    Label(frame_about, text = "DENOMBREMENT", font = font_about, fg = "green").grid(row = 1, column = 2, pady = 5)
    Label(frame_about, text = "Version :  1.2.0", font = font_about, fg = "black").grid(row = 2, column = 2, pady = 5)
    Message(frame_about, text = "Ce logiciel de DENOMBREMENT a été dévelopé par Ibrahima Gaye", width = "250",font = font_about, justify = "center", fg = "black").grid(row = 3, column = 2, pady = 5)

    frame_about.pack()

    

def contact():
    mailing = Toplevel(root)
    mailing.title('Contact me')
    frame = Frame(mailing)

    font_cnt = tkfont.Font(family = "Times New Roman", size = 12, weight = 'normal' )

    objet = StringVar()
    entry = Label(frame, text = 'Objet', font = font_cnt).grid(row = 1, column = 1, padx = 3, pady = 10)
    champ_objet = Entry(frame, width = 45, textvariable = objet, font = font_cnt)
    champ_objet.focus()
    champ_objet.grid(row = 1, column = 2, padx = 10, pady = 10)

    Label(frame, text = 'Message',font = font_cnt).grid(row = 2, column = 1, padx = 3, pady = 10)
    mssg_text = Text(frame, width = 45, height = 10, font= font_cnt)
    mssg_text.grid(row = 2, column = 2, padx = 10, pady = 10)

    btn = Button(frame, text = 'Envoyer', width = 15, font = font_cnt, command = lambda arg2=mssg_text.get('1.0','end'): mailto(message=arg2))
    btn.grid(row = 3,column = 2, pady = 10)

    frame.pack()

def save_pref(arg):
    conf.set("ensemencement", arg)



#mes fonctions

def mailto(subject = 'test in function', message = 'si ce message a été envoyé par erreur,merci de bien vouloire le supprimer'):
    import smtplib, ssl

    try:
        HOST = "smtp.gmail.com"
        subjct = subject
        TO = "gayeibra@gmail.com"
        FROM = "gaye145gaye@gmail.com"
        MSG = message
        BODY = "\r\n".join((f"from: {FROM}",
                            f"To: {TO}",
                            f"Subject: {subjct}", "",
                            MSG))
        ssl_context = ssl.create_default_context()

        print(f"objet: {subjct}")
        print(f"mail: {MSG}")
        server = smtplib.SMTP_SSL(HOST,465,ssl_context)
        server.login(FROM, "AidaNiass")
        server.sendmail(FROM, [TO], BODY)
        server.quit()
    except:
        msg.showinfo('INFO', "Cette Fonction n'est pas encore disponible")

def printError(self, Errors):
    
    text = ''
    #Formatage de la list des erreurs et l'enregistrer dans un variable de chaine
    for i in range(len(Errors)):    
        text += f"{i+1}: {Errors[i]}\n"

    msg.showerror("ERREUR", text)

class Config:
    import os, sys
    import pathlib
    import json

    def __init__(self):
        self.path = os.environ.get("LOCALAPPDATA") + "\denombrement\conf.json"
        self.data = {}
        try:
            self.data = self.get()
        except FileNotFoundError:
            os.makedirs(self.path)
            self.set("instance", "False", mode = "x")

    def set(self, key, value,  mode = 'w'):
        """
        permet d'enregistrer des données de configuration dans le fichier de configuration
        sous forme de clef/valeur   (key : value)
        """
        self.data[key]=value
        data = json.dumps(self.data)
        with open(self.path, mode) as f:
            f.write(data)


    def get(self):
        """
        permet d'obtenir une valeur de configuration dans le fichier de configuration
        en se basant sur la clef (key)
        """
        with open(self.path, 'r') as f:
            self.data = json.loads(f.read())        
        return self.data

    def load(self, frame):
        if self.get()["ensemencement"] == "profondeur":
            frame.volume.set(1.0)
        else:
            frame.volume.set(0.1)



if __name__ =='__main__':
    conf = Config()
    filename = sys.argv[0]
    dir = os.path.dirname(os.path.abspath(filename))
    if conf.get()["instance"] == "False":
        conf.set("instance","True")
        root = Tk()
        root.iconbitmap(dir+'\icon.ico')
        root.title("Dénombrement")
        root.maxsize(width = 1200, height =330)
        interface = box1(root)
        conf.load(interface)


        #la barre de menu
        monmenu = Menu(root)

        first_menu = Menu(monmenu)
        first_menu.add_command(label = 'Enregistrer', command = save)
        first_menu.add_command(label = 'Enregistrer sous', command = save_as)
        first_menu.add_command(label = 'Quitter', command = quit)

        second_menu = Menu(monmenu)
        second_menu.add_radiobutton(label = "Ensemencement en profondeur     (1ml)", variable = interface.volume, value = 1.0,  command= lambda arg="profondeur":save_pref(arg))
        second_menu.add_radiobutton(label = "Ensemencement en surface         (0.1ml)", variable = interface.volume, value = 0.1, command= lambda arg="surface":save_pref(arg))

        thirth_menu = Menu(monmenu)
        thirth_menu.add_command(label = "Contacter moi", command = contact)
        thirth_menu.add_command(label = 'A props de', command = about)

        monmenu.add_cascade(label = 'Fichier', menu = first_menu ) 
        monmenu.add_cascade(label = 'préfèrences', menu = second_menu ) 
        monmenu.add_cascade(label = 'Aide', menu = thirth_menu ) 
        root.config(menu = monmenu)

        
        interface.pack()
        root.mainloop()
        conf.set("instance","False")
    else:
        print("il passe")
        pass

