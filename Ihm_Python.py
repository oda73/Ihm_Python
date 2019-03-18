from tkinter import *
from tkinter.messagebox import *
import subprocess
import json
from tkinter import filedialog
import os


class ComboBox(Toplevel):
    "widget composite boite de liste"

    def __init__(self, boss, items=["choix1","choix2"],id=1):
        self.win=boss
        self.eventid=f'<<fincombobox{id}>>'
        self.liste = items
        self.fen_choix = Toplevel(boss)
        self.fen_choix.title("choix tech")
        self.fen_choix.geometry("+600+250")
        self.fen_choix.withdraw()
        lab = Label(self.fen_choix, text="selectionnez le tech à affecter au swan")
        lab.pack(padx=5, pady=5)
        self.liste_choix= Listbox(self.fen_choix, selectbackground="ORANGE", relief=RAISED,
                                    activestyle="none", bd=7)
        for index in range(len(self.liste)):
            self.liste_choix.insert(index + 1, self.liste[index])
        self.liste_choix.pack(fill=X, expand=YES, padx=5, pady=5)
        Button(self.fen_choix, text="OK", command=self.action).pack(pady=5)

    def aff(self):
        self.fen_choix.deiconify()

    def action(self):
        print(self.liste_choix.get(self.liste_choix.curselection()))
        self.resultat="resultats :" + self.liste_choix.get(self.liste_choix.curselection())
        self.fen_choix.withdraw()
        self.win.event_generate(self.eventid)









class interface_IHM(object):
    def __init__(self,action,repjson):

        """construteur de la fenetre principal"""

        with open(action) as action_ihm:
            self.action_ihm_dict =json.load(action_ihm)
        with open("json\\tech.json")  as mailstechs:
            self.mailstechs_dict = json.load(mailstechs)
        self.rep_json = repjson
        self.liste_action_ihm =[action for action in self.action_ihm_dict.values()]
        self.root = Tk()
        self.root.title("                            ROBOT SWAN INTERFACE      V2.0")
        self.root.geometry("+1+1")
        self.menu()
        self.traitement_lot = FALSE
        #self.mailtech=""
        self.makeWidgets()
        self.root.protocol("WM_DELETE_WINDOW", self.quitter)
        self.root.bind('<<thread_fini>>', self.nettoyagelog)
        self.root.bind('<<fincombobox1>>',self.actioncombo1)
        self.root.bind('<<fincombobox2>>', self.actioncombo2)
        self.root.mainloop()


    def makeWidgets(self):

        fr_bas = Frame(self.root)
        fr_bas.pack(fill=BOTH, side=BOTTOM)
        fr_gauche = Frame(self.root)
        fr_gauche.pack(fill=BOTH, side=LEFT)
        fr_droite = Frame(self.root)
        fr_droite.pack(fill=BOTH, side=RIGHT)
        self.fr_log = Frame(fr_bas, relief=SUNKEN,bd=5,bg="ORANGE")
        self.fr_log.pack(side=TOP,expand=1,pady=5)

        # frame numero swan
        fr_gauche_haut = Frame(fr_gauche, bd=8, relief=RAISED)
        fr_gauche_haut.pack(side=TOP, padx=10, pady=10)
        Label(fr_gauche_haut, text="N° de Swan à Traiter", font="arial 10 bold").grid(row=0, column=0)
        self.entree = Entry(fr_gauche_haut, font="arial 10 bold", width=11, justify="center", bg="ORANGE")
        self.entree.grid(row=1, column=0, pady=5)

        # frame traitement par lot
        fr_gauche_bas = Frame(fr_gauche, bd=8, relief=RAISED)
        fr_gauche_bas.pack(side=BOTTOM, padx=5, pady=5)
        Label(fr_gauche_bas, text="Traitement par lot", font="arial 10 bold").grid(row=0, columnspan=2)
        Label(fr_gauche_bas, text="Selection fichier d'import", ).grid(row=1, column=0)
        self.fic_trait = StringVar()
        self.fic_trait.set("")
        self.entree2 = Entry(fr_gauche_bas, width=35, bg="ORANGE", textvariable=self.fic_trait)
        self.entree2.grid(row=2, columnspan=2, pady=5)
        Button(fr_gauche_bas, text="Parcourir", command=self.thread_traitement_lot, height=1, width=10,
               relief=RAISED,
               bd=4).grid(row=1,
                          column=1, padx=5, pady=5
                          )
        Button(fr_gauche_bas, text="Modele", command=self.ouvrir_modele_traitlot, height=1, width=10,
               relief=RAISED, bd=4).grid(row=3,
                                         columnspan=2,
                                         padx=5,
                                         pady=5)

        # liste choix
        # frame selection action
        fr_d = Frame(fr_droite, bd=4, relief=RAISED)
        fr_d.pack(side=RIGHT, padx=10)
        liste_action = Variable(fr_d, (self.liste_action_ihm))

        self.liste = Listbox(fr_d, listvariable=liste_action, selectmode="single", selectbackground="ORANGE",
                             height=10, width=38,activestyle="none")
        self.liste.grid(row=0, column=0, padx=5)
        Button(fr_d, text="OK", command=self.action).grid(row=1, column=0, pady=10)

        # bouton quitter
        Button(fr_bas, text="Quitter", command=self.quitter, height=1, width=20).pack(side=BOTTOM, pady=5)

        # bandeaux log
        self.sv = StringVar()
        self.lab1 = Label(self.fr_log, textvariable=self.sv,bg="ORANGE")
        self.lab1.grid(row=0, pady=5)
        self.lab2 = Label(self.fr_log,bg="ORANGE", text="                                                                                                                                ")
        self.lab2.grid(row=1, pady=5)

        self.sv2 = StringVar()
        self.lab3 = Label(self.fr_log, textvariable=self.sv2, fg="blue",bg="ORANGE")
        self.lab3.grid(row=2,pady=5)
        self.lab4 = Label(self.fr_log, text="",bg="ORANGE")
        self.lab4.grid(row=3, pady=5)



    def nettoyagelog(self):
        # print("nettoyage log fin de thread")
        # self.sv.set("")
        # self.lab1.update()
        # self.barre1.barre.destroy()
        # self.mailtech = ""
        pass

    def menu(self):
        """"gestion menu fen root"""
        self.men = Menu(self.root)
        self.men_admin = Menu(self.men, tearoff=0)
        self.men.add_cascade(label="Administration", menu=self.men_admin)
        self.men_admin.add_command(label="Modif fichiers json", command=self.ouvrir_json)

        self.root.config(menu=self.men)

    def ouvrir_json(self):

        repfic = filedialog.askopenfilename(defaultextension=".json",initialdir=self.rep_json)
        self.sv.set(f"modif fichier {os.path.basename(repfic)} en cours .....")  # affichage du traitement en cours
        self.lab1.update()
        subprocess.call(["notepad.exe",repfic ])
        self.sv.set("")
        self.lab1.update()
        showwarning("Alerte     ROBOT SWAN    ",
                    message="Attention pour prendre en compte les modifications veuillez redémarrer l'application")
        self.quitter()

    def trait_fichier_lot(self):
        pass
    def thread_traitement_lot(self):
        pass
    def ouvrir_modele_traitlot(self):
        pass
        #subprocess.call(["C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE","C:\Applications\Robotswan\\trait_lot\modele_trait_lot_swan_1.xls"])

    def action(self):
        #print(self.liste.curselection()[0])

        # if self.liste.curselection()==() : # verification action sélectionnée
        #     self.sv.set("veuillez choisir une action !!! ")
        # elif self.entree.get()=="" :
        #     self.sv.set("Veuiller saisir une réference !!!!")  # affichage msg erreur dans bandeau
        # else :
            #realiser les actions en fonctions de l'action selectionnée

            # ref_saisie = self.entree.get()
            # self.choix = str(self.liste.get(self.liste.curselection()))
            # self.sv.set(f"traitement en cours : {ref_saisie} -> {self.choix}")
        if self.liste.curselection()==(0,) :
            self.p1=ComboBox(self.root,["choix1","choix2","choix3"])
            self.p1.aff()
        if self.liste.curselection()==(1,) :
            self.p2=ComboBox(self.root,["choix4","choix5","choix6"],2)
            self.p2.aff()

    def actioncombo1(self,e):
        print(self.p1.resultat)

    def actioncombo2(self, e):
        print(self.p2.resultat)



    def quitter(self):
        #swan.quit()
        self.root.quit()


if __name__ == "__main__":
    app=interface_IHM("json\\menu_ihm.json","json")
