
from module_perso.module_graphiques import progress_bar
from tkinter import *
from tkinter.messagebox import *
import subprocess
from datetime import datetime
import json
from tkinter import filedialog





class interface_IHM(object):
    def __init__(self,fic):

        """construteur de la fenetre principal"""
        with open(fic) as swan_data:
            self.data_dict = json.load(swan_data)
        self.root = Tk()
        self.root.title("                            ROBOT SWAN INTERFACE      V2.0")
        #self.root.geometry("675x430")
        self.menu()
        self.traitement_lot = FALSE
        self.mailtech=""
        self.makeWidgets()
        self.root.mainloop()



        self.choix_utilisateur = dict([
            (f' qualification -> commentaires : {self.data_dict["choix_1"]["commentaires"]}',self.choix_1),
            (" mail  demanderessources CP Rennes", self.choix_2),
            (" mail demande ressources CP Lannion", self.choix_3),
            (" activation + validation + RDV LA POSTE", self.choix_4),
            (" activation + validation", self.choix_5),
            (" validation + RDV EUROINFO", self.choix_6),
            (" Creation RDV", self.choix_7),
            (" Envoi_mail", self.choix_8),
            (" Activation NFIT", self.choix_9),
            (" Qualif_RIE", self.choix_10),
            (" Debug fonction", self.choix_11)
        ])

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
        self.entree = Entry(fr_gauche_haut, font="arial 10 bold", width=10, bg="ORANGE")
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
        liste_action = Variable(fr_d, (f' qualification -> commentaires : {self.data_dict["choix_1"]["commentaires"]}',
                                       " mail  demanderessources CP Rennes",
                                       " Qualif_RIE",
                                       " mail demande ressources CP Lannion",
                                       " activation + validation + RDV LA POSTE", " activation + validation",
                                       " validation + RDV EUROINFO", " Creation RDV", " Envoi_mail",
                                       " Activation NFIT", " Debug fonction"))
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
        self.lab3 = Label(self.fr_log,bg="ORANGE", text="                                                                                                                 ")
        self.lab3.grid(row=1, pady=5)

        self.sv2 = StringVar()
        self.lab2 = Label(self.fr_log, textvariable=self.sv2, fg="blue",bg="ORANGE")
        self.lab2.grid(row=2,pady=5)
        self.lab4 = Label(self.fr_log, text="",bg="ORANGE")
        self.lab4.grid(row=3, pady=5)

        self.root.protocol("WM_DELETE_WINDOW", self.quitter)
        self.root.bind('<<thread_fini>>', self.nettoyagelog1)


    def menu(self):
        """"gestion menu fen root"""
        self.men = Menu(self.root)
        self.men_admin = Menu(self.men, tearoff=0)
        self.men.add_cascade(label="Administration", menu=self.men_admin)
        self.men_admin.add_command(label="Modif fichier swan.json", command=self.ouvrir_swanjson)

        self.root.config(menu=self.men)

    def ouvrir_swanjson(self):
        self.sv2.set("modif fichier swan.json en cours .....")  # affichage du traitement en cours
        self.lab2.update()
        subprocess.call(["notepad.exe", "c:\Applications\Robotswan\json\swan.json"])
        self.sv2.set("")
        self.lab2.update()
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

        if self.liste.curselection()==() :
            showinfo("Alerte","veuillez choisir une action !!! ")
        else :
            self.choix = str(self.liste.get(self.liste.curselection()))
            print(self.choix)
            if self.choix == f' qualification -> commentaires : {self.data_dict["choix_1"]["commentaires"]}' :
                pass


    def choix_1(self):
        pass
    def choix_2(self):
        pass
    def CurSel(self,numswan):
        #numswan=self.entree.get()
        print (self.choix)
        if self.choix is NONE:
            self.choix = str(self.liste.get(self.liste.curselection()))



    def choix_tech(self):
        liste_tech=[tech for tech in self.data_dict["techs_la_poste"].keys()]
        self.fen_choix_tech=Toplevel(self.root)
        self.fen_choix_tech.title("choix tech")
        self.fen_choix_tech.geometry("+900+250")
        lab=Label(self.fen_choix_tech,text="selectionnez le tech à affecter au swan")
        lab.pack(padx=5,pady=5)
        self.liste_choix_tech =Listbox(self.fen_choix_tech,selectbackground="ORANGE",relief=RAISED,activestyle="none",bd=7)
        for index in range(len(liste_tech)) :
            self.liste_choix_tech.insert(index+1,liste_tech[index])
        self.liste_choix_tech.pack(fill=X,expand=YES,padx=5,pady=5)
        Button(self.fen_choix_tech, text="OK", command=self.action2).pack()

    def action2(self):
        pass


    def nettoyagelog1(self, e):
        pass

    def quitter(self):
        #swan.quit()
        self.root.quit()

    def alert(self):
        showinfo("alerte", "Menu en construction")


if __name__ == "__main__":

    from time import sleep
    from selenium.webdriver.support import ui
    from module_perso.module_office import Outlook

    import arrow
    import pyperclip


    #date = datetime.now()
    # recuperation des données du fichier swan.json dans un dictionnaire data_dict
    app=interface_IHM("json\swan.json")
    #with open("json\swan.json") as swan_data:
    #    data_dict = json.load(swan_data)
    #swan = swan()
    #swan.ouvrir()
    #robotswan = interface()

    # sleep(2)
    # swan.rech("T807046482")
    # swan.btnmodif()
    # swan.btnenrquit()
    # print(swan.recup_info())
    # sleep(1)
    # swan.ferme()
