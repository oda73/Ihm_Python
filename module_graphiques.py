from tkinter import *
import tkinter as tk
from tkinter import ttk
from time import sleep

class progress_bar():
    def __init__(self,fen,nb_it=1,longueur=200,col="blue",mode="determinate"):
        style = ttk.Style()
        style.theme_use('alt')
        style.configure("couleur.Horizontal.TProgressbar",
                        foreground=col, background=col)

        self.fen=fen
        self.mode=mode
        self.progress_var = tk.DoubleVar()
        self. pourcentage_var=tk.StringVar()
        self.pourcentage_var.set("     %")
        self.barre=ttk.Progressbar(fen,length=longueur,variable=self.progress_var,style="couleur.Horizontal.TProgressbar",mode=mode,maximum=5)
        if self.mode == "determinate" :
            self.lab_pourcent = tk.Label(fen, textvariable=self.pourcentage_var)
            self.barre = ttk.Progressbar(fen, length=longueur, variable=self.progress_var,style="couleur.Horizontal.TProgressbar", mode=mode, maximum=100)
        self.pas=float(100/nb_it)
        self.progress=0
        self.pas_indeterminate=longueur/10

    def progression(self):
        if self.mode == "determinate":
            self.progress+=self.pas
            self.progress_var.set(self.progress)
            self.pourcentage_var.set(f"{str(int(self.progress))} %")
            print(f"{str(int(self.progress))} %")
        if self.mode == "indeterminate":
            self.barre.start(200)

        self.fen.update()




if __name__ == "__main__":
    root=Tk()
    barre1=progress_bar(root,10,200,"yellow","indeterminate")
    barre1.barre.grid(row=1,column=0)
    #barre1.lab_pourcent.grid(row=0,column=1)
    barre1.progression()
    for i in range(1,10):
        sleep(0.5)
        #barre1.progression()
        root.update()
    barre1.barre.destroy()
    #barre1.lab_pourcent.destroy()
    root.mainloop()
