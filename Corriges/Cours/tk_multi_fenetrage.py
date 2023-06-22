# Se connecter avec un mot de passe puis ouvrir des fenêtres
"""Module d'exemple de multi fenêtrage"""

import tkinter as tk
from tkinter import ttk, messagebox

class Window(tk.Toplevel):
    def __init__(self, parent, nomFenetre="Fenetre"):
        super().__init__(parent)

        self.geometry("300x150+750+500")
        self.title(nomFenetre)
        ttk.Label(self, text=nomFenetre).pack(expand=True, anchor=tk.S)
        ttk.Button(self, text="Fermer", command=self.destroy).pack(expand=True, anchor=tk.N)
        self.grab_set()


class App(tk.Tk):
    """Fenetre Principale"""
    def __init__(self):
        super().__init__()

        self.geometry("800x600+500+300")
        self.title("Fenetre principale")

        ttk.Button(self, text="Ouvrir fenetre 1", command=self.ouvre_fenetre1).pack(expand=True)
        ttk.Button(self, text="Ouvrir fenetre 2", command=self.ouvre_fenetre2).pack(expand=True)

    def test_ok(self, event):
        """Test du mot de passe"""
        if self.entry.get() == "abc123":
            self.fenMDP.destroy()
        else:
            messagebox.showwarning("Attention", "Mauvais mot de passe !")
    
    def ouvre_fenetreMDP(self):
        """Fenetre de mot de passe"""
        self.fenMDP = tk.Toplevel(self)
        self.fenMDP.title("Fenêtre de connexion")
        self.fenMDP.geometry("300x150+750+500")
        ttk.Label(self.fenMDP, text='Veuillez saisir le mot de passe : abc123').pack(expand=True, anchor=tk.S)
        self.entry = ttk.Entry(self.fenMDP)
        self.entry.bind("<Return>", self.test_ok)
        self.entry.pack(expand=True, anchor=tk.N)
        self.entry.focus_set()
        # Fenetre transitoire. Elle apparait toujours devant sa fenetre parent.
        # Si la fenetre parents est iconifiée, la fenetre transitoire aussi.
        self.fenMDP.transient(self.fenMDP.master)
        # On utilise grab_set() pour que la fenêtre Toplevel puisse recevoir des événements
        # et empêcher les utilisateurs d'interagir avec la fenêtre principale.        
        self.fenMDP.grab_set()
        # annulation du clic sur la croix de la fenetre
        self.fenMDP.wm_protocol("WM_DELETE_WINDOW",self.faitRien)

    def faitRien(self):
        pass

    def ouvre_fenetre1(self):
        self.fen1 = Window(self, "Fenetre 1")

    def ouvre_fenetre2(self):
        self.fen2 = Window(self, "Fenetre 2")

    # ferme les fenetres 1 et 2 et supprimer leurs objets respectif
    # pour permette de quitter la fenetre pricipale
    def ferme_fenetres(self, event=None):
        """Ferme les fenêtres 1 et 2"""
        if hasattr(self, 'fen1'): 
            self.fen1.destroy()
            del self.fen1

        if hasattr(self, 'fen2'):
            self.fen2.destroy()
            del self.fen2

    # raccourci clavier
    def bind_touches(self, event=None):
        """Lie les raccourcis clavier"""
        # il est possible de ne pas indiquer "event" en parametre de fonction,
        # mais dans ce cas il faut utiliser "lamdbda event: fonction()" pour le binding
        self.bind("<Escape>", self.quitter)
        self.bind("<q>", self.ferme_fenetres)

    # quitte la fenetre principale si les deux autres fenetres sont fermer
    def quitter(self, event=None):
        if not hasattr(self, "fen1") and not hasattr(self, "fen2"):
            self.quit()


if __name__ == "__main__":
    app = App()
    app.ouvre_fenetreMDP()
    app.bind_touches()
    app.mainloop()
