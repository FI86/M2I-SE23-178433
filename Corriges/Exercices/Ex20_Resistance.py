# Exercice 20
# Notre application doit faire apparaître une fenêtre comportant un dessin de la résistance,
# ainsi qu’un champ d’entrée dans lequel l’utilisateur peut encoder une valeur numérique. 
# Un bouton <Montrer> déclenche la modification du dessin de la résistance, de telle façon que 
# les trois bandes de couleur se mettent en accord
# avec la valeur numérique introduite. 
# Contrainte : Le programme doit accepter toute entrée numérique dans les limites de 1 à 999 GΩ.

# Pour rappel, la fonction des résistances électriques consiste à s’opposer 
# (à résister) plus ou moins bien au passage du courant. Les résistances 
# se présentent concrètement sous la forme de petites pièces tubulaires cerclées de 
# bandes de couleur (en général 3 ou 4).
# Dans cette exercices nous utiliserons 3 bandes pour la valeur et 1 bande pour le coefficient multiplicateur.
# Nous n'inquerons pas la bande suivante qui est la tolerance.
# Ces bandes de couleur indiquent la valeur numérique de la résistance, en fonction du code suivant :
# Noir = 0, Brun = 1, Rouge = 2, Orange = 3, Jaune = 4, Vert = 5, Bleu = 6, Violet = 7, Gris = 8, Blanc = 9.
# On oriente la résistance de manière telle que les bandes colorées soient placées à gauche. 
# La valeur de la résistance – exprimée en ohms (Ω) – s’obtient en lisant ces bandes colorées 
# également à partir de la gauche : les trois premières bandes indiquent les trois premiers chiffres 
# de la valeur numérique ; il faut ensuite accoler à ces trois chiffres un coefficient multiplicateur 
# fournie par la quatrième bande.
# Le coef multiplicateur egale 1 pour noir, 10 pour marron, 100 pour rouge, jusqu'à 1 000 000 000 (1 GΩ) pour blanc.
# Pour les valeurs inferieures à 100 on rajoute des 0 pour être sur 3 chiffres et on appliquera un coef de 0,1 (gold) ou 0,01 (silver)
# Exemple 1 : si les bandes colorées sont jaune, violette, noir et jaune.
# La valeur de cette résistance est 4700000 Ω, ou 4700 kΩ, ou encore 4,7 MΩ.
# Exemple 2 : pour 12Ω les couleurs seront marron, rouge, noir et gold.

import tkinter as Tk

class Resistance:
    def __init__(self):
        """Constructeur de la fenêtre principale"""
        self.root = Tk.Tk()
        self.root.title("Resistance")
        
        # Gestions d'événements
        self.root.bind("<Escape>", self.quitter)
        self.root.bind("<Return>", self.changeCouleurs, add="+")
        
        self.dessineResistance()

        Tk.Label(self.root, text="Entrez la valeur de la résistance, en ohms : ").grid(row=2)
        Tk.Button(self.root, text="Montrer", command=self.changeCouleurs).grid(row=3, sticky="W") 
        Tk.Button(self.root, text="Quitter", command=self.quitter).grid(row=3, sticky="E")
        self.entree = Tk.Entry(self.root, width=14)
        self.entree.grid(row=3)
        # Code des couleurs pour les valeurs de 0 à 9 ainsi que les coefficients gold et silver
        self.couleurs = ["black", "brown", "red", "orange", "yellow", "green", "blue", "purple", "grey", "white", "silver", "gold"]
        self.entree.focus_set()

        self.root.mainloop()
    
    def quitter(self, event=None):
        self.root.quit()

    def dessineResistance(self):
        """Canevas avec un modèle de résistance à trois lignes colorées"""
        self.can = Tk.Canvas(self.root, width=250, height=100, bg="ivory")
        self.can.grid(row=1, pady=5, padx=5)
        # dessin du fils de la resistance
        self.can.create_line(10, 50, 240, 50, width=5)
        # dessin de la resistance resistance
        self.can.create_rectangle(65, 30, 185, 70, fill="light grey", width=2)
        # on mémorisera les 4 lignes dans une liste
        self.ligne = []
        # Dessin des 4 lignes colorées (noires au départ) :
        for x in range(4):
            self.ligne.append(self.can.create_rectangle(70 + x * 24, 30, 82 + x * 24, 70, fill="black", width=0))

    def changeCouleurs(self, event=None):
        """Affichage des couleurs correspondant à la valeur entrée"""
        self.valeurSaisie = self.arrondiNbrChiffre(3)
        self.nbChiffre = len(self.valeurSaisie)
        self.coef = 0

        if self.valeurSaisie.isdigit() == True:
            # si le nombre saisie est inferieur à 100 on rajoute des 0 pour être sur 3 chiffres,
            # et on applique un coef x0.1 (gold) ou x 0.01 (silver)
            while self.nbChiffre < 3:
                self.valeurSaisie += "0"
                self.coef += 1
                self.nbChiffre += 1
        else: 
            self.nbChiffre = -1

        if self.nbChiffre < 1 or self.nbChiffre > 12 or int(self.valeurSaisie) <= 0:
            self.signaleErreur()
        else:
            # Coloration des 3 lignes de valeurs
            for n in range(3):
                # itemconfigure permet de changer une propriété d'un objet tkinter déjà affiché.
                self.can.itemconfigure(self.ligne[n], fill = self.couleurs[int(self.valeurSaisie[n])])
            
            # Coloration de la ligne de coefficient
            self.can.itemconfigure(self.ligne[3], fill = self.couleurs[self.nbChiffre - (3 + self.coef)])

    def signaleErreur(self):
        self.entree.configure(bg = "red")		# colorer le fond du champ
        self.root.after(1000, self.videEntree)	# après 1 seconde, effacer
        
    def videEntree(self):
        self.entree.configure(bg = "white")		# rétablir le fond blanc
        self.entree.delete(0, Tk.END)		    # efface le contenu
    
    def lireEntree(self):
        return self.entree.get()
    
    def arrondiNbrChiffre(self, nb):
        valeur = self.lireEntree()

        if len(valeur) > nb:
            valeur = str(round(float(valeur[:nb] + "." + valeur[nb:])) * (10 ** (len(valeur) - nb)))

        return valeur
    
# Programme principal :
def main():
    Resistance()

if __name__ == "__main__":
    main()