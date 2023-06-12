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

def main():
    pass

if __name__ == "__main__":
    main()