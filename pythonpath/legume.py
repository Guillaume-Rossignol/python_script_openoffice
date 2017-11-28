from re import sub


class Legume:
    def __init__(
            self,
            nom,
            multicellules = 0,
            joursEnCellules = 0,
            recolte = 0,
            nombreRecolte =1,
            nombreRang = 1,
            espacement = 30,
            quantitePlanche = 0,
            packsPlanche = 0,
            plateauxPlanche = 0
            ):
        self.nom = nom
        self.multicellules = multicellules
        self.joursEnCellules = joursEnCellules
        self.recolte = recolte
        self.nombreRecolte = nombreRecolte
        self.nombreRang = nombreRang
        self.espacement = espacement
        self.quantitePlanche = quantitePlanche
        self.packsPlanche = packsPlanche
        self.plateauxPlanche = plateauxPlanche

def dictionnaireLegumes(feuille):
    #configuration de la feuille
    colonneNom = 0
    colonneMultiCellule = 1
    colonneJoursEnCellules = 2
    colonneRecolte = 3
    colonneNombreRecolte = 4
    colonneNombreRang = 5
    colonneEspacement = 6
    colonneQuantitePlanche = 9
    colonnePacksPlanche = 8
    colonnePlateauxPlanche = 10

    #Parcours de la feuille et génération du bouzin
    dictionnaire = {}

    ligne = 1
    while feuille.getCellByPosition(colonneNom, ligne).String != "":
        dictionnaire[standardisationNom(feuille.getCellByPosition(colonneNom, ligne).String)] = \
                     Legume(
                        feuille.getCellByPosition(colonneNom, ligne).String,
                        feuille.getCellByPosition(colonneMultiCellule, ligne).String,
                        feuille.getCellByPosition(colonneJoursEnCellules, ligne).String,
                        feuille.getCellByPosition(colonneRecolte, ligne).String,
                        int(feuille.getCellByPosition(colonneNombreRecolte, ligne).Value),
                        feuille.getCellByPosition(colonneNombreRang, ligne).String,
                        feuille.getCellByPosition(colonneEspacement, ligne).String,
                        feuille.getCellByPosition(colonneQuantitePlanche, ligne).String,
                        feuille.getCellByPosition(colonnePacksPlanche, ligne).String,
                        feuille.getCellByPosition(colonnePlateauxPlanche, ligne).String
                    )
        ligne += 1
    return dictionnaire
def standardisationNom(nom):
    return sub('[^a-z]', '', nom.lower())
