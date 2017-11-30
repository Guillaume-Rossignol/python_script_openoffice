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
            plateauxPlanche = 0,
            binage = 0,
            notes = "",
            insectes = "",
            quantitePack = 1,
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
        self.binage = binage
        self.notes = notes
        self.insectes = insectes
        self.quantitePack = quantitePack

def dictionnaireLegumes(feuille):
    #configuration de la feuille
    colonneNom = 0
    colonneMultiCellule = 1
    colonneJoursEnCellules = 2
    colonneRecolte = 3
    colonneNombreRecolte = 4
    colonneNombreRang = 5
    colonneEspacement = 6
    colonneQuantitePack = 7
    colonneQuantitePlanche = 9
    colonnePacksPlanche = 8
    colonnePlateauxPlanche = 10
    colonneBinage = 13
    colonneNotes = 14
    colonneInsectes = 15

    #Parcours de la feuille et génération du bouzin
    dictionnaire = {}

    ligne = 1
    while feuille.getCellByPosition(colonneNom, ligne).String != "":
        dictionnaire[standardisationNom(feuille.getCellByPosition(colonneNom, ligne).String)] = \
                     Legume(
                        feuille.getCellByPosition(colonneNom, ligne).String,
                        feuille.getCellByPosition(colonneMultiCellule, ligne).Value,
                        feuille.getCellByPosition(colonneJoursEnCellules, ligne).String,
                        feuille.getCellByPosition(colonneRecolte, ligne).String,
                        int(feuille.getCellByPosition(colonneNombreRecolte, ligne).Value),
                        feuille.getCellByPosition(colonneNombreRang, ligne).Value,
                        feuille.getCellByPosition(colonneEspacement, ligne).Value,
                        feuille.getCellByPosition(colonneQuantitePlanche, ligne).String,
                        feuille.getCellByPosition(colonnePacksPlanche, ligne).String,
                        feuille.getCellByPosition(colonnePlateauxPlanche, ligne).String,
                        feuille.getCellByPosition(colonneBinage, ligne).Value,
                        feuille.getCellByPosition(colonneNotes, ligne).String,
                        feuille.getCellByPosition(colonneInsectes, ligne).String,
                        feuille.getCellByPosition(colonneQuantitePack, ligne).Value,
                    )
        ligne += 1
    return dictionnaire
def standardisationNom(nom):
    return sub('[^a-z]', '', nom.lower())
