#importe uno
import uno
from re import sub

def test():
    document = XSCRIPTCONTEXT.getDocument()
    feuille = getSheetByName(document, "Legumes")
    dictionnaire = dictionnaireLegumes(feuille)
    feuille.getCellByPosition(0,0).String = ' | '.join(dictionnaire.keys())

colonneLegume = 3
def calendrier():
    # récupère le document actif
    document = XSCRIPTCONTEXT.getDocument()
    # prépare le travail sur les feuilles
    feuilleDate = getSheetByName(document, "Plan de jardin")
    feuilleSemi = getSheetByName(document, "Calendrier semis")
    feuilleCommande = getSheetByName(document, "Calendrier commandes")
    feuillePrepaPlanche = getSheetByName(document, "Calendrier prépa planches")
    feuilleTransplant = getSheetByName(document, "Calendrier transplant")
    feuilleRecolte = getSheetByName(document, "Calendrier récolte")

    colonneSemi = 5
    colonneCommande = 4
    colonnePrepaPlanche = 6
    colonneTransplant = 7
    colonneRecolte = 8

    calendrierSemi = {}
    calendrierCommande = {}
    calendrierPrepaPlanche = {}
    calendrierTransplant = {}
    calendrierRecolte = {}
    
    ligneLue = 1
    #PArcours du plan de jardin pour renseigner les differents dictionnaires
    while feuilleDate.getCellByPosition(0, ligneLue).String != "" :
        remplirCalendrier(feuilleDate, calendrierSemi,colonneSemi, ligneLue)
        remplirCalendrier(feuilleDate, calendrierCommande,colonneCommande, ligneLue)
        remplirCalendrier(feuilleDate, calendrierPrepaPlanche,colonnePrepaPlanche, ligneLue)
        remplirCalendrier(feuilleDate, calendrierTransplant,colonneTransplant, ligneLue)
        remplirCalendrier(feuilleDate, calendrierRecolte,colonneRecolte, ligneLue)
        ligneLue += 1
    #Remplissage des feuilles
    remplirFeuille(feuilleSemi,calendrierSemi)
    remplirFeuille(feuilleCommande,calendrierCommande)
    remplirFeuille(feuillePrepaPlanche,calendrierPrepaPlanche)
    remplirFeuille(feuilleTransplant,calendrierTransplant)
    remplirFeuille(feuilleRecolte,calendrierRecolte)

def getSheetByName(document, name):
    if document.Sheets.hasByName(name):
        return document.Sheets.getByName(name)
    document.Sheets.insertNewByName(name,0)

def remplirCalendrier(feuille, calendrier, colonne, ligne):
    date = feuille.getCellByPosition(colonne, ligne).Value
    #Petit workaround pour créer les listes ou les appends dans les dictionnaires
    if date > 0:
        if date in calendrier:
            calendrier[date].append(feuille.getCellByPosition(colonneLegume, ligne).String)
        else:
            calendrier[date] = [feuille.getCellByPosition(colonneLegume, ligne).String]    

def remplirFeuille(feuille, calendrier):
    feuille.getCellRangeByName("A2:B365").clearContents(7)
    line = 1
    for key in sorted(calendrier.keys()):
        feuille.getCellByPosition(0, line).Value = key
        feuille.getCellByPosition(1, line).String = " | ".join(calendrier[key])
        line +=1

class Legume:
    def __init__(self, \
                 nom,\
                 multicellules,\
                 joursEnCellules,\
                 recolte,\
                 nombreRecolte,\
                 nombreRang,\
                 espacement,\
                 quantitePlanche):
        self.nom = nom
        self.multicellules = multicellules
        self.joursEnCellules = joursEnCellules
        self.recolte = recolte
        self.nombreRecolte = nombreRecolte
        self.nombreRang = nombreRang
        self.espacement = espacement
        self.quantitePlanche = quantitePlanche


def standardisationNom(nom):
    return sub('[^a-z]', '', nom.lower())

def dictionnaireLegumes(feuille):
    #configuration de la feuille
    colonneNom = 0
    colonneMultiCellule = 1
    colonneJoursEnCellules = 2
    colonneRecolte = 3
    colonneNombreRecolte = 4
    colonneNombreRang = 5
    colonneEspacement = 6
    colonneQuantitePlanche = 8

    #Parcours de la feuille et génération du bouzin
    dictionnaire = {}

    ligne = 1
    while feuille.getCellByPosition(colonneNom, ligne).String != "":
        dictionnaire[standardisationNom(feuille.getCellByPosition(colonneNom, ligne).String)] = \
                     Legume(\
                        feuille.getCellByPosition(colonneNom, ligne).String,\
                        feuille.getCellByPosition(colonneMultiCellule, ligne).String,\
                        feuille.getCellByPosition(colonneJoursEnCellules, ligne).String,\
                        feuille.getCellByPosition(colonneRecolte, ligne).String,\
                        feuille.getCellByPosition(colonneNombreRecolte, ligne).String,\
                        feuille.getCellByPosition(colonneNombreRang, ligne).String,\
                        feuille.getCellByPosition(colonneEspacement, ligne).String,\
                        feuille.getCellByPosition(colonneQuantitePlanche, ligne).String\
                    )
        ligne += 1
    return dictionnaire

    
        

