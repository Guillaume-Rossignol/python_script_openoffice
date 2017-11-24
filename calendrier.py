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
    feuilleLegume = getSheetByName(document, 'Legumes')
    feuilleSemi = getSheetByName(document, "Calendrier semis")
    feuilleCommande = getSheetByName(document, "Calendrier commandes")
    feuillePrepaPlanche = getSheetByName(document, "Calendrier prépa planches")
    feuilleTransplant = getSheetByName(document, "Calendrier transplant")
    feuilleRecolte = getSheetByName(document, "Calendrier récolte")

    calendrierSemi = {}
    calendrierCommande = {}
    calendrierPrepaPlanche = {}
    calendrierTransplant = {}
    calendrierRecolte = {}

    legumes = dictionnaireLegumes(feuilleLegume)
    planches = listePlanche(feuilleDate, legumes)



    for planche in planches:
        remplirCalendrier(planche, calendrierSemi, "semi")
        remplirCalendrier(planche, calendrierCommande, "commande")
        remplirCalendrier(planche, calendrierPrepaPlanche, "preparation")
        remplirCalendrier(planche, calendrierTransplant, "transplant")
        remplirCalendrier(planche, calendrierRecolte, "premiereRecolte")

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

def remplirCalendrier(planche, calendrier, colonne):
    date = getattr(planche, colonne)
    #Petit workaround pour créer les listes ou les appends dans les dictionnaires
    if date > 0:
        if date in calendrier:
            calendrier[date].append(planche)
        else:
            calendrier[date] = [planche]    

def remplirFeuille(feuille, calendrier):
    feuille.getCellRangeByName("A2:B365").clearContents(7)
    line = 1
    for key in sorted(calendrier.keys()):
        feuille.getCellByPosition(0, line).Value = key
        value = ""
        def names(l):
            for planche in l:
                yield planche.legume.nom

        feuille.getCellByPosition(1, line).String = " | ".join(names(calendrier[key]))
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

class Planche:
    def __init__(self,\
                 legume,\
                 bloc,\
                 planche,\
                 iteration,\
                 commande,\
                 semi,\
                 preparation,\
                 transplant,\
                 premiereRecolte):
        self.legume = legume
        self.bloc = bloc
        self.planche = planche
        self.iteration = iteration
        self.commande = float(commande)
        self.semi = float(semi)
        self.preparation = float(preparation)
        self.transplant = float(transplant)
        self.premiereRecolte = float(premiereRecolte)
    def getId(self):
        return self.bloc+'-'+self.planche+'-'+self.preparation
    
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

def listePlanche(feuille, dictionnaireLegume):
    #configuration de la feuille
    colonneBloc = 0
    colonnePlanche = 1
    colonneIteration = 2
    colonneNom = 3
    colonneCommande = 4
    colonneSemi = 5
    colonnePreparation = 6
    colonneTransplant = 7
    colonnePremiereRecolte = 8

    #Parcours de la feuille et génération du bouzin
    planches = []

    ligne = 1
    while feuille.getCellByPosition(colonneBloc, ligne).String != "":
        nomLegume = feuille.getCellByPosition(colonneNom, ligne).String
        if standardisationNom(nomLegume) in dictionnaireLegume:
            legume = dictionnaireLegume[standardisationNom(nomLegume)]
        else:
            legume = Legume(nomLegume,0,0,0,0,0,0,0)

        planches.append(Planche(legume,
            feuille.getCellByPosition(colonneBloc, ligne).String,
            feuille.getCellByPosition(colonnePlanche, ligne).String,
            feuille.getCellByPosition(colonneIteration, ligne).String,
            feuille.getCellByPosition(colonneCommande, ligne).Value,
            feuille.getCellByPosition(colonneSemi, ligne).Value,
            feuille.getCellByPosition(colonnePreparation, ligne).Value,
            feuille.getCellByPosition(colonneTransplant, ligne).Value,
            feuille.getCellByPosition(colonnePremiereRecolte, ligne).Value
                                ))

        ligne += 1
    return planches

    
        

