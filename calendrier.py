#importe uno
import uno
from re import sub
from datetime import *

colonneLegume = 3
def calendrier():
    # récupère le document actif
    document = XSCRIPTCONTEXT.getDocument()

    # Crée les différentes feuilles
    feuilleDate = getSheetByName(document, "Plan de jardin")
    feuilleLegume = getSheetByName(document, 'Legumes')
    feuilleSemi = getSheetByName(document, "Calendrier semis")
    feuilleCommande = getSheetByName(document, "Calendrier commandes")
    feuillePrepaPlanche = getSheetByName(document, "Calendrier prepa planches")
    feuilleTransplant = getSheetByName(document, "Calendrier transplant")
    feuilleRecolte = getSheetByName(document, "Calendrier recolte")
    feuilleRecapCalendrier = getSheetByName(document, "Calendrier - Recap")

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

    def names(l):
        for planche in l:
            yield planche.legume.nom
    def plancheAndNames(l):
        for planche in l:
            yield planche.getId() + '-' + planche.legume.nom

    #Remplissage des feuilles
    remplirFeuille(feuilleSemi,calendrierSemi, names)
    remplirFeuille(feuilleCommande,calendrierCommande, names)
    remplirFeuille(feuillePrepaPlanche,calendrierPrepaPlanche, plancheAndNames)
    remplirFeuille(feuilleTransplant,calendrierTransplant, plancheAndNames)
    remplirFeuille(feuilleRecolte,calendrierRecolte, plancheAndNames)


    #Remplissage des données fixes de la feuille de recap
    feuilleRecapCalendrier.getCellRangeByName("A1:G1000").clearContents(7)

    #Recuperer la valeure du premier lundi de l'année pour Ooo
    firstMonday = next_weekday(date(2017, 1, 1), 0)
    feuilleRecapCalendrier.getCellByPosition(0, 0).setFormula(str(firstMonday.year)+"-"+str(firstMonday.month)+"-"+str(firstMonday.day))
    firstMonday = feuilleRecapCalendrier.getCellByPosition(0, 0).Value
    feuilleRecapCalendrier.getCellByPosition(0, 0).String = ""

    #Ecriture des colonnes
    feuilleRecapCalendrier.getCellByPosition(1, 0).String ="Commandes"
    feuilleRecapCalendrier.getCellByPosition(2, 0).String ="Semis"
    feuilleRecapCalendrier.getCellByPosition(3, 0).String ="PrepaPlanche"
    feuilleRecapCalendrier.getCellByPosition(4, 0).String ="Transplant"
    feuilleRecapCalendrier.getCellByPosition(5, 0).String ="Recolte"

    #Pour toutes les semaines de l'année
    for week in range(52):
        lineShift = week*10+1
        feuilleRecapCalendrier.getCellByPosition(0, lineShift).String = "Semaine "+str(week+1)
        for day in range(1, 8):
            currentDay = firstMonday+week*7+day-1
            feuilleRecapCalendrier.getCellByPosition(0, lineShift+day).Value = currentDay

            def checkCalendrierAndDisplay(calendrier, colonneShift, fonctionAffichage):
                if currentDay in calendrier:
                    feuilleRecapCalendrier.getCellByPosition(colonneShift, lineShift+day).setFormula("\n".join(fonctionAffichage(calendrier[currentDay])))

            checkCalendrierAndDisplay(calendrierCommande, 1, names)
            checkCalendrierAndDisplay(calendrierSemi, 2, names)
            checkCalendrierAndDisplay(calendrierPrepaPlanche, 3, plancheAndNames)
            checkCalendrierAndDisplay(calendrierTransplant, 4, plancheAndNames)
            checkCalendrierAndDisplay(calendrierRecolte, 5, plancheAndNames)

    for row in feuilleRecapCalendrier.getCellRangeByName("A1:A400").getRows():
        row.OptimalHeight = True
    for column in feuilleRecapCalendrier.getCellRangeByName("A1:H1").getColumns():
        column.OptimalWidth = True

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

def remplirFeuille(feuille, calendrier, aAfficher):
    feuille.getCellRangeByName("A2:B365").clearContents(7)
    line = 1
    for key in sorted(calendrier.keys()):
        feuille.getCellByPosition(0, line).Value = key
        value = ""

        feuille.getCellByPosition(1, line).String = " | ".join(aAfficher(calendrier[key]))
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
        return str(self.bloc)+'-'+str(self.planche)+'-'+str(self.iteration)
    
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

def next_weekday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead < 0: # Target day already past this week
        days_ahead += 7
    return d + timedelta(days_ahead)
        

