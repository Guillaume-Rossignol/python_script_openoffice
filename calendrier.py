#importe uno
import uno
from re import sub
from datetime import *

colonneLegume = 3
def calendrier():
    # récupère le document actif
    document = XSCRIPTCONTEXT.getDocument()


    # Crée et récupere les différentes feuilles
    def getSheetByName(document, name):
        if document.Sheets.hasByName(name):
            return document.Sheets.getByName(name)
        document.Sheets.insertNewByName(name, 0)
        return document.Sheets.getByName(name)

    feuilleDate = getSheetByName(document, "Plan de jardin")
    feuilleVariables = getSheetByName(document, "Variables")
    feuilleLegume = getSheetByName(document, 'Legumes')
    feuilleRecapCalendrier = getSheetByName(document, "Calendrier - Recap")

    calendrierSemi = {}
    calendrierCommande = {}
    calendrierPrepaPlanche = {}
    calendrierTransplant = {}
    calendrierRecolte = {}
    variables = {}

    legumes = dictionnaireLegumes(feuilleLegume)
    planches = listePlanche(feuilleDate, legumes)


    variables['dateFormatId'] = feuilleVariables.getCellByPosition(1, 1).NumberFormat
    variables['annee'] = int(feuilleVariables.getCellByPosition(1, 2).Value)

    #Trick : Recuperer la valeure du premier lundi de l'année pour Ooo
    firstMonday = next_weekday(date(variables['annee'], 1, 1), 0)
    feuilleRecapCalendrier.getCellByPosition(0, 0).setFormula(str(firstMonday.year)+"-"+str(firstMonday.month)+"-"+str(firstMonday.day))
    firstMonday = feuilleRecapCalendrier.getCellByPosition(0, 0).Value
    #Clean de la feuille
    feuilleRecapCalendrier.getCellByPosition(0, 0).String = ""

    def remplirCalendrier(planche, calendrier, colonne):
        date = getattr(planche, colonne)
        # Petit workaround pour créer les listes ou les appends dans les dictionnaires
        if date > 0:
            if date in calendrier:
                calendrier[date].append(planche)
            else:
                calendrier[date] = [planche]

    def remplirCalendrierRecolte(planche, calendrier):
        dates = planche.getListeRecolte()
        for date in dates:
            if date in calendrier:
                calendrier[date].append(planche)
            else:
                calendrier[date] = [planche]

    for planche in planches:
        remplirCalendrier(planche, calendrierSemi, "semi")
        remplirCalendrier(planche, calendrierCommande, "commande")
        remplirCalendrier(planche, calendrierPrepaPlanche, "preparation")
        remplirCalendrier(planche, calendrierTransplant, "transplant")
        remplirCalendrierRecolte(planche, calendrierRecolte)

    # Fonctions utilisées pour affichers les planches dans le tableau de récap
    def plancheNamesPlateaux(l):
        for planche in l:
            yield planche.getId() + '-' + planche.legume.nom + "(" + str(planche.legume.plateauxPlanche)  + ")"
    def namesAndPacks(l):
        for planche in l:
            yield planche.legume.nom + " " +str(planche.legume.packsPlanche) + " packs"
    def namesAndSemi(l):
        for planche in l:
            yield planche.legume.nom + "(" + planche.legume.quantitePlanche + ")"
    def plancheAndNames(l):
        for planche in l:
            yield planche.getId() + '-' + planche.legume.nom

    #Remplissage des données fixes de la feuille de recap
    feuilleRecapCalendrier.getCellRangeByName("A1:G1000").clearContents(7)


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
            feuilleRecapCalendrier.getCellByPosition(0, lineShift + day).NumberFormat = variables['dateFormatId']

            def checkCalendrierAndDisplay(calendrier, colonneShift, fonctionAffichage):
                if currentDay in calendrier:
                    feuilleRecapCalendrier.getCellByPosition(colonneShift, lineShift+day).setFormula("\n".join(fonctionAffichage(calendrier[currentDay])))

            checkCalendrierAndDisplay(calendrierCommande, 1, namesAndPacks)
            checkCalendrierAndDisplay(calendrierSemi, 2, namesAndPacks)
            checkCalendrierAndDisplay(calendrierPrepaPlanche, 3, plancheAndNames)
            checkCalendrierAndDisplay(calendrierTransplant, 4, plancheNamesPlateaux)
            checkCalendrierAndDisplay(calendrierRecolte, 5, plancheAndNames)

class Legume:
    def __init__(
            self,
            nom,
            multicellules,
            joursEnCellules,
            recolte,
            nombreRecolte,
            nombreRang,
            espacement,
            quantitePlanche,
            packsPlanche,
            plateauxPlanche
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

    def getListeRecolte(self):
        listeRecolte = []
        for indiceRecolte in range(self.legume.nombreRecolte):
            #Pour l'instant la fréquence de récolte est forcement 7 jours
            listeRecolte.append(self.premiereRecolte + 7*indiceRecolte)
        return listeRecolte

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
            legume = Legume(nomLegume, 0, 0, 0, 1, 0, 0, 0, 0, 0)

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
        

