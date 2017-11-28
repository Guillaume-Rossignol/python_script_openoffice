# -*- coding: utf-8 -*-
#importe uno
import uno
from datetime import *
from legume import *
from planche import *

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

def next_weekday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead < 0: # Target day already past this week
        days_ahead += 7
    return d + timedelta(days_ahead)
        
def createUnoService(service):
    ctx = uno.getComponentContext()
    return ctx.ServiceManager.createInstanceWithContext(service, ctx)
