# -*- coding: utf-8 -*-
#importe uno
import uno
from datetime import *
import os
import shutil
from legume import *
from planche import *
from messagebox import *
import feuillesplanche

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
# renommer les valeurs pour eviter possibles ambiguites
from com.sun.star.awt.MessageBoxResults import OK as MBR_OK

colonneLegume = 3

def GenerateFeuillesPlanche():
    doc = XSCRIPTCONTEXT.getDocument()

    #Demande une confirmation
    res = MessageBox(
        doc.CurrentController.Frame.ContainerWindow,
        "Générer les feuilles va détruire les précendentes feuilles générer.\n" +
        "Cette action prend un temps certain",
        "Confirmation génération des feuilles planches",
        WARNINGBOX,
        BUTTONS_OK_CANCEL
    )
    if res != MBR_OK:
        MessageBox(
            doc.CurrentController.Frame.ContainerWindow,
            "Rien n'a été fait",
            "Annulation",
            INFOBOX,
            MBR_YES
        )
        return

    # Crée et récupere les différentes feuilles
    def getSheetByName(document, name):
        if document.Sheets.hasByName(name):
            return document.Sheets.getByName(name)
        document.Sheets.insertNewByName(name, 0)
        return document.Sheets.getByName(name)

    feuilleDate = getSheetByName(doc, "Plan de jardin")
    feuilleLegume = getSheetByName(doc, 'Legumes')

    legumes = dictionnaireLegumes(feuilleLegume)
    planches = listePlanche(feuilleDate, legumes)


    feuillesplanche.generateFeuilles(XSCRIPTCONTEXT, planches)
    MessageBox(doc.CurrentController.Frame.ContainerWindow, 'Toutes les feuilles ont été créées', 'Création des feuilles')


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
    def addTest(planche, chaine):
        if planche.isATest():
            chaine += " test"
        return chaine
    def plancheNamesPlateaux(l):
        for planche in l:
            yield addTest(planche, planche.getId() + '-' + planche.legume.nom + "(" + str(planche.getPlateau())  + ")")
    def namesAndPacks(l):
        for planche in l:
            yield addTest(planche, planche.legume.nom + " " +str(planche.getPacks()) + " packs")
    def namesAndSemi(l):
        for planche in l:
            yield addTest(planche, planche.legume.nom + "(" + planche.getQuantitePlanche() + ")")
    def plancheAndNames(l):
        for planche in l:
            yield addTest(planche, planche.getId() + '-' + planche.legume.nom)

    #Remplissage des données fixes de la feuille de recap
    recapCellRange = feuilleRecapCalendrier.getCellRangeByName("A1:G1000")
    recapCellRange.clearContents(7)


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

    #mise en page
    for row in recapCellRange.getRows():
        row.setPropertyValue('OptimalHeight', True)
    for column in recapCellRange.getColumns():
        column.setPropertyValue('OptimalWidth', True)


def next_weekday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead < 0: # Target day already past this week
        days_ahead += 7
    return d + timedelta(days_ahead)
