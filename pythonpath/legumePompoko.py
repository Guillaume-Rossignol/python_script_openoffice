# -*- coding: utf-8 -*-
import math
import datetime
from re import sub
from messagebox import MessageBox

def convint(chaine):
    if chaine == "":
        return 0
    return int(chaine)

def convertSemaineToTrimestre(semaine):
    return math.ceil(semaine / 13)
class Legume:
    def __init__(
            self,
            nom,
            typeEmplacement = "PC", #PC pour ce qui est planté en sol, SA pour la serre
            trimestre = 0, #0: toute l'année, sinon, [1,4]
            elevage = 0, #temps en semis
            croissance = 0, #temps en terre
            recolte =1, # durée pendant laquelle on peut récolter apres la croissance (>0)
            fournisseur = "",
            nbCaisse = 30, #nombre de caisse par planche
            quantiteGraine = 0, #par plance
            masse = 0, #masse (en kg) recolté par planche
            formatMotte = 0, #0 : non concerné, 1 petite motte, 2 grosse motte
            grGramme = "", #aucune idée de ce que je doit faire de cette information
            densite = "", #densité ?!?
            ):
        self.recolte = recolte
        self.croissance = croissance
        self.fournisseur = fournisseur
        self.nbCaisse = nbCaisse
        self.masse = masse
        self.quantiteGraine = quantiteGraine
        self.formatMotte = formatMotte
        self.grGramme = grGramme
        self.densite = densite
        self.trimestre = trimestre
        self.elevage = elevage
        self.nom = nom
        self.typeEmplacement = typeEmplacement

    def isConfigured(self):
        return self.recolte > 0 and self.croissance > 0

    def aColoriser(self):
        return self.recolte + self.croissance

class DicoIntelligent:
    def __init__(self, dico):
        self.dico = dico

    def getInfos(self, nom, typeDeSol, semaine):
        trimName = standardisationNom(nom, typeDeSol, self.convertSemaineToTrimestre(semaine))
        genericName = standardisationNom(nom, typeDeSol, 0)

        if trimName in self.dico:
            return self.dico[trimName]
        if genericName in self.dico:
            return self.dico[genericName]

        return None

    def convertSemaineToTrimestre(self, semaine):
        return convertSemaineToTrimestre(semaine)



def dictionnaireLegumes(feuille):
    #configuration de la feuille
    colonneNom = 0
    colonnetypeSol = 1
    colonnetrimestre = 2
    colonneelevage = 3
    colonnecroissance = 4
    colonnerecolte = 5
    colonnefournisseur = 6
    colonnenbCaisse = 7
    colonnequantitegraine = 8
    colonnemasse = 9
    colonnetypeMotte = 10
    colonnegrGramme = 11
    colonnedensite = 12

    #Parcours de la feuille et génération du bouzin
    dictionnaire = {}

    ligne = 1
    while feuille.getCellByPosition(colonneNom, ligne).String != "":
        dictionnaire[standardisationNom(
            feuille.getCellByPosition(colonneNom, ligne).String,
            feuille.getCellByPosition(colonnetypeSol, ligne).String,
            feuille.getCellByPosition(colonnetrimestre, ligne).String
        )] = Legume(
            feuille.getCellByPosition(colonneNom, ligne).String,
            feuille.getCellByPosition(colonnetypeSol, ligne).String,
            convint(feuille.getCellByPosition(colonnetrimestre, ligne).String),
            convint(feuille.getCellByPosition(colonneelevage, ligne).String),
            convint(feuille.getCellByPosition(colonnecroissance, ligne).String),
            convint(feuille.getCellByPosition(colonnerecolte, ligne).String),
            feuille.getCellByPosition(colonnefournisseur, ligne).String,
            convint(feuille.getCellByPosition(colonnenbCaisse, ligne).String),
            feuille.getCellByPosition(colonnequantitegraine, ligne).String,
            feuille.getCellByPosition(colonnemasse, ligne).String,
            feuille.getCellByPosition(colonnetypeMotte, ligne).String,
            feuille.getCellByPosition(colonnegrGramme, ligne).String,
            feuille.getCellByPosition(colonnedensite, ligne).String
        )
        ligne += 1

    #Controle des données du dico
    erreur = [legume.nom for legume in dictionnaire.values() if (legume.recolte < 1 or legume.croissance < 1)]
    if len(erreur) > 0:
        feuille.getCellRangeByName('O2').setFormula("Erreur sur les légumes : "+"|".join(erreur))
        return None

    return DicoIntelligent(dictionnaire)

def standardisationNom(nom, typeDeSol, trimestre):
    return sub('[^a-z]', '', (nom.split(':')[0] + typeDeSol).lower()) + str(trimestre)

class Sheet:
    def __init__(self, oosheet, errorColor):
        self.errorColor = errorColor
        self.oosheet = oosheet
        if oosheet.Name == "PC":
            self.rangeLine = [
                range(3, 13), range(15, 30), range(32, 47), range(50, 64),
                range(66, 80), range(83, 94), range(95, 111), range(114, 127),
                range(131, 151), range(153, 165), range(168, 178), range(181, 191),
                range(194, 199), range(202, 220)
            ]
            self.start2021 = 177
        elif oosheet.Name == "SA":
            self.start2021 = 176
            self.rangeLine = [
                range(3, 11), range(13, 21), range(24, 29), range(31, 36),
                range(39, 53),
            ]

    def convertColonneToSemaine(self, colonne):
        return (colonne - self.start2021 + 1) % 52

    def convertColonneToYear(self, colonne):
        return 2021 + int((colonne - self.start2021 + 1) / 52)


    def getColonneYear(self):
        currentYear = datetime.datetime.now().year
        deltaYear = currentYear - 2021
        startColonne = self.start2021 + deltaYear * 52
        endColonne = startColonne + 52 * 2

        return range(startColonne, endColonne)

    def getValideCell(self):
        liste = []
        for range in self.rangeLine:
            for line in range:
                for colonne in self.getColonneYear():
                    cell = self.oosheet.getCellByPosition(colonne, line)
                    cellColor = cell.getPropertyValue("CellBackColor")
                    if cell.String != '' and (cellColor == -1 or cellColor == self.errorColor):
                        liste.append(cell)

        return liste

def getSheetByName(document, name):
    if document.Sheets.hasByName(name):
        return document.Sheets.getByName(name)
    document.Sheets.insertNewByName(name, 0)
    return document.Sheets.getByName(name)

def addMissingLegumes(sheet, legumes:set):
    emptyLine = 1
    while sheet.getCellByPosition(0, emptyLine).String != "":
        emptyLine += 1

    for (nom, typeDeSol, trimestre) in legumes:
        sheet.getCellByPosition(0, emptyLine).setFormula(nom)
        sheet.getCellByPosition(1, emptyLine).setFormula(typeDeSol)
        sheet.getCellByPosition(2, emptyLine).setFormula(trimestre)
        emptyLine +=1

def processComplet(document):
    bdd = getSheetByName(document, "base de données")
    config = getSheetByName(document, "configuration")

    blanc = config.getCellByPosition(0, 0).getPropertyValue('CellBackColor')
    plantation = config.getCellByPosition(1, 0).getPropertyValue('CellBackColor')
    croissance = config.getCellByPosition(1, 1).getPropertyValue('CellBackColor')
    recolte = config.getCellByPosition(1, 2).getPropertyValue('CellBackColor')
    erreur = config.getCellByPosition(1, 3).getPropertyValue('CellBackColor')

    pcSheet = Sheet(getSheetByName(document, "PC"), erreur)
    saSheet = Sheet(getSheetByName(document, "SA"), erreur)

    dicoLegume = dictionnaireLegumes(bdd)

    if dicoLegume is None:
        MessageBox(document, 'Completer la feuille «base de données». Un filtre sur la croissance et la récolte = 0 peut aider', 'Erreur bdd')
        document.getCurrentController().setActiveSheet(bdd)
        return None

    MessageBox(document, 'Lecture de la BDD fait', 'Step 1')


    def coloriser(sheet: Sheet, dicoLegumes):
        oosheet = sheet.oosheet
        typeDeSol = oosheet.Name
        erreurs = {"placePrise": [], "légume inexistant": [], 'ajout': {}}

        for cell in sheet.getValideCell():
            colonne = cell.CellAddress.Column
            year = sheet.convertColonneToYear(colonne)
            week = sheet.convertColonneToSemaine(colonne)
            ligne = cell.CellAddress.Row
            currentLegume = dicoLegumes.getInfos(cell.String, typeDeSol, sheet.convertColonneToSemaine(colonne))

            failed = False
            if currentLegume is None:
                erreurs["légume inexistant"].append((cell.String, typeDeSol, convertSemaineToTrimestre(sheet.convertColonneToSemaine(colonne))))
                failed = True
            else:
                for i in range(1, currentLegume.aColoriser() + 1):
                    if oosheet.getCellByPosition(colonne + i, ligne).getPropertyValue('CellBackColor') != blanc:
                        erreurs["placePrise"].append(cell)
                        failed = True
                        cell.setPropertyValue("CellBackColor", erreur)
                        break

            if failed == False:
                cell.setPropertyValue("CellBackColor", plantation)
                dureeCroissance = currentLegume.croissance
                dureeRecolte = currentLegume.recolte
                for i in range(1, currentLegume.croissance + 1):
                    oosheet.getCellByPosition(colonne + i, ligne).setPropertyValue('CellBackColor', croissance)
                for i in range(dureeCroissance + 1, dureeCroissance + dureeRecolte + 1):
                    oosheet.getCellByPosition(colonne + i, ligne).setPropertyValue('CellBackColor', recolte)

                if year not in erreurs['ajout']:
                    erreurs['ajout'][year] = {}
                if currentLegume not in erreurs['ajout'][year]:
                    erreurs['ajout'][year][currentLegume] = {}
                if week not in erreurs['ajout'][year][currentLegume]:
                    erreurs['ajout'][year][currentLegume][week] = 0
                erreurs['ajout'][year][currentLegume][week] += 1



        erreurs["légume inexistant"] = set(erreurs["légume inexistant"])
        return erreurs

    def completeSemis(bilan, document):
        for site in bilan:
            for year in bilan[site]:
                semisDoc = getSheetByName(document, "Semis "+str(year % 2000))
                emptyLine = 1
                while semisDoc.getCellByPosition(0, emptyLine).String != "":
                    emptyLine += 1
                emptyLine += 1
                for legume in bilan[site][year]:
                    for week in bilan[site][year][legume]:
                        quantite = bilan[site][year][legume][week]
                        semisDoc.getCellRangeByName("A"+str(emptyLine)).setFormula(legume.nom)
                        semisDoc.getCellRangeByName("C"+str(emptyLine)).setFormula(site)
                        semisDoc.getCellRangeByName("D"+str(emptyLine)).setFormula(week - legume.elevage)
                        semisDoc.getCellRangeByName("F"+str(emptyLine)).setFormula(week)
                        semisDoc.getCellRangeByName("J"+str(emptyLine)).setFormula(quantite)
                        semisDoc.getCellRangeByName("K"+str(emptyLine)).setFormula(legume.densite)
                        semisDoc.getCellRangeByName("N"+str(emptyLine)).setFormula(legume.nbCaisse * quantite)
                        semisDoc.getCellRangeByName("H"+str(emptyLine)).setFormula(legume.croissance + week)
                        emptyLine += 1

    bilanPC = coloriser(pcSheet, dicoLegume)
    MessageBox(document, 'Feuille PC : traitée', 'Step 2')


    bilanSA = coloriser(saSheet, dicoLegume)

    bilanAjout = {
        'PC': bilanPC['ajout'],
        'SA': bilanSA['ajout']
    }
    completeSemis(bilanAjout, document)
    MessageBox(document, 'Feuille SA : traitée', 'Step 3')

    missingLegumes = bilanSA["légume inexistant"].union(bilanPC["légume inexistant"])
    if len(missingLegumes) > 0:
        addMissingLegumes(bdd, missingLegumes)
        MessageBox(document, str(len(missingLegumes)) + ' légumes rajoutés à la feuille', 'Ajout légumes')
        document.getCurrentController().setActiveSheet(bdd)

    if (len(bilanSA["placePrise"]) + len(bilanPC["placePrise"])) > 0:
        MessageBox(document, str(len(bilanSA["placePrise"]) + len(bilanPC["placePrise"])) + ' légumes n\'ont pas assez de places et ont été mis en rouge', 'Soucis planning')

    MessageBox(document, 'Fini', 'Fini')

    return None



