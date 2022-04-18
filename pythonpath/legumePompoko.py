# -*- coding: utf-8 -*-
import math
import datetime
from re import sub


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
        return math.ceil(semaine/13)



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
            int(feuille.getCellByPosition(colonnetrimestre, ligne).String),
            int(feuille.getCellByPosition(colonneelevage, ligne).String),
            int(feuille.getCellByPosition(colonnecroissance, ligne).String),
            int(feuille.getCellByPosition(colonnerecolte, ligne).String),
            feuille.getCellByPosition(colonnefournisseur, ligne).String,
            int(feuille.getCellByPosition(colonnenbCaisse, ligne).String),
            feuille.getCellByPosition(colonnequantitegraine, ligne).String,
            feuille.getCellByPosition(colonnemasse, ligne).String,
            feuille.getCellByPosition(colonnetypeMotte, ligne).String,
            feuille.getCellByPosition(colonnegrGramme, ligne).String,
            feuille.getCellByPosition(colonnedensite, ligne).String
        )
        ligne += 1

    return DicoIntelligent(dictionnaire)

def convertColonneToSemaine(colonne):
    return (colonne-20) % 52

def standardisationNom(nom, typeDeSol, trimestre):
    return sub('[^a-z]', '', (nom.split(':')[0] + typeDeSol).lower()) + str(trimestre)

class Sheet:
    def __init__(self, oosheet, start2021, errorColor):
        self.errorColor = errorColor
        self.start2021 = start2021
        self.oosheet = oosheet
        if oosheet.Name == "PC":
            self.rangeLine = [
                range(3, 13), range(15, 30), range(32, 47), range(50, 64),
                range(66, 80), range(83, 94), range(95, 111), range(114, 127),
                range(131, 151), range(153, 165), range(168, 178), range(181, 191),
                range(194, 199), range(202, 220)
            ]


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

