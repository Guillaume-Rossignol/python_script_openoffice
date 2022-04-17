# -*- coding: utf-8 -*-
import re
import math
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


def standardisationNom(nom, typeDeSol, trimestre):
    return sub('[^a-z]', '', (nom.split(':')[0] + typeDeSol).lower()) + str(trimestre)












class DictionnaireLegume:
    def __init__(self, feuille):
        self.feuille = feuille
        self.legumesConfigures = {}
        self.nouveauxLegumes = []


    def get_legume(self, legumeName):
        if self.legumesConfigures.has_key(legumeName):

            return self.legumesConfigures[legumeName]

        self.nouveauxLegumes.append(legumeName)

        return False

    def save(self):
        for legume in self.nouveauxLegumes:
            print(legume)
