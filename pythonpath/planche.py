# -*- coding: utf-8 -*-

from legume import *

class Planche:
    def __init__(self,
                 legume,
                 bloc,
                 planche,
                 iteration,
                 commande,
                 semi,
                 preparation,
                 transplant,
                 premiereRecolte,
                 rangs,
                 espacement,
                 geotextile,
                 irrigation,
                 fournisseur,
                 notes
                 ):
        self.legume = legume
        self.bloc = bloc
        self.planche = planche
        self.iteration = iteration
        self.commande = float(commande)
        self.semi = float(semi)
        self.preparation = float(preparation)
        self.transplant = float(transplant)
        self.premiereRecolte = float(premiereRecolte)
        self.rangs = rangs
        self.espacement = espacement
        self.geotextile = geotextile
        self.irrigation = irrigation
        self.fournisseur = fournisseur
        self.notes = notes
    def getId(self):
        return str(self.bloc)+'-'+str(self.planche)+'-'+str(self.iteration)

    def getListeRecolte(self):
        listeRecolte = []
        for indiceRecolte in range(self.legume.nombreRecolte):
            #Pour l'instant la fréquence de récolte est forcement 7 jours
            listeRecolte.append(self.premiereRecolte + 7*indiceRecolte)
        return listeRecolte

    def getListeBinage(self):
        listeBinage = []

        # Le geotextile dispense de binage
        if self.hasGeotextile():
            return []

        # Meme s'il n'y a pas de frequence de binage, on bine toujours apres une semaine de transplant
        if not self.legume.binage:
            return [self.transplant + 7]

        # Le premier binage a toujours lieu 7 jours apres le transplant.
        currentBinage = self.transplant + 7
        while currentBinage < self.premiereRecolte:
            listeBinage.append(currentBinage)
            currentBinage += 7*self.legume.binage

        return listeBinage

    def isATest(self):
        return self.espacement != self.legume.espacement or self.rangs != self.legume.nombreRang

    def getQuantitePlanche(self):
        if not self.isATest():
            return self.legume.quantitePlanche
        if self.espacement:
            return 3000/self.espacement*self.rangs
        return -1

    def getPacks(self):
        if not self.isATest():
            return self.legume.packsPlanche
        if self.legume.quantitePack :
            return self.getQuantitePlanche()/self.legume.quantitePack
        return -1

    def getPlateau(self):
        if not self.isATest():
            return self.legume.plateauxPlanche
        if self.legume.multicellules>0:
            return self.getQuantitePlanche() / self.legume.multicellules
        return -1

    def hasGeotextile(self):
        return self.geotextile.strip().lower() == 'oui'


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
    colonneFournisseur = 9
    colonneRangs = 10
    colonneEspacement = 11
    colonneGeotextile = 12
    colonneIrrigation = 13
    colonneNotes = 14

    #Parcours de la feuille et génération du bouzin
    planches = []

    ligne = 1
    while feuille.getCellByPosition(colonneBloc, ligne).String != "":
        nomLegume = feuille.getCellByPosition(colonneNom, ligne).String
        if standardisationNom(nomLegume) in dictionnaireLegume:
            legume = dictionnaireLegume[standardisationNom(nomLegume)]
        else:
            legume = Legume(nomLegume)

        planches.append(Planche(legume,
            feuille.getCellByPosition(colonneBloc, ligne).String,
            feuille.getCellByPosition(colonnePlanche, ligne).String,
            feuille.getCellByPosition(colonneIteration, ligne).String,
            feuille.getCellByPosition(colonneCommande, ligne).Value,
            feuille.getCellByPosition(colonneSemi, ligne).Value,
            feuille.getCellByPosition(colonnePreparation, ligne).Value,
            feuille.getCellByPosition(colonneTransplant, ligne).Value,
            feuille.getCellByPosition(colonnePremiereRecolte, ligne).Value,
            feuille.getCellByPosition(colonneRangs, ligne).Value,
            feuille.getCellByPosition(colonneEspacement, ligne).Value,
            feuille.getCellByPosition(colonneGeotextile, ligne).String,
            feuille.getCellByPosition(colonneIrrigation, ligne).String,
            feuille.getCellByPosition(colonneFournisseur, ligne).String,
            feuille.getCellByPosition(colonneNotes, ligne).String,
                                ))

        ligne += 1
    return planches
