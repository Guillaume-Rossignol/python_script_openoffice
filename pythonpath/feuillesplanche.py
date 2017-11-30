import uno
import os
import shutil
from com.sun.star.beans import PropertyValue

def generateFeuilles(context, planches):
    doc = context.getDocument()
    oDesktop = context.getDesktop()

    url = doc.URL
    syspath = uno.fileUrlToSystemPath(url)

    currentDirectory = os.path.dirname(syspath)
    destinationDirectory = os.path.join(currentDirectory, 'Planches')

    #Clean du repertoire de destination
    shutil.rmtree(destinationDirectory, True)
    os.makedirs(destinationDirectory)

    fichierModele = os.path.join(currentDirectory, 'modelefeuille.ods')


    #Parcours de toutes les planches
    for planche in planches:
        fichierDestination = os.path.join(destinationDirectory, planche.getId()+" "+planche.legume.nom+'.ods')
        urlDestination = uno.systemPathToFileUrl(fichierDestination)
        #Copie bête et méchante du modéle
        shutil.copy(fichierModele, fichierDestination)

        #Ouverture de l'ods
        PropVal = PropertyValue()
        PropVal.Name = 'Hidden'
        PropVal.Value = True
        oDataDoc = oDesktop.loadComponentFromURL(urlDestination, '_blank', 0, (PropVal,))
        DataSheets = oDataDoc.getSheets()
        DataSheet = DataSheets.getByIndex(0)

        #Remplissage des cellules
        def remplir (cellule, value):
            DataSheet.getCellRangeByName(cellule).setFormula(value)


        remplir("A1", planche.legume.nom)
        remplir("E8", planche.legume.nom)
        remplir("B3", planche.bloc)
        remplir("B4", planche.planche)
        remplir("B5", planche.iteration)
        remplir("B8", planche.commande)
        remplir("B9", planche.semi)
        remplir("B10", planche.preparation)
        remplir("B11", planche.transplant)

        ligneRecolte=12
        for date in planche.getListeRecolte():
            remplir("B"+str(ligneRecolte), date)
            ligneRecolte += 1

        remplir("E2", planche.rangs)
        remplir("E3", planche.espacement)
        remplir("E4", planche.geotextile)
        remplir("E5", planche.irrigation)
        remplir("E6", planche.legume.multicellules)
        remplir("E7", planche.legume.joursEnCellules)
        remplir("E9", planche.fournisseur)
        if planche.legume.binage > 0:
            if planche.legume.binage ==1:
                remplir("E11", str(planche.legume.binage) + " semaine")
            else:
                remplir("E11", str(planche.legume.binage) + " semaines")

        remplir("D22", "Insectes : "+planche.legume.insectes+"\nNotes planche : " +planche.notes+"\nNotes légume : "+planche.legume.notes)

        if planche.isATest():
            remplir("F1", "TEST")

        oDataDoc.store()
