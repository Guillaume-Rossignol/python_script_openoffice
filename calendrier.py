#importe uno
import uno


colonneLegume = 3
def calendrier():
    # récupère le document actif
    document = XSCRIPTCONTEXT.getDocument()
    # prépare le travail sur les feuilles
    feuilleDate = getSheetByName(document, "Plan de jardin")
    feuilleSemi = getSheetByName(document, "Calendrier semis")
    feuilleCommande = getSheetByName(document, "Calendrier commandes")
    feuillePrepaPlanche = getSheetByName(document, "Calendrier prépa planches")
    feuilleTransplant = getSheetByName(document, "Calendrier transplant")
    feuilleRecolte = getSheetByName(document, "Calendrier récolte")

    colonneSemi = 5
    colonneCommande = 4
    colonnePrepaPlanche = 6
    colonneTransplant = 7
    colonneRecolte = 8

    calendrierSemi = {}
    calendrierCommande = {}
    calendrierPrepaPlanche = {}
    calendrierTransplant = {}
    calendrierRecolte = {}
    
    ligneLue = 1
    #PArcours du plan de jardin pour renseigner les differents dictionnaires
    while feuilleDate.getCellByPosition(0, ligneLue).String != "" :
        remplirCalendrier(feuilleDate, calendrierSemi,colonneSemi, ligneLue)
        remplirCalendrier(feuilleDate, calendrierCommande,colonneCommande, ligneLue)
        remplirCalendrier(feuilleDate, calendrierPrepaPlanche,colonnePrepaPlanche, ligneLue)
        remplirCalendrier(feuilleDate, calendrierTransplant,colonneTransplant, ligneLue)
        remplirCalendrier(feuilleDate, calendrierRecolte,colonneRecolte, ligneLue)
        ligneLue += 1
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

def remplirCalendrier(feuille, calendrier, colonne, ligne):
    date = feuille.getCellByPosition(colonne, ligne).Value
    #Petit workaround pour créer les listes ou les appends dans les dictionnaires
    if date > 0:
        if date in calendrier:
            calendrier[date].append(feuille.getCellByPosition(colonneLegume, ligne).String)
        else:
            calendrier[date] = [feuille.getCellByPosition(colonneLegume, ligne).String]    

def remplirFeuille(feuille, calendrier):
    feuille.getCellRangeByName("A2:B365").clearContents(7)
    line = 1
    for key in sorted(calendrier.keys()):
        feuille.getCellByPosition(0, line).Value = key
        feuille.getCellByPosition(1, line).String = " | ".join(calendrier[key])
        line +=1
