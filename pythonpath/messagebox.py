# -*- coding: utf-8 -*-

import uno

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE

from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
# renommer les valeurs pour eviter possibles ambiguites
from com.sun.star.awt.MessageBoxResults import YES as MBR_YES, NO as MBR_NO, CANCEL as MBR_CANCEL

# Message box utilisant le Toolkit de l'API
def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):

  ctx = uno.getComponentContext()
  sm = ctx.ServiceManager
  sv = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
  maBoite = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
  return maBoite.execute()


def TestMessageBox():
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow

  res = MBR_YES
  while res != MBR_NO:
    s = "Voulez-vous continuer ?"
    t = "Un message de Python"
    res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)

    s = "Reponse : " +str(res)
    MessageBox(parentwin, s, t, INFOBOX)
