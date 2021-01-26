def activateBouton(Bouton) :
    if Bouton["state"] == "disabled" :
        Bouton["state"] = "normal"

def arrestBouton(Bouton) :
    if Bouton["state"] == "normal" :
        Bouton["state"] = "disabled"