
def ChangeColor(Objet, LeaveColor, OnColor):
        
    def LeaveTool(e):
        Objet['background'] = LeaveColor
    Objet.bind("<Leave>" , LeaveTool)

    def UseTool(e):
        Objet['background'] = OnColor  
    Objet.bind("<Enter>" , UseTool)     
