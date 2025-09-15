Attribute VB_Name = "Salaire"


Sub Bouton_Nouveau_Fichier_Salaires()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
    Call Final_Salaire_Détail
    Call Final_Salaire_Gratification
    Call Final_Salaire_Salaire
    Call Final_Salaire_Donnée
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
    
    
End Sub

Sub Bouton_Année_Salaires()


Application.ScreenUpdating = False

Sheets("Salaire").Select
Rows("1:59").Select
    Selection.Insert Shift:=xlDown
Call Complément_Salaire_Salaire


End Sub


Sub Bouton_Fiche_Salaire()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Columns("A:U").Select
    Selection.Insert Shift:=xlToRight
    Call Fiche_Salaires_Fiche
    Call Transfert_Salaires__Données_à_Fiches


End Sub


Sub Bouton_Impression_Fiche_Salaire()
    
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate:=True
    
    
End Sub


Sub Bouton_Certificat_Salaires()


    Call Transfert_Salaires__Fiches_à_Certificat_Calculs
    Call Transfert_Salaires__Certificat_Calculs_à_Certificat_Salaire

End Sub



