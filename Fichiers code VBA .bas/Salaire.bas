Attribute VB_Name = "Salaire"


Sub Bouton_Nouveau_Fichier_Salaires()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Workbooks.Add
    Call Final_Salaire_D�tail
    Call Final_Salaire_Gratification
    Call Final_Salaire_Salaire
    Call Final_Salaire_Donn�e
Sheets(Array("Feuil1")).Select
    ActiveWindow.SelectedSheets.Delete
    
    
End Sub

Sub Bouton_Ann�e_Salaires()


Application.ScreenUpdating = False

Sheets("Salaire").Select
Rows("1:59").Select
    Selection.Insert Shift:=xlDown
Call Compl�ment_Salaire_Salaire


End Sub


Sub Bouton_Fiche_Salaire()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Columns("A:U").Select
    Selection.Insert Shift:=xlToRight
    Call Fiche_Salaires_Fiche
    Call Transfert_Salaires__Donn�es_�_Fiches


End Sub


Sub Bouton_Impression_Fiche_Salaire()
    
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate:=True
    
    
End Sub


Sub Bouton_Certificat_Salaires()


    Call Transfert_Salaires__Fiches_�_Certificat_Calculs
    Call Transfert_Salaires__Certificat_Calculs_�_Certificat_Salaire

End Sub



