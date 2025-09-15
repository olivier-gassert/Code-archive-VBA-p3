Attribute VB_Name = "Salaire_Gratification"
Sub Final_Salaire_Gratification()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil Gratification"

Sheets.Add.Name = "Gratification"
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(1.5)
        .BottomMargin = Application.InchesToPoints(0.25)
        .CenterHorizontally = True
        .CenterVertically = True
        .Order = xlOverThenDown
        .Zoom = 95
    End With
Range("A1").Select
    Call Complément_Salaire_Gratification

Application.StatusBar = False


End Sub

Sub Complément_Salaire_Gratification()


Call Fiche_Salaire_Gratification


End Sub


Sub Fiche_Salaire_Gratification()


Call Fiche_Salaire_Salaire
ActiveCell.Offset(1, -4).Range("A1").Select
ActiveCell.FormulaR1C1 = "Gratification"


End Sub

