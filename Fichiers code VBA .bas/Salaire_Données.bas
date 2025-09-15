Attribute VB_Name = "Salaire_Donn�es"

Sub Attachement_Salaires_Donn�es()


Range("E8").Select
ActiveCell.FormulaR1C1 = InputBox("Nom", "Salaires")
Range("I8").Select
ActiveCell.FormulaR1C1 = InputBox("Pr�mon", "Salaires")
Range("I13").Select
ActiveCell.FormulaR1C1 = InputBox("Adresse 1", "Salaires")
Range("I15").Select
ActiveCell.FormulaR1C1 = InputBox("Adresse 2", "Salaires")
Range("I17").Select
ActiveCell.FormulaR1C1 = InputBox("Code postal", "Salaires")
Range("I19").Select
ActiveCell.FormulaR1C1 = InputBox("T�lphone", "Salaires")
Range("I21").Select
ActiveCell.FormulaR1C1 = InputBox("Natel", "Salaires")
Range("I25").Select
ActiveCell.FormulaR1C1 = InputBox("Date de naissance - xx.xx.xxxx", "Salaires")
Range("I27").Select
ActiveCell.FormulaR1C1 = InputBox("Etat civil", "Salaires")
Range("I29").Select
ActiveCell.FormulaR1C1 = InputBox("No AVS - 13 chiffres", "Salaires")
Range("I31").Select
ActiveCell.FormulaR1C1 = InputBox("Engagement", "Salaires")
Range("I33").Select
ActiveCell.FormulaR1C1 = InputBox("Taux d'activit�", "Salaires")
Range("I35").Select
ActiveCell.FormulaR1C1 = InputBox("Remarque", "Salaires")
Range("I39").Select
ActiveCell.FormulaR1C1 = InputBox("Mois", "Salaires")
Range("I41").Select
ActiveCell.FormulaR1C1 = InputBox("Heures", "Salaires")
Range("I43").Select
ActiveCell.FormulaR1C1 = InputBox("Montant", "Salaires")
Range("I47").Select
ActiveCell.FormulaR1C1 = InputBox("Vacances - %", "Salaires")
Range("I49").Select
ActiveCell.FormulaR1C1 = InputBox("Jours f�ri�s - %", "Salaires")
Range("I53").Select
ActiveCell.FormulaR1C1 = InputBox("AVS - %", "Salaires")
Range("I55").Select
ActiveCell.FormulaR1C1 = InputBox("Ass. ch�mage - %", "Salaires")
Range("I57").Select
ActiveCell.FormulaR1C1 = InputBox("Ass. accident - %", "Salaires")
Range("I59").Select
ActiveCell.FormulaR1C1 = InputBox("Pr�v. professionnelle", "Salaires")
Range("I61").Select
ActiveCell.FormulaR1C1 = InputBox("Ass. maternit� - %", "Salaires")


End Sub


Sub Final_Salaire_Donn�e()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil Donn�e"

Sheets.Add.Name = "Donn�e"
    Cells.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = 95
    End With
Range("A1").Select
    Call Fiche_Salaire_Donn�e

Application.StatusBar = False



End Sub


Sub Fiche_Salaire_Donn�e()


    Call Mise_en_page_Salaire_Donn�e
ActiveCell.Offset(1, 1).Range("A1:K1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire"
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -6).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Donn�es"
ActiveCell.Offset(2, -3).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Nom"
ActiveCell.Offset(0, 2).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Pr�nom"
ActiveCell.Offset(1, -5).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -8).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 4).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(3, -5).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Adresse"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Infos"
ActiveCell.Offset(0, -6).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Adresse 1"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Adresse 2"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Code postal"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "T�l�phone"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Natel"
ActiveCell.Offset(3, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Situation"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Infos"
ActiveCell.Offset(0, -6).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Date de naissance"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Etat civil"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "No AVS"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Engagement"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Taux d'activit�"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Remarques"
ActiveCell.Offset(3, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Infos"
ActiveCell.Offset(0, -6).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mois"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Heures"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Montant"
ActiveCell.Offset(3, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Ind�mnit�"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Infos"
ActiveCell.Offset(0, -6).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Vacances"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Jours f�ri�s"
ActiveCell.Offset(3, -1).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Charges"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Infos"
ActiveCell.Offset(0, -6).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "AVS"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. ch�mage"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. accident"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Pr�v. professionnelle"
ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. maternit�"


End Sub


Sub Mise_en_page_Salaire_Donn�e()


ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 10.29
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 31
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 7).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.ColumnWidth = 31
ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 11).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 12).Columns("A:A").EntireColumn.ColumnWidth = 10.29


End Sub

