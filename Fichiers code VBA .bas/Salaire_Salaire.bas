Attribute VB_Name = "Salaire_Salaire"
Sub Final_Salaire_Salaire()


Application.ScreenUpdating = False
Application.StatusBar = "Feuil Salaire"

Sheets.Add.Name = "Salaire"
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
    Call Complément_Salaire_Salaire

Application.StatusBar = False


End Sub


Sub Complément_Salaire_Salaire()

Range("A1").Select
    Call Fiche_Salaire_Salaire
Range("V1").Select
    Call Fiche_Salaire_Salaire
Range("AQ1").Select
    Call Fiche_Salaire_Salaire
Range("BL1").Select
    Call Fiche_Salaire_Salaire
Range("CG1").Select
    Call Fiche_Salaire_Salaire
Range("DB1").Select
    Call Fiche_Salaire_Salaire
Range("DW1").Select
    Call Fiche_Salaire_Salaire
Range("ER1").Select
    Call Fiche_Salaire_Salaire
Range("FM1").Select
    Call Fiche_Salaire_Salaire
Range("GH1").Select
    Call Fiche_Salaire_Salaire
Range("HC1").Select
    Call Fiche_Salaire_Salaire
Range("HX1").Select
    Call Fiche_Salaire_Salaire
Rows("60").Select
    ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
    
    
End Sub

Sub Fiche_Salaire_Salaire()


    Call Mise_en_page_Salaire_Salaire
ActiveCell.Offset(1, 3).Range("A1:O1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire"
ActiveCell.Offset(0, 1).Range("A1:B1").Select
    Selection.Interior.ColorIndex = 15
ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -15).Range("A1:O1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.FormulaR1C1 = "Fiche"
ActiveCell.Offset(2, -2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Nom"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Prénom"
ActiveCell.Offset(0, 2).Range("A1:G1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Période"
ActiveCell.Offset(1, -11).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
ActiveCell.Offset(1, -16).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 6).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(0, 6).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
ActiveCell.Offset(3, -13).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Fix"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Nombre"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Montant"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Total"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 0).Range("A1").Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mois"
ActiveCell.Offset(0, 16).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Heures"
ActiveCell.Offset(0, 16).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-4]"
ActiveCell.Offset(2, -17).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(0, 12).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire"
ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-4]C+R[-2]C"
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium
ActiveCell.Offset(4, -18).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Indemnités"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire soumis"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Taux"
ActiveCell.Offset(0, 6).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Total"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 0).Range("A1").Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Jours fériés"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-6]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Vacances"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-8]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -17).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(0, 12).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Indemnités"
ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-4]C+R[-2]C"
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium
ActiveCell.Offset(4, -18).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Charges"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire soumis"
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Taux"
ActiveCell.Offset(0, 6).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Total"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(2, 0).Range("A1").Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "AVS"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-16]C[12]+R[-6]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. chômage"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-18]C[12]+R[-8]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. accident"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-20]C[12]+R[-10]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Prév. professionnelle"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    'ActiveCell.FormulaR1C1 = "=R[-18]C[12]+R[-8]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -16).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ass. maternité"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-24]C[12]+R[-14]C[12]"
ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-12]/100*RC[-8]"
ActiveCell.Offset(2, -17).Range("A1:K1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
ActiveCell.Offset(0, 12).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Charges"
ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-10]C+R[-8]C+R[-6]C+R[-4]C+R[-2]C"
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium


ActiveCell.Offset(4, -18).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Calcul"
ActiveCell.Offset(0, 6).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire brut"
'Données
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Charges"
'données
ActiveCell.Offset(0, 2).Range("A1:C1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 15
    ActiveCell.FormulaR1C1 = "Salaire net"
ActiveCell.Offset(0, -16).Range("A1:S1").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous

ActiveCell.Offset(2, 0).Range("A1").Select
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ""


ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-32]C[8]+R[-22]C[8]"
ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=R[-6]C[4]"
ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlDouble
    Selection.Borders(xlEdgeLeft).Weight = xlThick
    Selection.Borders(xlEdgeTop).LineStyle = xlDouble
    Selection.Borders(xlEdgeTop).Weight = xlThick
    Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
    Selection.Borders(xlEdgeBottom).Weight = xlThick
 ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlDouble
    Selection.Borders(xlEdgeTop).Weight = xlThick
    Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
    Selection.Borders(xlEdgeBottom).Weight = xlThick
    Selection.NumberFormat = "#,##0.00"
    ActiveCell.FormulaR1C1 = "=RC[-8]-RC[-4]"
ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Borders(xlEdgeTop).LineStyle = xlDouble
    Selection.Borders(xlEdgeTop).Weight = xlThick
    Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
    Selection.Borders(xlEdgeBottom).Weight = xlThick
    Selection.Borders(xlEdgeRight).LineStyle = xlDouble
    Selection.Borders(xlEdgeRight).Weight = xlThick
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select

'ActiveCell.Offset(2, -17).Range("A1").Select
 '   Selection.HorizontalAlignment = xlCenter
  '  ActiveCell.FormulaR1C1 = "ELISA GASSERT"
'ActiveCell.Offset(1, 0).Range("A1").Select
 '   Selection.HorizontalAlignment = xlCenter
  '  ActiveCell.FormulaR1C1 = "Rue des Eaux-Vives 59"
'ActiveCell.Offset(1, 0).Range("A1").Select
 '   Selection.HorizontalAlignment = xlCenter
  '  ActiveCell.FormulaR1C1 = "1207 Genève"
'ActiveCell.Offset(1, 0).Range("A1").Select
 '   Selection.HorizontalAlignment = xlCenter
  '  ActiveCell.FormulaR1C1 = "0041 22 786 45 40"
'ActiveCell.Offset(1, 0).Range("A1").Select
   ' Selection.HorizontalAlignment = xlCenter
  '  ActiveCell.FormulaR1C1 = "olivier@elisa-gassert.ch"
    
ActiveCell.Offset(0, 19).Columns("A:A").EntireColumn.Select
    ActiveWindow.SelectedSheets.VPageBreaks.Add Before:=ActiveCell
    
'ActiveCell.Offset(59, 0).Rows("1:1").EntireRow.Select
   ' ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell




End Sub


Sub Mise_en_page_Salaire_Salaire()
    
    
ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 5
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 14
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 5).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 6).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 7).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 9).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 10).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 11).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 12).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 13).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 14).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 15).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 16).Columns("A:A").EntireColumn.ColumnWidth = 1
ActiveCell.Offset(0, 17).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 18).Columns("A:A").EntireColumn.ColumnWidth = 10.71
ActiveCell.Offset(0, 19).Columns("A:A").EntireColumn.ColumnWidth = 0.5
ActiveCell.Offset(0, 20).Columns("A:A").EntireColumn.ColumnWidth = 5
    

End Sub

