Attribute VB_Name = "Salaire_Transfert"
Sub Transfert_Salaires__Données_à_Fiches_Janvier()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("C8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("I8")
Range("Q8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("G13")
Range("G13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("K15")
Range("K15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("O15")
Range("O15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("K23")
Range("K23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("K25")
Range("K25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("K33")
Range("K33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("K35")
Range("K35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("K37")
Range("K37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("S39")
Range("S39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("K41")
Range("K41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Février()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("X8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("AD8")
Range("AJ8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("AB13")
Range("AB13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("AF15")
Range("AF15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("AJ15")
Range("AJ15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("AF23")
Range("AF23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("AF25")
Range("AF25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("AF33")
Range("AF33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("AF35")
Range("AF35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("AF37")
Range("AF37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("AN39")
Range("AN39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("AF41")
Range("AF41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Mars()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("AS8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("AY8")
Range("BE8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("AW13")
Range("AW13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("BA15")
Range("BA15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("BE15")
Range("BE15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("BA23")
Range("BA23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("BA25")
Range("BA25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("BA33")
Range("BA33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("BA35")
Range("BA35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("BA37")
Range("BA37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("BI39")
Range("BI39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("BA41")
Range("BA41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Avril()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("BN8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("BT8")
Range("BZ8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("BR13")
Range("BR13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("BV15")
Range("BV15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("BZ15")
Range("BZ15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("BV23")
Range("BV23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("BV25")
Range("BV25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("BV33")
Range("BV33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("BV35")
Range("BV35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("BV37")
Range("BV37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("CD39")
Range("CD39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("BV41")
Range("BV41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Mai()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("CI8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("CO8")
Range("CU8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("CM13")
Range("CM13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous5
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("CQ15")
Range("CQ15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("CU15")
Range("CU15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("CQ23")
Range("CQ23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("CQ25")
Range("CQ25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("CQ33")
Range("CQ33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("CQ35")
Range("CQ35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("CQ37")
Range("CQ37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("CY39")
Range("CY39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("CQ41")
Range("CQ41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Juin()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("DD8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("DJ8")
Range("DP8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("DH13")
Range("DH13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("DL15")
Range("DL15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("DP15")
Range("DP15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("DL23")
Range("DL23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("DL25")
Range("DL25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("DL33")
Range("DL33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("DL35")
Range("DL35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("DL37")
Range("DL37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("DT39")
Range("DT39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("DL41")
Range("DL41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Juillet()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("DY8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("EE8")
Range("EK8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("EC13")
Range("EC13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("EG15")
Range("EG15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("EK15")
Range("EK15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("EG23")
Range("EG23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("EG25")
Range("EG25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("EG33")
Range("EG33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("EG35")
Range("EG35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("EG37")
Range("EG37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("EO39")
Range("EO39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("EG41")
Range("EG41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Août()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("ET8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("EZ8")
Range("FF8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("EX13")
Range("EX13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("FB15")
Range("FB15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("FF15")
Range("FF15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("FB23")
Range("FB23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("FB25")
Range("FB25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("FB33")
Range("FB33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("FB35")
Range("FB35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("FB37")
Range("FB37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("FJ39")
Range("FJ39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("FB41")
Range("FB41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Septembre()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("FO8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("FU8")
Range("GA8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("FS13")
Range("FS13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("FW15")
Range("FW15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("GA15")
Range("GA15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("FW23")
Range("FW23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("FW25")
Range("FW25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("FW33")
Range("FW33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("FW35")
Range("FW35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("FW37")
Range("FW37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("GE39")
Range("GE39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("FW41")
Range("FW41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Octobre()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("GJ8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("GP8")
Range("GV8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("GN13")
Range("GN13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("GR15")
Range("GR15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("GV15")
Range("GV15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("GR23")
Range("GR23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("GR25")
Range("GR25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("GR33")
Range("GR33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("GR35")
Range("GR35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("GR37")
Range("GR37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("GZ39")
Range("GZ39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("GR41")
Range("GR41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Novembre()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("HE8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("HK8")
Range("HQ8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("HI13")
Range("HI13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("HM15")
Range("HM15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("HQ15")
Range("HQ15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("HM23")
Range("HM23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("HM25")
Range("HM25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("HM33")
Range("HM33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("HM35")
Range("HM35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("HM37")
Range("HM37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("HU39")
Range("HU39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("HM41")
Range("HM41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub


Sub Transfert_Salaires__Données_à_Fiches_Décembre()


Sheets("Salaire").Select
Sheets("Donnée").Range("C8").Copy Sheets("Salaire").Range("HZ8")
Sheets("Donnée").Range("I8").Copy Sheets("Salaire").Range("IF8")
Range("IL8").Select
    Selection.NumberFormat = "mmmm yyyy"
    ActiveCell.FormulaR1C1 = Date
Sheets("Donnée").Range("I41").Copy Sheets("Salaire").Range("ID13")
Range("ID13").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I43").Copy Sheets("Salaire").Range("IH15")
Range("IH15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I45").Copy Sheets("Salaire").Range("IL15")
Range("IL15").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I52").Copy Sheets("Salaire").Range("IH23")
Range("IH23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I50").Copy Sheets("Salaire").Range("IH25")
Range("IH25").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I57").Copy Sheets("Salaire").Range("IH33")
Range("IH33").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I59").Copy Sheets("Salaire").Range("IH35")
Range("IH35").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I61").Copy Sheets("Salaire").Range("IH37")
Range("IH37").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Donnée").Range("I63").Copy Sheets("Salaire").Range("IP39")
Range("IP39").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.NumberFormat = "#,##0.00"
Sheets("Donnée").Range("I65").Copy Sheets("Salaire").Range("IH41")
Range("IH41").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous


End Sub

