Option Explicit

'********************************************************************************
'Objetivo:          Classificação                                               *
'Autor:             Isac Canedo                                                 *
'Data Criação:      10/03/2000                                                  *
'Data Atualização:  10/03/2000                                                  *
'********************************************************************************
Dim classificar As Integer

Dim xline as Range
For Each xline in Selection.Rows
xline.Sort xline.Cells(1), xlAscending,
Header:=xlNo, Orientation:=xlSortRows
Next xline

End Sub
