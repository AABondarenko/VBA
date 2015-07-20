Attribute VB_Name = "Module1"
Sub numbers()

ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo               '_______ _________

Columns("A:A").Select                                                            '_________
    ActiveWorkbook.Worksheets("____1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("____1").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("____1").Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Range("A1").Copy Range("B1")                                                     '________ ______ ________
Range("A1").Select

Do While ActiveCell.Offset(1, 0) > 0
   ActiveCell.Offset(1, 0).Select
Loop

D = ActiveCell.value - Range("A1").value - 1

Range("B2").Select

'MsgBox D

For i = 0 To D                                                                   '_____ _____ - _______ ____ _ ___ ______
   ActiveCell = ActiveCell.Offset(-1, 0) + 1                                     '_ _______ B ___________ ______
   ActiveCell.Offset(1, 0).Select                                                '______ ______ _______
Next i

Range("B1").Select

Do While ActiveCell > 0                                                          '____ ________ _ _______ B ______ ____
   If ActiveCell.value = ActiveCell.Offset(0, -1) Then                           '____ ________ __ ____ _______ _____
      ActiveCell.Offset(1, 0).Select                                             '_________ ____
   Else
      ActiveCell.Copy ActiveCell.Offset(0, 2)                                    '_____ ________ ________ __ B _ C
      ActiveCell.Offset(0, -1).Select
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove        '_ _________ ______ ______ _ _
      ActiveCell.Offset(1, 1).Select
   End If
Loop

Columns("D:D").Select                                                            '__________ _______ ___ ________ ______ _____
    ActiveWorkbook.Worksheets("____1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("____1").Sort.SortFields.Add Key:=Range("D1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("____1").Sort
        .SetRange Range("D:D")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Selection.Font.Bold = True                                                       '______________
Range("D1").Select

MsgBox ("______! _____________ ______ ___________ _ _______ D.")

End Sub
