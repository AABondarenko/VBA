Attribute VB_Name = "Module1"
Sub find_and_shift()

    For Each cell In Range("H1:H2500")
        If cell.Value Like "*.*" Then
            cell.Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    Next cell
          
    


End Sub
