Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim boo As Boolean

    boo = False


    If boo = True Then

        Dim num, unit_P As Long
        Dim cell_Adr As String

        cell_Adr = Target.Address(rowabsolute:=False, columnabsolute:=False)

        CellTextCopy (ActiveCell.Value)

    End If
End Sub
