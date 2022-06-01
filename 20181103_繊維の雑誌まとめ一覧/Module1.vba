'Dim RowNumber As Integer
Const RowNumber As Integer = 218 '120
Const VolNum As Integer = 8 '「号」の列の文字列
Const VolDate As Integer = 7 '「年」の列の文字列
Const MaxArticleNumber As Integer = 19

Public Sub DecideRowNumber()
    Dim i As Integer
    While ActiveSheet.Cells(i, VolNum).Value Like "*No.*"
        RowNumber = i
        i = i + 1
    Wend

End Sub

Public Sub ArrowSeperate()
    Dim i As Integer

    RowNumber = 120

    For i = 1 To RowNumber
        Select Case Left(ActiveSheet.Cells(i, VolNum).Value, 1)
            Case "←" 'キープ
                ActiveSheet.Cells(1, i).Value = "True"
                ActiveSheet.Cells(VolNum, i).Value = Replace(ActiveSheet.Cells(VolNum, i), "←", "")
            Case "↓" '記録済み
                ActiveSheet.Cells(2, i).Value = "True"
                ActiveSheet.Cells(VolNum, i).Value = Replace(ActiveSheet.Cells(VolNum, i), "↓", "")
            Case "↑" '予約済み
                ActiveSheet.Cells(3, i).Value = "True"
                ActiveSheet.Cells(VolNum, i).Value = Replace(ActiveSheet.Cells(VolNum, i), "↑", "")
        End Select

    Next i

End Sub
Public Sub DescriptionAboutFiber() '「繊維」、「FR」、「Fiber」の記述なしだったらn
    Dim i As Integer
    For i = 0 To RowNumber
        If Right(ActiveSheet.Cells(VolNum, i).Value, 1) = "n" Then
            ActiveSheet.Cells(1, i).Value = "n"
            ActiveSheet.Cells(VolNum, i).Value = Replace(ActiveSheet.Cells(VolNum, i), "n", "")
        End If
    Next i

End Sub
Public Sub EnumeratePerArticle()
    Dim i As Integer: Dim j As Integer: Dim k As Integer
    Dim ArticleNumber As Integer



    For i = 0 To RowNumber
        ArticleNumber = 0
        For j = 0 To MaxArticleNumber
            If ActiveSheet.Cells(8 + j * 3, i).Value = "" Then ' Iの列
                Exit For
            Else
                ArticleNumber = ArticleNumber + 1
            End If
        Next j
        ActiveSheet.Rows.Add (j)
        For j = 0 To ArticleNumber
            For k = 0 To 2
                Range(ActiveSheet.Cells(8 + j, i), ActiveSheet.Cells(8 + j + k, i)).Cut
            Next k
        Next j
    Next i

End Sub

Public Sub RewriteVolumeText()

    For i = 1 To RowNumber
        ActiveSheet.Cells(i, VolNum).Value = Replace(ActiveSheet.Cells(i, VolNum).Value, "(", "_")
        ActiveSheet.Cells(i, VolNum).Value = Replace(ActiveSheet.Cells(i, VolNum).Value, ")", "")
    Next i
End Sub

Public Sub InspectVolumeTail()

    Dim VolTail As Variant

    For i = 1 To RowNumber
        VolTail = Right(ActiveSheet.Cells(i, VolNum).Value, 2)
        If InStr(VolTail, "_") > 0 Then
            ActiveSheet.Cells(i, VolNum).Value = Left(ActiveSheet.Cells(i, VolNum).Value, 3) & "0" & Right(VolTail, 1) '数字を1桁から2桁にする

        End If
    Next i
End Sub

Public Sub TallyingSum()

    Dim h As Integer: Dim i As Integer: Dim j As Integer

    Dim SumArticle As Integer '記事の数
    Dim SARow As Integer
    Const SACol1 As Integer = 12
    Const SACol2 As String = "L"


    Dim SumRow As Integer
    Dim aRow As Integer
    Dim bRow As Integer
    Dim cRow As Integer
    Dim dRow As Integer
    Dim eRow As Integer
    Dim fRow As Integer

    Dim SumCarbon As Integer
    Dim CarbonCol1 As Integer
    Dim CarbonCol2 As String

    Dim SumAll As Integer

    If ActiveSheet.Name = "雑誌の号数と年月の照合" Then

            For i = 1 To 300
                If ActiveSheet.Cells(i, SACol1).Value = "記事の数" Then
                    SARow = i + 2
                    Exit For
                End If
            Next i
        For h = 0 To 13
            For i = 1 To 4
                For j = 1 To 260
                    If Worksheets(i).Cells(j, SACol1).Value = "記事の数" Then
                        SumRow = j + 2
                        SumAll = SumAll + Worksheets(i).Cells(SumRow, SACol1 + h).Value
                        Exit For
                    End If
                Next j
            Next i

        'Select Case h
        '    Case 1
                ActiveSheet.Cells(SARow, SACol1 + h).Value = SumAll
                SumAll = 0
        'End Select

        Next h
    Else

    End If

End Sub

Public Function CountColorText(Rng As Range) As Long
    Dim myRng As Range
    Dim count As Long

    Application.Volatile
    count = 0

    For Each myRng In Rng
        If myRng.Font.ColorIndex <> 1 Then
            count = count + 1
        End If
    Next myRng
    CountColorText = count

End Function

Public Function CountAllData(SheetIndex1 As Long, SheetIndex2 As Long) As Long
    Dim count As Long
    Dim BaseRow As Long
    Const BaseCol As Long = 12 'Lの列
    Dim myRow As Long
    myRow = ActiveCell.Row
    Dim myCol As Long
    myCol = ActiveCell.Column
    Dim DeltaRow As Long
    Dim DeltaCol As Long
    Dim i As Long: Dim j As Long

    Application.Volatile
    count = 0

    For i = 1 To 300 '集計用シート
        If ActiveSheet.Cells(i, BaseCol).Value = "記事数の遷移" Then
            BaseRow = i
            Exit For
        End If
    Next i
    i = 0

    DeltaRow = myRow - BaseRow
    DeltaCol = myCol - BaseCol

    For i = SheetIndex1 To SheetIndex2
        For j = 1 To 300 'データシート
            If Worksheets(i).Cells(j, BaseCol).Value = "記事数の遷移" Then
                count = count + Worksheets(i).Cells(j + DeltaRow, BaseCol + DeltaCol).Value
                Exit For
            End If
        Next j
    Next i
    CountAllData = count

End Function
