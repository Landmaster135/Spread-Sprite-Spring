Public Sub PasteToInDesign()

    Dim i As Integer: Dim j As Integer
    Dim k As Integer: Dim l As Integer: Dim m As Integer
    Const a As String = "建築用途": Const b As String = "所在地"
    Const c As String = "部材": Const d As String = "用いられた技術"
    Const e As String = "繊維の形状": Const f As String = "用いられた材料"
    Const g As String = "建設年代": Const h As String = "施工"
    Const v As String = "補強年代": Const w As String = "施工面積 [m2]"
    Const x As String = "掲載誌"
    'Const y As String = "補強概要"

    Dim at As String: Dim bt As String
    Dim ct As String: Dim dt As String
    Dim et As String: Dim ft As String
    Dim gt As String: Dim ht As String
    Dim vt As String: Dim wt As String
    Dim xt As String: Dim yt As String


    Dim nr As Integer: Dim nc As Integer
    Const katamari As Integer = 7

    nc = 2
    Const r As Integer = 133 '最後の事例の行インデックス 133

    Dim Range1 As Range

    If ActiveSheet.Name = "InDesign貼る" Then

        'at = Worksheets(1).Cells(1, 3).Value
        'bt = Worksheets(1).Cells(1, 35).Value
        'ct = Worksheets(1).Cells(1, 5).Value
        'dt = Worksheets(1).Cells(1, 7).Value
        'et = Worksheets(1).Cells(1, 9).Value
        'ft = Worksheets(1).Cells(1, 8).Value
        'gt = Worksheets(1).Cells(1, 12).Value
        'ht = Worksheets(1).Cells(1, 10).Value
        'vt = Worksheets(1).Cells(1, 13).Value
        'wt = Worksheets(1).Cells(1, 27).Value
        'xt = Worksheets(1).Cells(1, 3).Value
        'yt = Worksheets(1).Cells(1, 36).Value

        For i = 2 To r
            at = Worksheets(1).Cells(i, 3).Value
            bt = Worksheets(1).Cells(i, 35).Value
            ct = Worksheets(1).Cells(i, 5).Value
            dt = Worksheets(1).Cells(i, 7).Value
            et = Worksheets(1).Cells(i, 9).Value
            ft = Worksheets(1).Cells(i, 8).Value
            gt = Worksheets(1).Cells(i, 12).Value
            ht = Worksheets(1).Cells(i, 11).Value
            vt = Worksheets(1).Cells(i, 13).Value
            wt = Worksheets(1).Cells(i, 27).Value
            For j = 0 To 2
                If Worksheets(1).Cells(i, 19 + j * 2).Value = "" Then
                    Exit For
                Else
                    If j = 0 Then
                        If Worksheets(1).Cells(i, 19 + j * 2).Value = "セメント・コンクリート" Then
                            xt = Worksheets(1).Cells(i, 19 + j * 2).Value & " ： No. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
                        Else 'Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリートテクノ" Or Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリート工学"
                            xt = Worksheets(1).Cells(i, 19 + j * 2).Value & " ： " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
                        End If
                    Else
                        If Worksheets(1).Cells(i, 19 + j * 2).Value = "セメント・コンクリート" Then
                            xt = xt & vbLf & Worksheets(1).Cells(i, 19 + j * 2).Value & " ： No. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
                        Else 'Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリートテクノ" Or Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリート工学"
                            xt = xt & vbLf & Worksheets(1).Cells(i, 19 + j * 2).Value & " ： Vol. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
                        End If
                    End If
                End If
            Next j
            'yt = Worksheets(1).Cells(i, 36).Value

            ActiveSheet.Cells(i * katamari, nc - 1).Value = i - 1
            ActiveSheet.Cells(i * katamari, nc - 1).Interior.Color = RGB(255, 255, 0)

            ActiveSheet.Cells(i * katamari, nc).Value = a
            ActiveSheet.Cells(i * katamari, nc + 1).Value = at
            ActiveSheet.Cells(i * katamari, nc + 2).Value = b
            ActiveSheet.Cells(i * katamari, nc + 3).Value = bt
            ActiveSheet.Cells(i * katamari + 1, nc).Value = c
            ActiveSheet.Cells(i * katamari + 1, nc + 1).Value = ct
            ActiveSheet.Cells(i * katamari + 1, nc + 2).Value = d
            ActiveSheet.Cells(i * katamari + 1, nc + 3).Value = dt
            ActiveSheet.Cells(i * katamari + 2, nc).Value = e
            ActiveSheet.Cells(i * katamari + 2, nc + 1).Value = et
            ActiveSheet.Cells(i * katamari + 2, nc + 2).Value = f
            ActiveSheet.Cells(i * katamari + 2, nc + 3).Value = ft
            ActiveSheet.Cells(i * katamari + 3, nc).Value = g
            ActiveSheet.Cells(i * katamari + 3, nc + 1).Value = gt
            ActiveSheet.Cells(i * katamari + 3, nc + 2).Value = h
            ActiveSheet.Cells(i * katamari + 3, nc + 3).Value = ht
            ActiveSheet.Cells(i * katamari + 4, nc).Value = v
            ActiveSheet.Cells(i * katamari + 4, nc + 1).Value = vt
            ActiveSheet.Cells(i * katamari + 4, nc + 2).Value = w
            ActiveSheet.Cells(i * katamari + 4, nc + 3).Value = wt
            ActiveSheet.Cells(i * katamari + 5, nc).Value = x
            ActiveSheet.Cells(i * katamari + 5, nc - 1).Interior.Color = RGB(255, 130, 0)
            ActiveSheet.Cells(i * katamari + 6, nc).Value = xt
            Set Range1 = Range(Cells(i * katamari + 6, nc), Cells(i * katamari + 6, nc))
            Range1.HorizontalAlignment = xlLeft
            Range1.VerticalAlignment = xlTop

            'ActiveSheet.Cells(i * katamari + 7, nc).Value = y
            'ActiveSheet.Cells(i * katamari + 8, nc).Value = yt
        Next i

        Dim bord As Borders
        For k = 2 To i
            For l = 0 To katamari - 1
                For m = 0 To 3
                    Set bord = Range(Cells(k * katamari + l, nc + m), Cells(k * katamari + l, nc + m)).Borders
                    bord.Weight = xlThin
                    bord.LineStyle = xlContinuous
                Next m
            Next l
            Set bord = Nothing
        Next k

    End If

End Sub

Public Function VolNo(ByVal i As String, ByVal j As String) As String

    Dim ArticleName As String
    Dim xtttt As String

    ArticleName = Worksheets(1).Cells(i, 19 + j * 2).Value

    If j = 0 Then
        If Worksheets(1).Cells(i, 19 + j * 2).Value = "セメント・コンクリート" Then
            VolNo = Worksheets(1).Cells(i, 19 + j * 2).Value & " ： No. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
        Else 'Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリートテクノ" Or Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリート工学"
            VolNo = Worksheets(1).Cells(i, 19 + j * 2).Value & " ： " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
        End If
    Else
        If Worksheets(1).Cells(i, 19 + j * 2).Value = "セメント・コンクリート" Then
            VolNo = VolNo & vbLf & Worksheets(1).Cells(i, 19 + j * 2).Value & " ： No. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
        Else 'Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリートテクノ" Or Worksheets(1).Cells(i, 19 + j * 2).Value = "コンクリート工学"
            VolNo = VolNo & vbLf & Worksheets(1).Cells(i, 19 + j * 2).Value & " ： Vol. " & Worksheets(1).Cells(i, 19 + j * 2 + 1).Value
        End If
    End If


End Function

Public Sub CellTextCopy(text As String)

    'C:\Program Files (x86)\Microsoft Office\root\vfs\SystemX86\FM20.DLL　を参照

    Dim buf As String

    buf = text

    With New MSForms.DataObject
        .SetText buf      '変数の値をDataObjectに格納する
        .PutInClipboard   'DataObjectのデータをクリップボードに格納する
    End With

End Sub
