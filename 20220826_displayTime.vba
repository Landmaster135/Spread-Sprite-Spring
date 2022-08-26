'WindowsAPI の timeGetTime 関数（システム時刻をミリ秒単位で取得）を宣言
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Function GetTime() As Double
 
    Dim TStart As Long
    Dim TEnd As Long
    
    '処理開始時間の取得
    TStart = timeGetTime
    
    'TODO
    Sleep 1000
    
    
    '処理開終了間の取得
    TEnd = timeGetTime
    
    '処理経過時間を表示
    GetTime = (TEnd - TStart) / 1000
 
End Function

Sub DisplayTime()

    Dim Time As Long
    Time = GetTime()
    MsgBox "処理時間：" & Time & " 秒"
    
End Sub
