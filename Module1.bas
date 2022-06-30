Attribute VB_Name = "Module1"
Option Explicit

Public Sub work_input_op()          '2022-06-23 新規作成

    '時刻入力のためのデータ型変数を宣言
    Dim myDateA As Date
    Dim myDateB As Date
    Dim myDateC As Date
    
    myDateA = TimeValue("8:45:00")       '始業時間
    myDateB = TimeValue("17:15:00")      '終業時間
    myDateC = TimeValue("1:00:00")       '休憩時間
    
    '条件パラメータが10以上の場合、一括処理を実行
    If Range("O2") >= 10 Then
    
    '時刻表示形式を設定（時分）
    With Range("C21:C25")                 '始業時間用
        .NumberFormat = "hh:mm"
        .Value = myDateA
        
        With Range("C28:C31")
            .NumberFormat = "hh:mm"
            .Value = myDateA
        End With
        
    End With
    
    With Range("D21:D25")                  '終業時間用
        .NumberFormat = "hh:mm"
        .Value = myDateB
        
        With Range("D28:D31")
            .NumberFormat = "hh:mm"
            .Value = myDateB
        End With
        
    End With
    
    With Range("E21:E25")                  '休憩時間用
        .NumberFormat = "hh:mm"
        .Value = myDateC
        
        With Range("E28:E31")
            .NumberFormat = "hh:mm"
            .Value = myDateC
        End With
        
    End With
    
    Else
        MsgBox "処理は実行せずに終了します。"
    End If
    
End Sub
