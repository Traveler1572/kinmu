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
    
    '時刻表示形式を設定（時分）
    With Range("O7:O11")                 '始業時間用
        .NumberFormat = "hh:mm"
        .Value = myDateA
    End With
    
    With Range("P7:P11")                  '終業時間用
        .NumberFormat = "hh:mm"
        .Value = myDateB
    End With
    
    With Range("Q7:Q11")                  '休憩時間用
        .NumberFormat = "hh:mm"
        .Value = myDateC
    End With
    
End Sub
