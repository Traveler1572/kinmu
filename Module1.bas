Attribute VB_Name = "Module1"
Option Explicit

Public Sub work_input_op()          '2022-06-23 �V�K�쐬

    '�������͂̂��߂̃f�[�^�^�ϐ���錾
    Dim myDateA As Date
    Dim myDateB As Date
    Dim myDateC As Date
    
    myDateA = TimeValue("8:40:00")       '�n�Ǝ���
    myDateB = TimeValue("17:10:00")      '�I�Ǝ���
    myDateC = TimeValue("1:00:00")       '�x�e����
    
    '�����\���`����ݒ�i�����j
    With Range("C21:C25")                 '�n�Ǝ��ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateA
        
        With Range("C28:C31")
            .NumberFormat = "hh:mm"
            .Value = myDateA
        End With
        
    End With
    
    With Range("D21:D25")                  '�I�Ǝ��ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateB
        
        With Range("D28:D31")
            .NumberFormat = "hh:mm"
            .Value = myDateB
        End With
        
    End With
    
    With Range("E21:E25")                  '�x�e���ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateC
        
        With Range("E28:E31")
            .NumberFormat = "hh:mm"
            .Value = myDateC
        End With
        
    End With
    
End Sub
