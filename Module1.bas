Attribute VB_Name = "Module1"
Option Explicit

Public Sub work_input_op()          '2022-06-23 �V�K�쐬

    '�������͂̂��߂̃f�[�^�^�ϐ���錾
    Dim myDateA As Date
    Dim myDateB As Date
    Dim myDateC As Date
    
    myDateA = TimeValue("8:45:00")       '�n�Ǝ���
    myDateB = TimeValue("17:15:00")      '�I�Ǝ���
    myDateC = TimeValue("1:00:00")       '�x�e����
    
    '�����\���`����ݒ�i�����j
    With Range("O7:O11")                 '�n�Ǝ��ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateA
    End With
    
    With Range("P7:P11")                  '�I�Ǝ��ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateB
    End With
    
    With Range("Q7:Q11")                  '�x�e���ԗp
        .NumberFormat = "hh:mm"
        .Value = myDateC
    End With
    
End Sub
