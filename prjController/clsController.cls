VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////
'/////////////// ����� ������������ �������� ���������� ///////////////
'//////////////////////////////////////////////////////////////////////

Private pElements As Collection '//��������� ���������.

Public Sub PassControls(ParamArray Controls() As Variant)
'//�����, ���������������� ������ Controller. �� ��������, ���� ������ ��� ��� ���������������.
    
    Dim Ctrl As Variant
    Dim i%
    Dim Wrapper As clsWrapper
    
    If Not pElements Is Nothing Then
        Debug.Print "Controls are already passed"
        Exit Sub
    End If
    
    Set pElements = New Collection

    For Each Ctrl In Controls
    
        Set Wrapper = New clsWrapper '// ������ "������"
        Set Wrapper.Control = Ctrl   '// ������������ ������� ����������
        Set Wrapper.Controller = Me  '// ��������� ����������� ������ - ������� ��������� ������ clsController
        
        pElements.Add Wrapper '// ��������� "���������" ������� � ���������
        
        Set Wrapper = Nothing
    Next
    
End Sub

'/////////////////////////////////////////////////////////
'/////////////// ����������� ������� /////////////////////
'/////////////////////////////////////////////////////////

'// ������ ������������ ���������� ����� ���� �����

Public Sub MyEvent1(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent1!")
End Sub

Public Sub MyEvent2(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent2!")
End Sub

Public Sub MyEvent3(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent3!")
End Sub

Public Sub MyEvent4(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent4!")
End Sub