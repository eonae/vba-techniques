VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////
'/////////////// �����-������ ��� ��������� ���������� ///////////////
'//////////////////////////////////////////////////////////////////////

'// ����������� �� ������� ��� ������ WithEvents ����� ������ ��� ����������� ���� ��������,
'// ������� ���������� �������� ������� ��� �������� ���� ���������.

Private WithEvents pButton As MSForms.CommandButton
Private WithEvents pTextBox As MSForms.TextBox
Attribute pTextBox.VB_VarHelpID = -1
Private WithEvents pComboBox As MSForms.ComboBox
Attribute pComboBox.VB_VarHelpID = -1
Private WithEvents pListBox As MSForms.ListBox
Attribute pListBox.VB_VarHelpID = -1
Private WithEvents pOptionButton As MSForms.OptionButton
Attribute pOptionButton.VB_VarHelpID = -1
Private WithEvents pCheckBox As MSForms.CheckBox
Attribute pCheckBox.VB_VarHelpID = -1

Private pControlType As String '// ��� �������� ����������

'// ����� ����� ������� ������ Controller, ������� ����� ������������ �������. ��� ����� ����� ��
'// ��������� ��������������� ��������� ����, � ���������������� ��� ����� ��������������� ��� ��������
'// ��������� � Controller (� ������, ��� �� "�������������")

Private pController As clsController '// ��� Controller

Private Sub Class_Initialize() '// ����� � �� ������ ����� ����

    Set pButton = Nothing
    Set pTextBox = Nothing
    Set pComboBox = Nothing
    Set pListBox = Nothing
    Set pOptionButton = Nothing
    Set pCheckBox = Nothing
    
    Set pController = Nothing
    
    pControlType = ""
    
End Sub

Public Property Set Controller(ByRef � As clsController)

    If Not pController Is Nothing Then Exit Property '//���������� �������� ����������.
    
    Set pController = �
    
End Property

Public Property Set Control(ByRef C As Object)
    
    If pControlType <> "" Then Exit Property '//�������� ���������������� ���������� ����������.

    pControlType = TypeName(C)
    
'// � ����������� �� ���� ������������� ��������, ���������� ��� � ��������������� ����.
    Select Case pControlType
        Case "CommandButton":
            Set pButton = C
        Case "TextBox":
            Set pTextBox = C
        Case "ComboBox":
            Set pComboBox = C
        Case "ListBox":
            Set pListBox = C
        Case "OptionButton":
            Set pOptionButton = C
        Case "CheckBox":
            Set pCheckBox = C
        Case Else:
            pControlType = ""
            Debug.Print "Wrong type!"
    End Select
    
End Property

Public Property Get Control() As Object

'// ���������� ��������� ������� ����������.

    Select Case pControlType
        Case "CommandButton":
            Set Control = pButton
        Case "TextBox":
            Set Control = pTextBox
        Case "ComboBox":
            Set Control = pComboBox
        Case "ListBox":
            Set Control = pListBox
        Case "OptionButton":
            Set Control = pOptionButton
        Case "CheckBox":
            Set Control = pCheckBox
        Case Else:
            Debug.Print "Wrong type!"
    End Select
    
End Property

'/////////////////////////////////////////////////////////
'/////////////// �������������� ������� //////////////////
'/////////////////////////////////////////////////////////

'// ������ ������������ ���������� ����� ���� �����

Private Sub pButton_Click()
    Call pController.MyEvent1(Me)
End Sub

Private Sub pTextBox_Change()
    Call pController.MyEvent2(Me)
End Sub

Private Sub pComboBox_Change()
    Call pController.MyEvent2(Me)
End Sub

Private Sub pListBox_Change()
    Call pController.MyEvent2(Me)
End Sub


