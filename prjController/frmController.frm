VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmController 
   Caption         =   "UserForm1"
   ClientHeight    =   3768
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3324
   OleObjectBlob   =   "frmController.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// ������ ������� ��������� ���������� ����� �������������� ���������������� �����������
'// ��������������� � ������ �����. � ����� ������� ���� ��� ����������:
'//
'//     * � ������, ���� ���� ��������� ����������� ����, ���������� ��� �������������
'//     ������� ���������� � ������ �� ���, ��� ���� � ����������� ���� � ��������� ���
'//     ���������.
'//
'//     * ���� �� ����� ������������ ���������� �������� ����������, ����������� ����������
'//     ������ ��� ���� ��������. (�� ���� �� �������������).
'//
'// � ������������ ������� �� ������ ����� (clsController), � ���������� �������� �������� �����
'// ���������� ��������� ���������� ����� � ������������� �� �� �������. ��� ���� ���
'// ����������� ��� ��������� ���������� ��� �� ����� ��������� � �������� ������ ������.
'//
'// �������� ����� ����������� � ���, ��� ������������ �������� ����� WithEvents (�. �. ����������� �� �������)
'// �������� �������� ������ ��� ���������� �������, �� �� ���������. ��� ������ ����� �����������, ������ �������
'// ���������� "��������������" ������ ���������� ������ clsWrapper, �������, ������� ������� �� ��������
'// � �� �������� �������� ��������������� ����� ��������� ������ clsController.

Private Controller As clsController '// ��������� ��� "����������"

Private Sub btn2_Click()
    Debug.Print "Genuine event is working too!"
End Sub

Private Sub UserForm_Initialize()

'// ��� ������������� ����� ������ ��������� ������ � ������� � ���� �������, �������� �� ������ ���������.

    Set Controller = New clsController
    Call Controller.PassControls(Me.btn1, Me.btn2, Me.btn3, Me.TextBox1)
    
End Sub


