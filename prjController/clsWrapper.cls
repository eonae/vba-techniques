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
'/////////////// Класс-обёртка для элементов управления ///////////////
'//////////////////////////////////////////////////////////////////////

'// Подписаться на события при помощи WithEvents можно только для конкретного типа объектов,
'// поэтому приходится объявить заранее все основные типы элементов.

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

Private pControlType As String '// Тип элемента управления

'// Очень важно указать объект Controller, который будет обрабатывать события. Для этого здесь мы
'// объявляем соответствующее приватное поле, а инициализировать его будем непосредственно при передаче
'// контролов в Controller (а точнее, при их "заворачивании")

Private pController As clsController '// Наш Controller

Private Sub Class_Initialize() '// Можно и не делать этого явно

    Set pButton = Nothing
    Set pTextBox = Nothing
    Set pComboBox = Nothing
    Set pListBox = Nothing
    Set pOptionButton = Nothing
    Set pCheckBox = Nothing
    
    Set pController = Nothing
    
    pControlType = ""
    
End Sub

Public Property Set Controller(ByRef С As clsController)

    If Not pController Is Nothing Then Exit Property '//Переписать значение невозможно.
    
    Set pController = С
    
End Property

Public Property Set Control(ByRef C As Object)
    
    If pControlType <> "" Then Exit Property '//Повторно инициализировать переменную невозможно.

    pControlType = TypeName(C)
    
'// В зависимости от типа передаваемого элемента, записываем его в соответствующее поле.
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

'// Возвращаем вложенный элемент управления.

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
'/////////////// Обрабатываемые события //////////////////
'/////////////////////////////////////////////////////////

'// Список передаваемых параметром может быть любой

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


