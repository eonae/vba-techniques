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
'/////////////// Класс объединяющий элементы управления ///////////////
'//////////////////////////////////////////////////////////////////////

Private pElements As Collection '//Коллекция контролов.

Public Sub PassControls(ParamArray Controls() As Variant)
'//Метод, инициализирующий объект Controller. Не работает, если объект уже был инициализирован.
    
    Dim Ctrl As Variant
    Dim i%
    Dim Wrapper As clsWrapper
    
    If Not pElements Is Nothing Then
        Debug.Print "Controls are already passed"
        Exit Sub
    End If
    
    Set pElements = New Collection

    For Each Ctrl In Controls
    
        Set Wrapper = New clsWrapper '// Создаём "обёртку"
        Set Wrapper.Control = Ctrl   '// Заворачиваем элемент управления
        Set Wrapper.Controller = Me  '// Указываем управляющий объект - текущий экземпляр класса clsController
        
        pElements.Add Wrapper '// Добавляем "завёрнутый" элемент в коллекцию
        
        Set Wrapper = Nothing
    Next
    
End Sub

'/////////////////////////////////////////////////////////
'/////////////// Обработчики событий /////////////////////
'/////////////////////////////////////////////////////////

'// Список передаваемых параметров может быть любой

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
