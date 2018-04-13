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

'// Обычно события элементов управления формы обрабатываются соответствующими процедурами
'// непосредственно в модуле формы. У этого подхода есть два недостатка:
'//
'//     * В случае, если есть несколько аналогичных форм, идентичный код обработчкиков
'//     придётся копировать в каждый из них, что ведёт к разрастанию кода и сложности его
'//     поддержки.
'//
'//     * Если на форме присутствуют одинаковые элементы управления, обработчики приходится
'//     писать для всех отдельно. (со теми же последствиями).
'//
'// В предложенном решении мы создаём класс (clsController), в экземпляре которого собираем любое
'// количество элементов управления формы и подписываемся на их события. При этом все
'// обработчики для элементов управления или их групп находятся в описании самого класса.
'//
'// Проблема здесь заключается в том, что использовать ключевое слово WithEvents (т. е. подписаться на событие)
'// напрямую возможно только для отдельного объекта, но не коллекции. Для обхода этого ограничения, каждый элемент
'// управления "заворачивается" внутрь экземпляра класса clsWrapper, который, получая событие от лежащего
'// в нём элемента вызывает соответствующий метод обработки класса clsController.

Private Controller As clsController '// Объявляем наш "контроллер"

Private Sub btn2_Click()
    Debug.Print "Genuine event is working too!"
End Sub

Private Sub UserForm_Initialize()

'// При инициализации формы создаём экземпляр класса и передаём в него конролы, которыми он должен управлять.

    Set Controller = New clsController
    Call Controller.PassControls(Me.btn1, Me.btn2, Me.btn3, Me.TextBox1)
    
End Sub


