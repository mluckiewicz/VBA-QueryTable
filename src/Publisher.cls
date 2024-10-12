VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Publisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasa o nazwie "Publisher"
' Ta klasa bêdzie wydawaæ zdarzenie "OnStateChanged"
' i przechowywaæ wartoœæ do œledzenia

Option Explicit

' Deklaracja zdarzenia
Public Event OnStateChanged(ByVal Success As Boolean)
Private WithEvents QT As Excel.QueryTable
Attribute QT.VB_VarHelpID = -1
Private currentState As Boolean ' Prywatna zmienna przechowuj¹ca stan

' Przypisujemy Excel.QueryTable dla Ÿród³a danych
Public Sub SetQueryTable(ByVal qtObject As Excel.QueryTable)
    Set QT = qtObject
End Sub

'Emitowanie zdarzenia po zakoñczeniu odœwie¿ania
Private Sub QT_AfterRefresh(ByVal Success As Boolean)
    currentState = Success
    RaiseEvent OnStateChanged(Success)
    'Debug.Print "Odœwie¿ono " & ObjPtr(Me)
End Sub

' Metoda zwracaj¹ca bie¿¹cy stan
Public Function GetState() As String
    GetState = currentState
End Function

