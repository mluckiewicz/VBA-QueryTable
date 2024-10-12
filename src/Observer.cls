VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Observer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasa Observer - mo¿e odbieraæ powiadomienia od wielu Publisherów
Option Explicit

' Kolekcja obiektów PublisherEventHandler, które nas³uchuj¹ zdarzeñ od Publisherów
Private publisherHandlers As Collection
Private stateChangeCount As Long

' Inicjalizacja klasy
Private Sub Class_Initialize()
    Set publisherHandlers = New Collection
    stateChangeCount = 0
End Sub

' Rejestracja nowego Publishera do obserwowania
Public Sub RegisterPublisher(ByRef publisher As publisher)
    ' Utworzenie nowego obiektu PublisherEventHandler dla tego Publishera
    Dim handler As PublisherEventHandler
    Set handler = New PublisherEventHandler
    handler.Init publisher, Me ' Przypisanie Publishera i siebie jako Observera

    ' Dodanie handlera do kolekcji
    publisherHandlers.Add handler
    
    Debug.Print "Zarejestrowano publisher:  " & ObjPtr(publisher) & " w obserwatorze " & ObjPtr(Me)
    
End Sub

' Metoda obs³uguj¹ca zmiany stanu od dowolnego Publishera
Public Sub HandleStateChange(ByRef publisher As publisher, ByVal Success As Boolean)
    stateChangeCount = stateChangeCount + 1
    Debug.Print "Obserwer " & ObjPtr(Me) & " otrzyma³ powiadomienie z obiektu Publisher " & ObjPtr(publisher) & ": " & Success
    Debug.Print "Odœwie¿ono stan : " & stateChangeCount & " z " & publisherHandlers.Count & " Publisherów"
    
    If stateChangeCount = publisherHandlers.Count Then
        Debug.Print "Zmieniono stan wszystkich Publisherów: " & stateChangeCount & " z " & publisherHandlers.Count
    End If
    
End Sub

