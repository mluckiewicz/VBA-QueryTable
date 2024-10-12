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

' S³ownik przechowuj¹cy handlerów i stany Publisherów
Private publishers As Object

' Inicjalizacja klasy
Private Sub Class_Initialize()
    ' Tworzymy s³ownik dla przechowywania handlerów i stanów
    Set publishers = CreateObject("Scripting.Dictionary")
End Sub

' Rejestracja nowego Publishera do obserwowania
' Publisher: Instancja klasy Publisher, któr¹ chcemy zarejestrowaæ
Public Sub RegisterPublisher(ByRef Publisher As Publisher)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))

    ' SprawdŸ, czy Publisher jest ju¿ zarejestrowany
    If Not IsPublisherRegistered(publisherId) Then
        ' Zainicjuj i dodaj nowego Publishera
        Dim handler As PublisherEventHandler
        Set handler = CreateNewHandler(Publisher)
        
        RegisterNewPublisher publisherId, handler
        LogMessage "Zarejestrowano publisher: " & publisherId & " w obserwatorze " & ObjPtr(Me)
    Else
        LogMessage "Publisher o ID " & publisherId & " ju¿ jest zarejestrowany."
    End If
End Sub

' Sprawdza, czy Publisher o danym Id jest ju¿ zarejestrowany
' publisherId: Id Publishera, który chcemy sprawdziæ
' Zwraca True, jeœli Publisher jest zarejestrowany, w przeciwnym razie False
Private Function IsPublisherRegistered(ByVal publisherId As String) As Boolean
    IsPublisherRegistered = publishers.Exists(publisherId)
End Function

' Tworzy nowy obiekt PublisherEventHandler i przypisuje go do Publishera
' Publisher: Instancja klasy Publisher, dla której tworzymy handler
' Zwraca nowy obiekt PublisherEventHandler
Private Function CreateNewHandler(ByRef Publisher As Publisher) As PublisherEventHandler
    Dim handler As PublisherEventHandler
    Set handler = New PublisherEventHandler
    handler.Init Publisher, Me ' Inicjalizacja handlera
    Set CreateNewHandler = handler
End Function

' Rejestruje nowego Publishera w g³ównym s³owniku
' publisherId: Id Publishera, który chcemy zarejestrowaæ
' handler: Obiekt PublisherEventHandler zwi¹zany z tym Publisherem
Private Sub RegisterNewPublisher(ByVal publisherId As String, ByRef handler As PublisherEventHandler)
    ' Tworzymy nowy s³ownik dla danego Publishera
    Dim publisherData As Object
    Set publisherData = CreateObject("Scripting.Dictionary")
    
    ' Przypisujemy handler i inicjalizujemy stan
    publisherData.Add "handler", handler
    publisherData.Add "state", False ' Pocz¹tkowy stan
    
    ' Dodajemy do g³ównego s³ownika publishers
    publishers.Add publisherId, publisherData
End Sub

' Wyrejestrowanie Publishera (opcjonalna metoda)
' Publisher: Instancja klasy Publisher, któr¹ chcemy wyrejestrowaæ
Public Sub UnregisterPublisher(ByRef Publisher As Publisher)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))
    
    If publishers.Exists(publisherId) Then
        publishers.Remove publisherId
        LogMessage "Wyrejestrowano publisher: " & publisherId
    Else
        LogMessage "Publisher o ID " & publisherId & " nie istnieje."
    End If
End Sub

' Metoda obs³uguj¹ca zmiany stanu od dowolnego Publishera
' Publisher: Instancja klasy Publisher, która zg³asza zmianê stanu
' Success: Zmienna logiczna informuj¹ca, czy zmiana stanu by³a udana
Public Sub HandleStateChange(ByRef Publisher As Publisher, ByVal Success As Boolean)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))

    ' Sprawdzanie i aktualizacja stanu danego Publishera
    If publishers.Exists(publisherId) Then
        UpdatePublisherState publisherId, Success
        LogMessage "Obserwer " & ObjPtr(Me) & " otrzyma³ powiadomienie z obiektu Publisher " & ObjPtr(Publisher) & ": " & Success

        ' Sprawdzanie, czy wszystkie Publishery s¹ zaktualizowane
        If AllPublishersUpdated() Then
            HandleAllPublishersUpdated
        End If
    End If
End Sub

' Aktualizuje stan dla danego Publishera
' publisherId: Id Publishera, którego stan chcemy zaktualizowaæ
' Success: Zmienna logiczna informuj¹ca o nowym stanie Publishera
Private Sub UpdatePublisherState(ByVal publisherId As String, ByVal Success As Boolean)
    publishers(publisherId)("state") = Success
End Sub

' Sprawdza, czy wszystkie Publishery zosta³y zaktualizowane
' Zwraca True, jeœli wszystkie Publishery s¹ zaktualizowane, w przeciwnym razie False
Private Function AllPublishersUpdated() As Boolean
    Dim key As Variant
    For Each key In publishers.Keys
        If publishers(key)("state") = False Then
            AllPublishersUpdated = False
            Exit Function
        End If
    Next key
    AllPublishersUpdated = True
End Function

' Dzia³ania wykonywane, gdy wszystkie Publishery s¹ zaktualizowane
Private Sub HandleAllPublishersUpdated()
    LogMessage "Zmieniono stan wszystkich Publisherów."
    ResetPublisherStates
End Sub

' Resetowanie stanów wszystkich Publisherów
Private Sub ResetPublisherStates()
    Dim key As Variant
    For Each key In publishers.Keys
        publishers(key)("state") = False ' Resetujemy stan
    Next key
    LogMessage "Stany Publisherów zresetowane."
End Sub

' Uniwersalna metoda do logowania wiadomoœci
' message: Treœæ wiadomoœci do zalogowania
Private Sub LogMessage(ByVal message As String)
    Debug.Print message
End Sub

