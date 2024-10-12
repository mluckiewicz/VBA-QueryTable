Attribute VB_Name = "Tester"
' Przyk³ad kodu w module standardowym
' Testowanie wzorca Obserwator z wieloma Publisherami
Option Explicit

Sub TestObserverPatternMultiplePublishers()

    Dim qt1 As Excel.QueryTable
    Dim qt2 As Excel.QueryTable
    Dim qt3 As Excel.QueryTable

    Set qt1 = Arkusz1.ListObjects(1).QueryTable
    Set qt2 = Arkusz2.ListObjects(1).QueryTable
    Set qt3 = Arkusz3.ListObjects(1).QueryTable

    ' Utworzenie instancji Publisherów
    Dim pub1 As New publisher
    Dim pub2 As New publisher
    Dim pub3 As New publisher
    pub1.SetQueryTable qt1
    pub2.SetQueryTable qt2
    pub3.SetQueryTable qt3
    
    ' Utworzenie instancji Observera
    Dim obs As Observer
    Set obs = New Observer
    
    ' Po³¹czenie obserwatora z dwoma Publisherami
    obs.RegisterPublisher pub1
    obs.RegisterPublisher pub2
    obs.RegisterPublisher pub3
    
End Sub

