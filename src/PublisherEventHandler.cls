VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublisherEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Klasa PublisherEventHandler - s³u¿y do nas³uchiwania zdarzeñ dla ka¿dego Publishera z osobna
Option Explicit

Private WithEvents Publisher As Publisher
Attribute Publisher.VB_VarHelpID = -1
Private observerRef As Observer ' Referencja do g³ównego obserwatora

' Inicjalizacja klasy - powi¹zanie z Publisherem i Obserwatorem
Public Sub Init(ByRef pub As Publisher, ByRef obs As Observer)
    Set Publisher = pub
    Set observerRef = obs
End Sub

' Obs³uga zdarzenia OnStateChanged - przekazywanie go do klasy Observer
Private Sub publisher_OnStateChanged(ByVal Success As Boolean)
    ' Wywo³anie metody w klasie Observer
    observerRef.HandleStateChange Publisher, Success
End Sub

