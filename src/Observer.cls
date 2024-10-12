VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Observer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Observer Class - can receive notifications from multiple Publishers
Option Explicit

' Dictionary to store handlers and states of Publishers
Private publishers As Object

' Class Initialization
Private Sub Class_Initialize()
    ' Create a dictionary for storing handlers and states
    Set publishers = CreateObject("Scripting.Dictionary")
End Sub

' Register a new Publisher to observe
' Publisher: Instance of the Publisher class to be registered
Public Sub RegisterPublisher(ByRef Publisher As Publisher)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))

    ' Check if the Publisher is already registered
    If Not IsPublisherRegistered(publisherId) Then
        ' Initialize and add the new Publisher
        Dim handler As PublisherEventHandler
        Set handler = CreateNewHandler(Publisher)
        
        RegisterNewPublisher publisherId, handler
        LogMessage "Registered publisher: " & publisherId & " in observer " & ObjPtr(Me)
    Else
        LogMessage "Publisher with ID " & publisherId & " is already registered."
    End If
End Sub

' Checks if the Publisher with the given Id is already registered
' publisherId: ID of the Publisher to check
' Returns True if the Publisher is registered, otherwise False
Private Function IsPublisherRegistered(ByVal publisherId As String) As Boolean
    IsPublisherRegistered = publishers.Exists(publisherId)
End Function

' Creates a new PublisherEventHandler object and assigns it to the Publisher
' Publisher: Instance of the Publisher class for which the handler is created
' Returns a new PublisherEventHandler object
Private Function CreateNewHandler(ByRef Publisher As Publisher) As PublisherEventHandler
    Dim handler As PublisherEventHandler
    Set handler = New PublisherEventHandler
    handler.Init Publisher, Me ' Initialize the handler
    Set CreateNewHandler = handler
End Function

' Registers a new Publisher in the main dictionary
' publisherId: ID of the Publisher to register
' handler: PublisherEventHandler object associated with this Publisher
Private Sub RegisterNewPublisher(ByVal publisherId As String, ByRef handler As PublisherEventHandler)
    ' Create a new dictionary for the given Publisher
    Dim publisherData As Object
    Set publisherData = CreateObject("Scripting.Dictionary")
    
    ' Assign the handler and initialize the state
    publisherData.Add "handler", handler
    publisherData.Add "state", False ' Initial state
    
    ' Add to the main publishers dictionary
    publishers.Add publisherId, publisherData
End Sub

' Unregister a Publisher (optional method)
' Publisher: Instance of the Publisher class to unregister
Public Sub UnregisterPublisher(ByRef Publisher As Publisher)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))
    
    If publishers.Exists(publisherId) Then
        publishers.Remove publisherId
        LogMessage "Unregistered publisher: " & publisherId
    Else
        LogMessage "Publisher with ID " & publisherId & " does not exist."
    End If
End Sub

' Method that handles state changes from any Publisher
' Publisher: Instance of the Publisher class that reports a state change
' Success: Boolean indicating whether the state change was successful
Public Sub HandleStateChange(ByRef Publisher As Publisher, ByVal Success As Boolean)
    Dim publisherId As String
    publisherId = CStr(ObjPtr(Publisher))

    ' Check and update the state of the given Publisher
    If publishers.Exists(publisherId) Then
        UpdatePublisherState publisherId, Success
        LogMessage "Observer " & ObjPtr(Me) & " received notification from Publisher " & ObjPtr(Publisher) & ": " & Success

        ' Check if all Publishers are updated
        If AllPublishersUpdated() Then
            HandleAllPublishersUpdated
        End If
    End If
End Sub

' Updates the state for a given Publisher
' publisherId: ID of the Publisher whose state needs to be updated
' Success: Boolean indicating the new state of the Publisher
Private Sub UpdatePublisherState(ByVal publisherId As String, ByVal Success As Boolean)
    publishers(publisherId)("state") = Success
End Sub

' Checks if all Publishers have been updated
' Returns True if all Publishers are updated, otherwise False
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

' Actions performed when all Publishers are updated
Private Sub HandleAllPublishersUpdated()
    LogMessage "All Publishers' states have changed."
    ResetPublisherStates
End Sub

' Reset the states of all Publishers
Private Sub ResetPublisherStates()
    Dim key As Variant
    For Each key In publishers.Keys
        publishers(key)("state") = False ' Reset the state
    Next key
    LogMessage "Publishers' states have been reset."
End Sub

' Universal method for logging messages
' message: Content of the message to log
Private Sub LogMessage(ByVal message As String)
    Debug.Print message
End Sub

