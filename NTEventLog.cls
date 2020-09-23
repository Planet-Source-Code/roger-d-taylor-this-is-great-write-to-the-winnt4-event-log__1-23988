VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Roger Taylor
'Roger_D_Taylor@Hotmail.com

'PLEASE READ:
'This software has been given in good will to you. i therefore request that you respect this and credit me within your program
'please feel free to pass this code on to your friends but please keep this text intact.
'
'
'Revision History:
'           None.
'
'To do:
' as of 05th June 2001 i have not yet managed to create the required resource file required by the windows nt event log.
' this resource file will stop the annoying message that you get in the event log.
'
' Your messages will be written to the event log but will have the text (Cannot find resource) or something like that prefixing
' your messages. if anyone manages to create this file before i do, can you please mail me with instructions on its construction and implementation please
'
'
'
'----------------------------------------------------
'Force Varaibles to be declared prior to execution
'----------------------------------------------------
Option Explicit

'----------------------------------------------------
'Private Declarations Section
'----------------------------------------------------
Private Const m_ClassName = "clsEV1"
Private mlngEventLog                As Long             'Handle of the NT Event log
Private mbolNTEventsLoaded          As Boolean          'Have we loaded NT Messaging?
Private mlngNTEventMessageCount     As Long             'NT Event message count
Private mtypNTEventMessages()       As NTEVENT_TYPE     'Array of NT Messages
Private mstrNTEventMachine          As String           'Destination of event "" = local machine
Private mstrAppName                 As String
Private Const mstrVersionNumber = "1.0.0"
Private Const NTEVENT_SOURCE = "My New Event Source"
Private Const NTEVENT_CATEGORYID = 1

Private Declare Sub NTEvent_MemCpy Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function NTEvent_RegisterSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function NTEvent_DeregisterSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function NTEvent_ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Function NTEvent_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

'
'USAGE:
' Dim x as new NTEventLog
' x.NTEventRegisterLog
' x.NTEventAddMessage 0,EVENTLOG_SUCCESS,Err.Description
' x.NTEventWrite 0, "Some additional error text"
' set x = nothing
'
'----------------------------------------------------
' Public Type Declarations Section
'----------------------------------------------------

'--------------------------------------------------------------------------------------------------------
' Public Function NTEventAddMessage(lngEventID As Long,
'                                   lngEventType As Long,
'                                   strMessage As String) As Long
'
' [in] lngEventID        =  A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] lngEventType      =  This is a numeric value that is retrieved from a public enum which has the following format
'                           EVENTLOG_SUCCESS          = 0
'                           EVENTLOG_ERROR_TYPE       = 1
'                           EVENTLOG_WARNING_TYPE     = 2
'                           EVENTLOG_INFORMATION_TYPE = 4
'                           EVENTLOG_AUDIT_SUCCESS    = 8
'                           EVENTLOG_AUDIT_FAILURE    = 10
'
' [in] strMessage        =  The event error description that is to be writtent to the event log
' [out] NTEventAddMessage  =  A long integer value that defines the error / success code
'
'
' Such That:
' A Type array is built of NT Event log messages that are to be written to the NT event Log
'
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------
'
Public Function NTEventAddMessage(lngEventID As Long, lngEventType As Long, strMessage As String) As Long

    
    'Increment the message count by one and resize the data
    'array for the NT Event messages
    mlngNTEventMessageCount = mlngNTEventMessageCount + 1
    ReDim Preserve mtypNTEventMessages(1 To mlngNTEventMessageCount)

    'Add the event information...
    mtypNTEventMessages(mlngNTEventMessageCount).intEventID = lngEventID
    mtypNTEventMessages(mlngNTEventMessageCount).intEventType = lngEventType
    mtypNTEventMessages(mlngNTEventMessageCount).strMessage = strMessage

    'Success
    NTEventAddMessage = 1


End Function

'--------------------------------------------------------------------------------------------------------
'Public Function NTEventRegisterLog() As Long
'
' [out] NTEventRegisterLog =  A long integer value that defines the error / success code
'
'
' Such That:
' The creator of the NT event message is registered with the event log as an event log user.
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------

Public Function NTEventRegisterLog() As Long
    
    
    'Check to see if NT Events have already been loaded
    If mbolNTEventsLoaded = True Then
        Exit Function
    End If
    
    ' Register this app. as new source, mhEventLog is handle
    mlngEventLog = NTEvent_RegisterSource(mstrNTEventMachine, NTEVENT_SOURCE)
   
   If mlngEventLog = Null Then
    NTEventRegisterLog = 3
    Exit Function
   End If
   
   'Success reported
   NTEventRegisterLog = 1
   
    'We have now connected to the NT Event manager
    mbolNTEventsLoaded = True
    

End Function

'--------------------------------------------------------------------------------------------------------
'Public Function NTEventWrite(lngEventID As Long,
'                             strAdditional As String) As Long
'
' [in] lngEventID    = A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] strAdditional = A string which defines additional text that will describe the event that occured
' [out] NTEventWrite =  A long integer value that defines the error / success code
'
' Such That:
' An Event is written to the windows NT 4 Event log
'
'
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'
'--------------------------------------------------------------------------------------------------------
'
Public Function NTEventWrite(lngEventID As Long, strAdditional As String) As Long
    
    
        
    Dim strMessage      As String          'Error message
    Dim lngEventType    As Long            'Error type
    Dim lngMsg          As Long            'Handle used for memory space allocation
    Dim lngLen          As Long            'String length
    Dim lngRet          As Long            'Function return
    
    'Configure additional text
    strAdditional = Trim(strAdditional)
    If Len(strAdditional) > 0 Then
        strAdditional = ": " & strAdditional
    End If
     

  
    'Now retrieve the error message from our array of messages
    Call NTEventGetMessage(lngEventID, strMessage, lngEventType)

  
    'Configure text of message and set up length
    strMessage = vbCrLf & vbCrLf & "App: " & mstrAppName & ", Ver: " & mstrVersionNumber & ", Date: " & Format(Now, "dd/mmm/yyyy") & vbCrLf & strMessage & strAdditional
    lngLen = Len(strMessage) + 1
    
    'Now allocate an amount of global memory for our error message
    lngMsg = GlobalAlloc(&H40, lngLen)
    
    'Copy message string into memory space obtained from GlobalAlloc
    NTEvent_MemCpy ByVal lngMsg, ByVal strMessage, lngLen
    
      'Now actually write the message to the log...
    lngRet = NTEvent_ReportEvent(mlngEventLog, lngEventType, NTEVENT_CATEGORYID, lngEventID, 0&, 1, lngLen, lngMsg, lngMsg)
    
    'Now free off our global memory
    Call GlobalFree(lngMsg)

    NTEventWrite = 1


End Function

'--------------------------------------------------------------------------------------------------------
'Private Sub NTEventGetMessage(lngEventID As Long,
'                              strMessage As String,
'                              lngEventType As Long)
'
' [in] lngEventID        =  A long integer value indicating the event id that is to be used when registering the event with the Evetn Log
' [in] strMessage        =  The event error description that is to be writtent to the event log
' [in] lngEventType      =  This is a numeric value that is retrieved from a public enum which has the following format
'                           EVENTLOG_SUCCESS          = 0
'                           EVENTLOG_ERROR_TYPE       = 1
'                           EVENTLOG_WARNING_TYPE     = 2
'                           EVENTLOG_INFORMATION_TYPE = 4
'                           EVENTLOG_AUDIT_SUCCESS    = 8
'                           EVENTLOG_AUDIT_FAILURE    = 10
'
' Such That:
' The NT event is retrieved from the event array as a call from NTEventWrite
'
' Author             : Roger Taylor
' Date               : 13th Dec 2000
' Last Modified Date :
'
' Revision history:
' - None
'--------------------------------------------------------------------------------------------------------
'
Private Sub NTEventGetMessage(lngEventID As Long, strMessage As String, lngEventType As Long)

    On Error GoTo Error_handler
        
    Dim lngCount            As Long         'Counter
  
    'Loop through all the messages and select the correct
    'message and event type for the event id
    For lngCount = 1 To mlngNTEventMessageCount
        If lngEventID = mtypNTEventMessages(lngCount).intEventID Then
            'Assign message and id and exit sub
            lngEventType = mtypNTEventMessages(lngCount).intEventType
            strMessage = mtypNTEventMessages(lngCount).strMessage
            Exit Sub
        End If
    Next
  
    'Did not find an entry for the event id so return a
    'standard message and info type
    strMessage = "Undefined Event"
    lngEventType = EVENTLOG_INFORMATION_TYPE
  
Exit Sub
Error_handler:

End Sub

