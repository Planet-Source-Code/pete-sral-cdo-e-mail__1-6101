VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDOMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarSendToEmail As String 'local copy
Private mvarSubject As String 'local copy
Private mvarMessage As String 'local copy
Public Property Let Message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Message = 5
    mvarMessage = vData
End Property


Public Property Get Message() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Message
    Message = mvarMessage
End Property



Public Property Let Subject(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Subject = 5
    mvarSubject = vData
End Property


Public Property Get Subject() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Subject
    Subject = mvarSubject
End Property



Public Property Let SendToEmail(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SendToEmail = 5
    mvarSendToEmail = vData
End Property


Public Property Get SendToEmail() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SendToEmail
    SendToEmail = mvarSendToEmail
End Property




Public Function SendMail(psProfileName As String, Optional psSentToEmail As String, Optional psSubject As String, Optional psMessage As String)

On Error GoTo SendMail_Error

   Dim oSession As MAPI.Session
   Dim oMessage As Message
   Dim oRecip As Recipient
   'local
   Dim sRecip As String
   Dim sSubject As String
   Dim sMessage
   
   'Handle Optional parameters
   If Len(psSentToEmail) <> 0 Then
    sRecip = psSentToEmail
   Else
    sRecip = mvarSendToEmail
    End If
    
    If Len(psSubject) <> 0 Then
        sSubject = psSubject
    Else
        sSubject = mvarSubject
    End If
    
    If Len(psMessage) <> 0 Then
        sMessage = psMessage
    Else
        sMessage = mvarMessage
    End If
    
    'TODO: need to check if class variables are not empty
    
   Set oSession = CreateObject("MAPI.Session")
   oSession.Logon psProfileName
   Set oMessage = oSession.Outbox.Messages.Add
   oMessage.Subject = sSubject
   Set oRecip = oMessage.Recipients.Add(sRecip)
   oRecip.Resolve
   oMessage.Text = sMessage
   
   oMessage.Send , False
   oSession.Logoff
   Exit Function
   
SendMail_Error:
    MsgBox Str(Err) & " - " & Error & ". Sending e-mail to " & sRecip, vbExclamation
End Function
 
Public Sub Clear()
mvarSendToEmail = ""
mvarSubject = ""
mvarMessage = ""
End Sub
