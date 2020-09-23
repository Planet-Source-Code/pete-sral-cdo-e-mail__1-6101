VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CDO E-mail"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4275
      TabIndex        =   7
      Top             =   2385
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   2895
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      Height          =   1200
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmMain.frx":0442
      Top             =   1110
      Width           =   5355
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   780
      TabIndex        =   4
      Text            =   "This is a TEST"
      Top             =   495
      Width           =   4620
   End
   Begin VB.TextBox txtTo 
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Text            =   "Pete@pjs-inc.com"
      Top             =   90
      Width           =   2595
   End
   Begin VB.Label Label3 
      Caption         =   "Message"
      Height          =   300
      Left            =   135
      TabIndex        =   2
      Top             =   885
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Subject"
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   525
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      Height          =   300
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsMail As New clsCDOMail


Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdSend_Click()

    clsMail.SendToEmail = txtTo.Text
    clsMail.Subject = txtSubject.Text
    clsMail.Message = txtMessage.Text
    
    If MsgBox("Send Mail?", vbQuestion + vbYesNoCancel) = vbYes Then
        clsMail.SendMail "Pete Sral"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (clsMail Is Nothing) Then
        Set clsMail = Nothing
    End If
    
End Sub
