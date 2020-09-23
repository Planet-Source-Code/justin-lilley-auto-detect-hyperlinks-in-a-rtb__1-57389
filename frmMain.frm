VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   360
      Top             =   3240
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By: Justin Lilley
'http://mindzpro.com
'Please Vote and leave feedback!
'Just made real quit cuZ theres nothin on PSC Like it =P

Private Sub Command1_Click()
Dim yournamehere As String
If txtSend.Text = "" Then
Exit Sub
Else
yournamehere = "Please Vote"
AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, txtSend.Text & vbCrLf, vbWhite
    txtSend.Text = ""
    txtSend.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    If Len(frmMain.RTB1.Text) >= 1000 Then
        frmMain.RTB1 = ""
    Else
    End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
Dim yournamehere As String
If txtSend.Text = "" Then
txtSend.SetFocus
Exit Sub
Else
If KeyAscii = 13 Then
yournamehere = "Please Vote"
AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, txtSend.Text & vbCrLf, vbWhite
    KeyAscii = 0
    txtSend.Text = ""
    txtSend.SetFocus
End If
End If
End Sub


Private Sub Form_Load()
Dim yournamehere As String
yournamehere = "Please Vote"
    Me.Caption = "RTB Colors and Hyperlink Auto Detect. By Justin Lilley."
    EnableURLDetect RTB1.hwnd, Me.hwnd
    AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, "Detects Both The Http and Www" & vbCrLf, vbWhite
        AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, "Double Click And Go..." & vbCrLf, vbWhite
AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, "www.Mindzpro.com" & vbCrLf, vbWhite
AddTxt vbWhite, "<|", vbWhite, yournamehere, vbWhite, "|> ", vbWhite, "http://Mindzpro.com" & vbCrLf, vbWhite
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect
End Sub
