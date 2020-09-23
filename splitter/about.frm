VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   ControlBox      =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton unloadfrm 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "EMail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   360
         Picture         =   "about.frx":000C
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lblmsg 
         Caption         =   "Copyright (c) 2004 Nilanjan Sen"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lblurl 
         Caption         =   "http://www.geocities.com/nilanjansen03"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         MouseIcon       =   "about.frx":5C1E
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Visit developer's homepage on the net "
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblmail 
         Caption         =   "nilanjansen03@yahoo.co.in"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         MouseIcon       =   "about.frx":5F28
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "EMail to the developer"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lbldeveloper 
         Alignment       =   2  'Center
         Caption         =   "AppName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblappname 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AppName and Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const AppName = "FileSplitter"
Dim message As String

Private Sub unloadfrm_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     Me.Caption = "About " & AppName
     lblappname.Caption = AppName & " Version :  " & App.Major & "." & App.Minor & "  Build :  " & App.Revision
     lbldeveloper.Caption = AppName & " is developed by Nilanjan Sen"
End Sub

Private Sub lblmail_Click()
     Call ShellExecute(&O0, vbNullString, "mailto:" & lblmail.Caption & "?Subject=" & App.Title, vbNullString, vbNullString, vbNormalFocus)
     End Sub

Private Sub lblurl_Click()
     Call ShellExecute(&O0, vbNullString, lblurl.Caption, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblmsg_Click()
     message = "This is a freeware version of " & AppName & "." & vbCrLf
     message = message + "You can distribute it among your friends without any permission." & vbCrLf & vbCrLf
     message = message + "This software is AS IS without warranty of any kind."
     message = message + " While every possible care is taken, to ensure that the software is efficient and bug free."
     message = message + " However any kind of suggestions are invited from the users to make the software more efficient and full-proof."
     Call MsgBox(message, vbInformation + vbOKOnly, AppName)
End Sub
