VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileSplitter"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   4680
   Icon            =   "splitter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Fuser"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox txtfuse 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdfuse 
         Caption         =   "Fuse File"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdbrowse2 
         Caption         =   "Browse"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "File to fuse :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Splitter"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdsplit 
         Caption         =   "Split File"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdbrowse1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtsplit 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "File to split :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About FileSplitter..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
    '* This program was created by Nilanjan Sen.       *
    '* Please do not change/remove this text           *
    '* Feel free to edit the code as you wish          *
    '* Send comments to nilanjansen03@yahoo.co.in      *
    '* Components: Common dialog control 6.0           *
'*********************************************************
     
     Dim i, splitsize, cnt, splitcnt As Integer
     
Private Sub about_Click()
     Form3.Show
End Sub

Private Sub cmdbrowse1_Click()
     CommonDialog1.Filter = "All files(*.*)|*.*"
     CommonDialog1.ShowOpen
     txtsplit.Text = CommonDialog1.FileName
End Sub

Private Sub cmdbrowse2_Click()
     CommonDialog2.Filter = "Split Files(*.spt)|*.SPT|All files(*.*)|*.*"
     CommonDialog2.ShowOpen
     txtfuse.Text = CommonDialog2.FileName
End Sub

Private Sub cmdfuse_Click()
     Dim FileNum As Integer
     Dim FileNum2 As Integer
     Dim fusefile As String
     On Error GoTo errortrap
     FileName = Left(txtfuse.Text, Len(txtfuse.Text) - 5)
     For i = 2 To 9
           On Error GoTo jump
           If FileLen(FileName & Format(i, ".spt")) <> 0 Then splitcnt = i
     Next i
jump:
     FileNum = FreeFile
     Close FileNum
     Open FileName For Binary As FileNum
     For cnt = 1 To splitcnt
           FileNum2 = FreeFile
           Close FileNum2
           Open FileName & Format(cnt, ".spt") For Binary As FileNum2
                fusefile = Space(LOF(FileNum2))
                Get FileNum2, 1, fusefile
           Close FileNum2
           Put FileNum, LOF(FileNum) + 1, fusefile
           Kill FileName & Format(cnt, ".spt")
     Next cnt
     Close FileNum
     MsgBox ("File fused.")
     Exit Sub
errortrap:
     If Err = 5 Then MsgBox ("File not specified.")
     Exit Sub
End Sub

Private Sub cmdsplit_Click()
     Dim FileNum As Integer
     Dim FileNum2 As Integer
     Dim splitfile As String
     On Error GoTo errortrap
     str1 = FileLen(txtsplit.Text) / splitsize
     If str1 = Int(str1) Then
           splitcnt = str1
     Else
           splitcnt = Int(str1) + 1
     End If
     If splitcnt > 9 Then
           MsgBox ("File is too large to split.")
           Exit Sub
     End If
     If splitcnt < 2 Then
           MsgBox ("File is too small to split.")
           Exit Sub
     End If
     FileNum = FreeFile
     Close FileNum
     FileName = txtsplit.Text
     Open FileName For Binary As FileNum
     For cnt = 1 To splitcnt
           FileNum2 = FreeFile
           Close FileNum2
           Open FileName & Format(cnt, ".spt") For Binary As FileNum2
                splitfile = Space(splitsize)
                Get FileNum, (splitsize * (cnt - 1)) + 1, splitfile
                If cnt = splitcnt Then splitfile = Left(splitfile, LOF(FileNum) - (splitsize * (splitcnt - 1)))
                Put FileNum2, 1, splitfile
           Close FileNum2
     Next cnt
     Close FileNum
     MsgBox ("File splitted.")
     Exit Sub
errortrap:
     If Err = 53 Then MsgBox ("File not specified.")
End Sub

Private Sub Form_Load()
splitsize = 1457664
End Sub
