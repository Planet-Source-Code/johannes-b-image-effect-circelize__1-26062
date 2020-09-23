VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circelize: By JOHANNES B 2001"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5280
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "4"
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   3720
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   10
      Top             =   0
      Width           =   1575
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   120
      ScaleHeight     =   4260
      ScaleWidth      =   5670
      TabIndex        =   8
      Top             =   120
      Width           =   5700
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H8000000E&
         Height          =   4260
         Left            =   0
         Picture         =   "Circelize.frx":0000
         ScaleHeight     =   280
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   374
         TabIndex        =   9
         Top             =   0
         Width           =   5670
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "7"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete effect from picture"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save picture..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   5655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load picture..."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "4"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Circelize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "Thickness:"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Step Y:"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Circle size:"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Step X:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer, B As Integer, Sx As Integer, Sy As Integer, St As Integer, Th As Integer
Dim PT As String
Private Sub Command1_Click()
On Error Resume Next

Set Picture3.Picture = Picture1.Image
'Reset head
A = 0
B = 0
'Set step
Sx = Text1.Text
Sy = Text3.Text
'Set size of circle
S = Text2.Text


Do
'Move head right
A = A + Sx
If A > Picture1.ScaleWidth Then
'Move head to X = 1 and move one step down
A = 1
B = B + Sy
Command1.Caption = "PLEASE WAIT... ROW " & B & " OF " & Picture1.ScaleHeight + S
Picture1.Refresh
End If
'Get color at head position
Picture1.ForeColor = GetPixel(Picture3.hdc, A, B)

'Draw circle at head position
Picture1.Circle (A, B), S

'Loop until head is reaching bottom of picture
Loop Until B > Picture1.ScaleHeight + S
'Done!
Picture1.Refresh
Command1.Caption = "Circelize"
End Sub

Private Sub Command2_Click()
On Error GoTo kalle
'Filter extansions that will be showed
CM.Filter = "Bitmap image|*.bmp"
'So it wount continue if you press cancel
CM.CancelError = True
'Show the dialog
CM.ShowSave
'Save the picture
SavePicture Picture1.Image, CM.FileName
Exit Sub
'If you press cancel or if the filename is invalid
kalle:
Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo kalleanka
'Filter extansions that will be showed
CM.Filter = "Image files|*.bmp;*.jpg;*.gif;*.wmf;*.emf;*.cur;*.ico"
'So it wount continue if you press cancel
CM.CancelError = True
'Show the dialog
CM.ShowOpen
'Open the picture
Picture1.Picture = LoadPicture(CM.FileName)
Exit Sub
'If you press cancel or if the picture is invalid
kalleanka:
Exit Sub
End Sub

Private Sub Command4_Click()
Picture1.Cls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Do
Form1.Top = Form1.Top + 1
Loop Until Form1.Top > Screen.Height

MsgBox "Hope you liked it. Please vote!", vbInformation

End

End Sub

Private Sub Text4_Change()
On Error GoTo kaka
Picture1.DrawWidth = Text4.Text
Exit Sub
kaka:
Text4.Text = "1"
Exit Sub
End Sub


