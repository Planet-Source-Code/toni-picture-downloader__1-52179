VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3150
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   2160
      Width           =   135
   End
   Begin Project1.Downloader Downloader1 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Max

Private Sub Downloader1_Change(Download As Variant)
On Error Resume Next
If Max = 0 Then Max = 1
If Download = 0 Then Download = 1
Picture1.Width = ((Me.ScaleWidth / Maxi) * Download)
'Debug.Print Max, Download
End Sub

Private Sub Downloader1_Maxim(Maxi As Variant)
Max = Maxi
End Sub

'Private Sub Downloader1_Stat(Statist As Variant)
'On Error Resume Next
'If Max = 0 Then Max = 1
'Picture1.Width = ((Me.ScaleWidth / Max)*Statist)
'End Sub

Private Sub Downloader1_Vallue(Valu As Variant)
Set Me.Picture = Valu
End Sub

Private Sub Form_Load()
Me.Caption = "Drag a Link On me"
Downloader1.PictureFromURL = "http://rr.exhedra.com/upload_PSC/AuthorPhotos/AUTHOR_PHOTO2003413327101424.bmp"
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Downloader1.PictureFromURL = Data.GetData(vbCFText)
End Sub

Private Sub Form_Resize()
Picture1.Move 0, Me.ScaleHeight - Picture1.Height
End Sub
