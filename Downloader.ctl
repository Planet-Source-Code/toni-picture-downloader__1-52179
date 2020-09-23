VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   Picture         =   "Downloader.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   510
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrPictureFromURL As String
Event Change(Download)
Event Maxim(Maxi)
Event Vallue(Valu)
'Event Stat(Statist)
           
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error Resume Next
   Select Case AsyncProp.PropertyName
      Case "PictureFromURL"
    RaiseEvent Vallue(AsyncProp.Value)
   End Select
Exit Sub
End Sub

Public Property Get PictureFromURL() As String
   PictureFromURL = mstrPictureFromURL
End Property

Public Property Let PictureFromURL(ByVal NewString As String)
On Error Resume Next
   ' (Code to validate path or URL omitted.)
   mstrPictureFromURL = NewString
   If (Ambient.UserMode = True) And (NewString <> "") Then
      ' If program is in run mode, and the URL string
      ' is not empty, begin the download.
      AsyncRead NewString, vbAsyncTypePicture, "PictureFromURL"
   End If
   PropertyChanged "PictureFromURL"
End Property

Public Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent Maxim(AsyncProp.BytesMax)
    RaiseEvent Change(AsyncProp.BytesRead)
    'RaiseEvent Stat(AsyncProp.Status)
End Sub

Private Sub UserControl_Resize()
UserControl.Size 480, 480
End Sub

