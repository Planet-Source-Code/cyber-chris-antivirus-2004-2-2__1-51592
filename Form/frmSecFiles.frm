VERSION 5.00
Begin VB.Form frmSecFiles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Secured Files"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSecFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdDesecure 
      Caption         =   "Desecure the file"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox flSec 
      Height          =   3015
      Left            =   0
      Pattern         =   "*.secure"
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmSecFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdDesecure_Click()

  Dim sXor As New clsSimpleXOR

    If MsgBox("Do you really want to desecure the file?", vbYesNo + vbCritical) = vbYes Then
        With App
            FileCopy .Path & "\secure\" & flSec.FileName, .Path & "\secure\" & Mid$(flSec.FileName, 1, Len(flSec.FileName) - 7)
            Kill .Path & "\secure\" & flSec.FileName
            sXor.DecryptFile .Path & "\secure\" & flSec.FileName, .Path & "\secure\" & flSec.FileName, AV.AVname
        End With
        Set sXor = Nothing
        flSec.Refresh
    End If

End Sub

Private Sub Form_Load()

    Me.flSec.Path = App.Path & "\secure\"

End Sub
