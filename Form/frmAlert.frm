VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivir 2004"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdSecure 
      Caption         =   "&Secure"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   585
      Left            =   360
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   4
      Top             =   240
      Width           =   645
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   12
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   11
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Virus found!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub BuildAlert()

    On Error Resume Next
    Log "Virus found: " & Virus.Reason & " in " & Virus.FileName
        lblText(1).Caption = Virus.FileName
    lblText(1).ToolTipText = Virus.FileName & "  (" & FileLen(Virus.FileName) & " Bytes )"
    lblText(2).Caption = Virus.Reason
    lblText(8).Caption = FileLen(Virus.FileName) & " Bytes"
    If Virus.Type = Executable Then
        lblText(7).Caption = "Executable File"
    End If
    If Virus.Type = Script Then
        lblText(7).Caption = "Script"
    End If
    picIcon.Picture = LoadIcon(Large, Virus.FileName)
    On Error GoTo 0

End Sub

Private Sub cmdIgnore_Click()

    Log "Virus Alert ignored!"
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    
    Log "Virus removed!"
    RemoveFile (Virus.FileName)

End Sub

Private Sub cmdSecure_Click()
On Error Resume Next
    MsgBox "The File will be secured, that means everytime you want to start it, you'll get a prompt." & vbCrLf & _
        "This will avoid unwanted starts!", vbInformation + vbOKOnly
    Dim sXor As New clsSimpleXOR
     sXor.EncryptFile Virus.FileName, Virus.FileName, AV.AVname
    Set sXor = Nothing
    MkDir App.Path & "\Secure\"
    FileCopy Virus.FileName, App.Path & "\Secure\" & Virus.FileNameShort & ".secure"
    Log "File secured: " & Virus.FileName
    Kill Virus.FileName
End Sub

Private Sub cmdView_Click()
    
    
    If MsgBox("WARNING! This will execute the file with the associated program !WARNING" & vbCrLf & "Continue?", vbCritical + vbYesNo, AV.AVname) = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", Virus.FileName, vbNullString, "c:\", 1)
        Log "Potential Virus executed"
    End If

End Sub



Private Sub Form_Load()

    BuildAlert

End Sub

