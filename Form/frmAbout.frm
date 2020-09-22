VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CC Antivir 2004"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdExit 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Thanks to:"
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4095
         Begin VB.Label lblThanks3 
            BackStyle       =   0  'Transparent
            Caption         =   "Johannes Knapp: for storing my files on his      Webscpace"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label lblThanks2 
            BackStyle       =   0  'Transparent
            Caption         =   "Paul: for his great CRC32 Checksum Algorithm "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label lblThanks 
            BackStyle       =   0  'Transparent
            Caption         =   "Patabugen: for storing my files on his Webspace (www.patabugen.co.uk ) "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   3735
         End
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Email: cyber_chris235@gmx.net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "© Copyright by Cyber Chris"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim myArticleAddr As String
    If MsgBox("Would you please vote on PSC Website in case you like this program?", vbQuestion + vbYesNo, "Your vote will be very well appreciated ...") = vbYes Then
        myArticleAddr = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=51592&optCodeRatingValue=5"
        Call ShellExecute(Me.hWnd, "Open", myArticleAddr, vbNullString, vbNullString, 1)
        MsgBox "Thank you very much. I really appreciate that :-) ", , "Thanks a million..."
    End If

End Sub

Private Sub lblCopyright_Click(Index As Integer)

    Call ShellExecute(Me.hWnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks2_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:wpsjr1@succeed.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:dude@patabugen.co.uk", vbNullString, "c:\", 1)

End Sub

