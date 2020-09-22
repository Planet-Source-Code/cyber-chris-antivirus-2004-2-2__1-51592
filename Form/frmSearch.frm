VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivirus 2004"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   3240
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picHelpAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2895
      TabIndex        =   27
      Top             =   840
      Width           =   2895
      Begin VB.PictureBox picAbout 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":8A12
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   30
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox picHelp 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":C0A4
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblAbuthelp 
         Caption         =   "   Help / About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2895
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   8
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   7
         X1              =   2880
         X2              =   0
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   6
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1680
      End
   End
   Begin VB.PictureBox picOther 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   23
      Top             =   480
      Width           =   2895
      Begin VB.PictureBox picUpdate 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":E1CE
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox picSec 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":1036C
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   5
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   4
         X1              =   2880
         X2              =   0
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Label lblOther 
         Caption         =   "   Extra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picScan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   18
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox picFastSearchx 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmSearch.frx":136DE
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   21
         Top             =   1560
         Width           =   2055
      End
      Begin VB.PictureBox picPathsearch 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":16FCC
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox picFileSearch 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":1AAD6
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   2
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   1
         X1              =   2880
         X2              =   0
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Label lblFileScan 
         Caption         =   "   File Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   11400
      ScaleHeight     =   2955
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan the selected File"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checksum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   825
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   8
      X1              =   6840
      X2              =   6840
      Y1              =   -720
      Y2              =   0
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   7
      X1              =   6360
      X2              =   6360
      Y1              =   -720
      Y2              =   0
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run on startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tray window:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Signatures:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   3120
      X2              =   3120
      Y1              =   4080
      Y2              =   -120
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   2
      X1              =   8640
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anti Virus Definitions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   1
      X1              =   8640
      X2              =   3120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Files checked:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private WithEvents X     As cCommonDialog
Attribute X.VB_VarHelpID = -1
Private CurrentFile      As String
Dim sPicScan, SpicOther, sHelpAbout As pStatus

Private Sub cmdScan_Click()

    CheckFile (CurrentFile)

End Sub

Private Sub Form_Load()

    Set X = New cCommonDialog
    Set ccClass = X
    frmMain.Cls
    BuildUI
    sPicScan = Min
    SpicOther = Min
    sHelpAbout = Max
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblAbuthelp_Click()
Dim temp, temp2 As Integer

    
If sHelpAbout = Max Then
        For temp = 1 To 1695 - 255
            picHelpAbout.Height = picHelpAbout.Height - 1
            DoEvents
        Next
        sHelpAbout = Min
ElseIf sHelpAbout = Min Then
        If sPicScan = Max Then lblFileScan_Click
        If SpicOther = Max Then lblOther_Click
        For temp = 1 To 1695 - 255
            picHelpAbout.Height = picHelpAbout.Height + 1
            DoEvents
        Next
        sHelpAbout = Max
End If
End Sub

Private Sub lblFileScan_Click()
Dim temp, temp2 As Integer

If sPicScan = Max Then
        For temp = 1 To 2415 - 255
            picScan.Height = picScan.Height - 1
            picOther.Top = picScan.Top + picScan.Height + 20
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next
        sPicScan = Min
ElseIf sPicScan = Min Then
        If sHelpAbout = Max Then lblAbuthelp_Click
        If SpicOther = Max Then lblOther_Click
        For temp = 1 To 2415 - 255
            picScan.Height = picScan.Height + 1
            picOther.Top = picScan.Top + picScan.Height + 20
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next
        sPicScan = Max
End If

End Sub

Private Sub lblOther_Click()
Dim temp, temp2 As Integer
If SpicOther = Max Then

        For temp = 1 To 1815 - 255
            picOther.Height = picOther.Height - 1
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next
        SpicOther = Min
ElseIf SpicOther = Min Then
        If sHelpAbout = Max Then lblAbuthelp_Click
        If sPicScan = Max Then lblFileScan_Click

        For temp = 1 To 1815 - 255
            picOther.Height = picOther.Height + 1
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next
        SpicOther = Max
End If
End Sub

Private Sub lblText_Click(Index As Integer)

    On Error Resume Next
    If Index = 6 Then
        If lblText(7).Caption = "OFF" Then
            frmTray.Show , Me
         Else 'NOT LBLTEXT(7).CAPTION...
            Unload frmTray
        End If
     ElseIf Index = 9 Then 'NOT INDEX...
        If lblText(8).Caption = "OFF" Then
            SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname, App.Path & "\" & App.EXEName & ".exe /T", 1
            lblText(8).Caption = "ON"
            SaveSetting AV.AVname, "Settings", "Startup", "ON"
         Else 'NOT LBLTEXT(8).CAPTION...
            DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname
            lblText(8).Caption = "OFF"
            SaveSetting AV.AVname, "Settings", "Startup", "OFF"
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub picAbout_Click()

    frmAbout.Show , Me

End Sub

Private Sub picFastsearch_Click()

    X.ControlToSetNewParent = Picture1
    Debug.Print X.ShowOpen(Me.hWnd)
    

End Sub

Private Sub picFastSearchx_Click(Index As Integer)

    X.ControlToSetNewParent = Picture1
    Debug.Print X.ShowOpen(Me.hWnd)

End Sub

Private Sub picFileSearch_Click()

    Call ShowFileSearch

End Sub

Private Sub picHelp_Click()

    frmHelp.Show , Me

End Sub

Private Sub picPathsearch_Click()

    Checkfolder

End Sub

Private Sub picSec_Click()
    frmSecFiles.Show
End Sub

Private Sub picUpdate_Click()

    frmUpdate.Show , Me

End Sub

Public Sub ShowFileSearch()

  Dim strfilename As String

    On Error Resume Next
    strfilename = (ShowOpenDlg(Me, , "All Files|*.*", , "Scan File"))
    If FileLen(strfilename) <> 0 Then
        CheckFile (strfilename)
    End If
    On Error GoTo 0

End Sub

Private Sub X_FileChanged(ByVal FileName As String)

    lblFileName.Caption = Mid$(FileName, InStrRev(FileName, "\") + 1)
    lblText(12).Caption = FileLen(FileName) & " Bytes"
    lblText(14).Caption = CalcCRC(FileName)
    CurrentFile = FileName

End Sub

