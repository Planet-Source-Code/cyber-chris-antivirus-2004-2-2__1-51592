Attribute VB_Name = "modSearch"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sign(4096)               As String    'The Signatures will be loaded into this array
Private SignVirusType(4096)      As String * 1
Private SignVirusName(4096)      As String

Public Sub BuildSigns()

  'This builds the Signature - Array
  
  Dim sIn        As String
  Dim swords()   As String
  Dim X          As Long
  Dim Data()     As String
    sIn = FileText(AV.Signature.SignatureFilename)
    swords = Split(sIn, vbCrLf)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        Data = Split(swords(X) & ":" & ":", ":")
        Sign(X) = Data(0)
        SignVirusType(X) = Data(1)
        SignVirusName(X) = Data(2)
    Next X
    Sign(X + 1) = "#END#"
    AV.Signature.SignatureDate = Sign(0)
    AV.Signature.SignatureCount = UBound(swords) - 1

Exit Sub

err:
    MsgBox "An error has occured while loading the signature File!" & vbCrLf & "This could be caused by an empty or damaged file!" & vbCrLf & vbCrLf & "The error message was: " & err.Description, vbCritical + vbOKOnly, "Error"

End Sub

Public Function Search(ByVal strfilename As String) As String

  Dim Current  As Long
  Dim CRC      As String

    CRC = CalcCRC(strfilename)
    For Current = 1 To 4096
        If Sign(Current) = "#END#" Or LenB(Sign(Current)) = 0 Then
            GoTo Finish
        End If
        If CRC = Sign(Current) Then
            DoEvents
            Search = SignVirusName(Current)
            Select Case SignVirusType(Current)
             Case "E"
                Virus.Type = Executable
             Case "S"
                Virus.Type = Script
            End Select
            Exit Function
         Else 'NOT FINDTERM(FNAME,...'NOT CRC...
            Search = "NOTHING"
        End If
        DoEvents
    Next Current
Finish:

End Function

