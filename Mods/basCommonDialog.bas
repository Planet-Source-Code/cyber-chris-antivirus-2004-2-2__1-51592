Attribute VB_Name = "basCommonDialog"

Option Explicit
Public ccClass     As cCommonDialog

Public Function ComDlgCallback(ByVal hWndDlg As Long, _
                               ByVal Msg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long

    If Not ccClass Is Nothing Then
        ccClass.pIncomingMessage hWndDlg, Msg, wParam, lParam
    End If

End Function

