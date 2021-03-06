VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCYSCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub Execute(Optional ByVal pOfficer As Boolean = False)
    If gbConnected Then
        MsgBox "No SQL Server connection found!"
    Else
        If pOfficer Then
            frmCYSCCR.Show
        Else
            frmCYSCCR.Show 1
        End If
    End If
End Sub

Public Sub Disconnect()
    On Error Resume Next
    gcnnBilling.Close
    Set gcnnBilling = Nothing
End Sub

Public Function ConnectByStr(ByVal pCnnStr As String, ByVal pUserID As String) As Boolean
Dim errBilling As ADODB.Error
Dim lsErrStr As String
   
    ' Open the database.
    On Error GoTo err_Connect
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open pCnnStr
    gUserID = pUserID
    ConnectByStr = True
   
    Exit Function
    
err_Connect:
    ConnectByStr = False: gbConnected = False
    For Each errBilling In gcnnBilling.Errors
        With errBilling
            lsErrStr = "Connection Error. " & .Description & vbLf & _
            "Verify Log On then retry.  Contact MIS for assistance."
        End With
        MsgBox lsErrStr, vbCritical
    Next
End Function

Public Sub ReprintCCR(Optional ByVal pOfficer As Boolean = False)
    If gbConnected Then
        MsgBox "No SQL Server connection found!"
    Else
       If pOfficer Then
          frmReprintCCR.Show
       Else
          frmReprintCCR.Show 1
       End If
    End If
End Sub

Public Sub CCRSuper(ByVal pSupervisor As String)
   vSupervisor = pSupervisor
End Sub
