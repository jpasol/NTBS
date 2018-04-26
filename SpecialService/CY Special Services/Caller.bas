Attribute VB_Name = "modCaller"
Option Explicit

Sub Main()
Dim gConnStr As String
Dim c As Object

    gConnStr = "Provider=sqloledb; Data Source=itssdevapps4; Initial Catalog=BILLING; Integrated Security=SSPI"

    Set c = CreateObject("SubicCYSCCR.cCYSCCR")
    With c
        Call .ConnectByStr(gConnStr, "glacorte")
        Call .Execute
        Call .Disconnect
    End With
    Set c = Nothing
End Sub
