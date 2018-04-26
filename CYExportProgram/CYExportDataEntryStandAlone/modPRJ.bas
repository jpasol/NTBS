Attribute VB_Name = "modPRJ"
Option Explicit
Public gConnStr As String
Private Sub Main()
    Dim mp As clsCCRde06
    gConnStr = "Provider=sqloledb" & _
        ";Data Source=" & Trim("SBITCBILLING") & _
        ";Initial Catalog=" & Trim("BILLING") & _
        ";Integrated Security=SSPI"
    Set mp = New clsCCRde06
    mp.Userid = "rcalvo"
    mp.ConnectByStr gConnStr
    mp.Execute
    mp.Disconnect
    Set mp = Nothing
End Sub
