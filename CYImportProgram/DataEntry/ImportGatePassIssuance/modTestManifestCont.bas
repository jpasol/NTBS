Attribute VB_Name = "modTestManifestCont"
Option Explicit
Dim gConnStr As String
Private Sub main()
   Dim MS As Object
   gConnStr = "Provider=sqloledb" & _
        ";Data Source=itssdevapps4" & _
        ";Initial Catalog=sbitcbilling" & _
        ";Integrated Security=SSPI"
    Set MS = CreateObject("prjManifestCont.clsCYMDE01")
    'MS.Userid = "borillano"
    MS.ConnectByStr gConnStr
    MS.Execute "lTenorio"
    MS.Disconnect
    Set MS = Nothing
End Sub


