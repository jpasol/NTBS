VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Tariff Rates Maintenance"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "C L O S E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "L O G I N "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ConnectToServer() As Boolean
    Dim gconnstr As String
    Dim x As Object
    
    
    gconnstr = "Provider=sqloledb" & _
                ";Data Source=" & "sbitcbilling" & _
                ";Initial Catalog=" & "BILLING" & _
                ";Integrated Security=SSPI"
    
    Set x = CreateObject("CYRatesMaintenance.clsCYRates")
    With x
        .ConnectByStr (gconnstr)
        .Execute
        .Disconnect
    End With
    Set x = Nothing
End Function

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdLogin_Click()
    Call ConnectToServer
End Sub

