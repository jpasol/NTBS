VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSubicINVReprint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Invoice Reprint"
   ClientHeight    =   8055
   ClientLeft      =   4950
   ClientTop       =   2895
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6015
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   10815
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   6600
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtInvNum 
      Height          =   420
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display"
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtNumDay 
      Height          =   420
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin Crystal.CrystalReport CYInvoice 
      Left            =   6000
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Invoice Preview"
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Invoice Number"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "No. of Days (SA only)"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmSubicINVReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gcnnBilling As ADODB.Connection
Dim gbConnected As Boolean

'MDC (20131205)
Dim sqlConBilling As String
Dim sqlConNavis As String

Private Sub cmdDisplay_Click()
Dim rsINVict As New ADODB.Recordset
Dim tmpInvNo As Long

'    If Len(Trim(txtRefNum)) = 0 And Len(Trim(txtInvNum)) = 0 Then
'        MsgBox "Please specify valid entries.", vbExclamation, "Error Message"
'        Exit Sub
'    End If
    
    Screen.MousePointer = vbHourglass
    
'        If Len(Trim(txtNumDay)) > 0 Then    ' SA
'           CYInvoice.ReportFileName = App.Path & "\SubicINVSA1.rpt"
'            CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.ParameterFields(2) = "NumDays; " & Trim(txtNumDay) & ";TRUE"
'            CYInvoice.PrintReport
'        Else
'             CYInvoice.ReportFileName = App.Path & "\SubicInvoice1.rpt"
'             CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.ReportFileName = "c:\ntbs\cyinvoice\reprinting\SubicInvoice1.rpt"
'            CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum) & ";TRUE"
'            CYInvoice.PrintReport
'        End If
'
'    Else                             'Use of Invoice Number (for MR/Reg. bills only)
    
'        If Not gbConnected Then ConnectToBilling
'        With rsINVict
'            .Open "SELECT * FROM INVICT WHERE (invnum = " & txtInvNum & ")", _
'                   gcnnBilling, , , adCmdText
'            If Not .EOF Then
'                tmpInvNo = .Fields("refnum")
'                tmpInvNo = .Fields("invnum")
'            End If
'            .Close
'        End With
        'CYInvoice.ReportFileName = App.Path & "\SubicInvoice.rpt"
        'CYInvoice.ParameterFields(1) = "InvoiceNo; " & Trim(tmpInvNo) & ";TRUE"
        SubicInvoice.DiscardSavedData
        SubicInvoice.ParameterFields.GetItemByName("InvoiceNo").AddCurrentValue CDbl(Trim(txtInvNum.Text))
        CRViewer1.ReportSource = SubicInvoice
        CRViewer1.ViewReport
        
        Set SubicInvoice = Nothing
       
    'End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ConnectToBilling()
Dim gConnStr As String
Call ReadConfig
    gConnStr = sqlConBilling '"Provider=SQLOLEDB;Data Source=sbitcbilling;Initial Catalog=Billing;Integrated Security=SSPI"
    Set gcnnBilling = New ADODB.Connection
    gcnnBilling.Open gConnStr
    gbConnected = True
End Sub

Public Sub ReadConfig()
Dim Xcnt As Integer
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1

Do While Not EOF(1)
    Xcnt = Xcnt + 1
    Select Case Xcnt
        Case 1
            Line Input #1, sqlConBilling
        Case 2
            Line Input #1, sqlConNavis
    End Select
Loop
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
If ScaleHeight > CRViewer1.Top Then CRViewer1.Height = ScaleHeight - CRViewer1.Top
CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbConnected Then gcnnBilling.Close
End Sub

'Private Sub txtRefNum_GotFocus()
'    SendKeys "{HOME}": SendKeys "+{END}"
'End Sub
Private Sub txtInvNum_Change()

End Sub

Private Sub txtInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     cmdDisplay.SetFocus
  End If
End Sub
