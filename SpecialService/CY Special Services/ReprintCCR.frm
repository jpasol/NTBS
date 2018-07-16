VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmReprintCCR 
   Caption         =   "CY Specical Services CCR Reprinting"
   ClientHeight    =   2385
   ClientLeft      =   4275
   ClientTop       =   3525
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReprintCCR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   5010
   Begin VB.CommandButton cmdReprint 
      Caption         =   "Re&print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   375
      TabIndex        =   2
      Top             =   1575
      Width           =   2265
   End
   Begin MSMask.MaskEdBox txtCCRRefNo 
      Height          =   465
      Left            =   2925
      TabIndex        =   0
      Top             =   300
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtCCRNum 
      Height          =   465
      Left            =   2925
      TabIndex        =   1
      ToolTipText     =   " Leave blank for all CCRs "
      Top             =   900
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "CCR Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   375
      TabIndex        =   6
      Top             =   900
      Width           =   2340
   End
   Begin VB.Label lblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   420
      Index           =   1
      Left            =   3675
      TabIndex        =   5
      Top             =   1650
      Width           =   1005
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Reference No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   375
      TabIndex        =   3
      Top             =   300
      Width           =   2340
   End
End
Attribute VB_Name = "frmReprintCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsCCRReprint As Object

Private Sub cmdReprint_Click()
'    With clsCCRReprint
''        On Error GoTo err_Reprint
'        .CCRSupervisor vSupervisor
'        .CCRNumber = CLng("0" & txtCCRNum)
'        .PrintCCR CLng("0" & Trim(txtCCRRefNo))
        OutCCRPC CStr(Trim(txtCCRRefNo)), CStr(Trim(txtCCRNum))
        txtCCRRefNo = Space(txtCCRRefNo.MaxLength)
        txtCCRNum = Space(txtCCRNum.MaxLength)
'    End With
'    Exit Sub
'err_Reprint:
'    On Error GoTo 0
'    MsgBox "Reference/CCR number " & Trim(txtCCRRefNo) & " not found", vbInformation
'    txtCCRRefNo.SetFocus
End Sub

Private Sub cmdReprint_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub Form_Load()
'<Version/>
Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
'</Version>
    'Set clsCCRReprint = CreateObject("CCRPR03.clsCCRPR03")
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set clsCCRReprint = Nothing
End Sub

Private Sub txtCCRNum_GotFocus()
    With txtCCRNum
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCCRNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub txtCCRNum_LostFocus()
    With txtCCRNum
        .BackColor = vbWindowBackground
    End With
End Sub

Private Sub txtCCRRefNo_Change()
    cmdReprint.Enabled = (CLng(parsezero("0" & txtCCRRefNo)) > 0)
    
    Dim ado As Recordset

End Sub

Private Function parsezero(var As Variant) As Variant
On Error GoTo parse
parsezero = CLng(var)
Exit Function
parse:
parsezero = 0
End Function
Private Sub txtCCRRefNo_GotFocus()
    With txtCCRRefNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCCRRefNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case Else
    End Select
End Sub

Private Sub txtCCRRefNo_LostFocus()
    Dim getCCR As ADODB.Recordset
    
    With txtCCRRefNo
        .BackColor = vbWindowBackground
    End With
End Sub


Private Function NumToText(dblValue As Currency) As String
    Static ones(0 To 9) As String
    Static teens(0 To 9) As String
    Static tens(0 To 9) As String
    Static thousands(0 To 4) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strTemp As String
    Dim tmpBuff As String
    Dim strSign As String
    Dim negativeSign As Boolean

    ones(0) = "zero"
    ones(1) = "one"
    ones(2) = "two"
    ones(3) = "three"
    ones(4) = "four"
    ones(5) = "five"
    ones(6) = "six"
    ones(7) = "seven"
    ones(8) = "eight"
    ones(9) = "nine"

    teens(0) = "ten"
    teens(1) = "eleven"
    teens(2) = "twelve"
    teens(3) = "thirteen"
    teens(4) = "fourteen"
    teens(5) = "fifteen"
    teens(6) = "sixteen"
    teens(7) = "seventeen"
    teens(8) = "eighteen"
    teens(9) = "nineteen"

    tens(0) = ""
    tens(1) = "ten"
    tens(2) = "twenty"
    tens(3) = "thirty"
    tens(4) = "forty"
    tens(5) = "fifty"
    tens(6) = "sixty"
    tens(7) = "seventy"
    tens(8) = "eighty"
    tens(9) = "ninety"
    
    thousands(0) = ""
    thousands(1) = "thousand"
    thousands(2) = "million"
    thousands(3) = "billion"
    thousands(4) = "trillion"

    'Trap errors
    On Error GoTo NumToTextError
    'Get fractional part
    If dblValue < 0 Then
        negativeSign = True
        dblValue = Abs(dblValue)
    Else
        negativeSign = False
    End If
    strResult = "& " & Format((dblValue - Int(dblValue)) * 100, "00") & "/100"
    If negativeSign Then
        strSign = "NEGATIVE "
    Else
        strSign = ""
    End If
    strTemp = CStr(Int(dblValue))
    'Iterate through string
    For i = Len(strTemp) To 1 Step -1
        'Get value of this digit
        nDigit = Val(Mid$(strTemp, i, 1))
        'Get column position
        nPosition = (Len(strTemp) - i) + 1
        'Action depends on 1's, 10's or 100's column
        Select Case (nPosition Mod 3)
            Case 1  '1's position
                bAllZeros = False
                If i = 1 Then
                    tmpBuff = ones(nDigit) & " "
                ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
                    tmpBuff = teens(nDigit) & " "
                    i = i - 1   'Skip tens position
                ElseIf nDigit > 0 Then
                    tmpBuff = ones(nDigit) & " "
                Else
                    'If next 10s & 100s columns are also
                    'zero, then don't show 'thousands'
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    tmpBuff = tmpBuff & thousands(nPosition / 3) & " "
                End If
                strResult = tmpBuff & strResult
            Case 2  'Tens position
                If nDigit > 0 Then
                    strResult = tens(nDigit) & " " & strResult
                End If
            Case 0  'Hundreds position
                If nDigit > 0 Then
                    strResult = ones(nDigit) & " hundred " & strResult
                End If
        End Select
    Next i
    'Convert first letter to upper case
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If

EndNumToText:
    'Return result
    NumToText = Trim(strSign) & strResult
    Exit Function

NumToTextError:
    strResult = "#Error#"
    Resume EndNumToText
End Function


Private Sub OutCCRPC(pRefnum As Long, pCCRNum As Long)
' *************************
' ** Printing of receipt **
' *************************
Dim ctrCnt As Integer
Dim tmp1 As String * 30
Dim tmp2 As String * 30
Dim tmpString As String
Dim Word1 As String * 36
Dim Word2 As String * 36
Dim Word3 As String * 36
Dim strEntry As String * 80
Dim Refn As String * 10
Dim Seqf As String * 10
Dim CCRf As String * 10
Dim DateTime As String
Dim strExporter As String * 30
Dim strSize As String * 4
Dim strCtnnum As String * 12

Dim vslName As String * 10
Dim X As Integer
Dim strRemarks As String * 15
Dim rsCCRDetail As ADODB.Recordset
Dim strRemarkOut As String * 16
Dim remark1 As String
Dim remark3 As String
Dim UserName As String
Dim strValidation As String * 35

Dim sRateCode As String
Dim sRateDescription As String * 29
Dim sRateAmount As Currency
Dim sDays As String
Dim docRefNo As String * 6
Dim sValidUntil As String
Dim sValidUntilText As String
Dim Amount As Currency
Dim TotalAmount As Currency
Dim TotalVatAmount As Currency
Dim TotalTaxAmount As Currency
Dim TotalCheckAmount As Currency
Dim strChqAmt As String
Dim strCshAmt As String
Dim rsCCRPay As ADODB.Recordset
Dim strAdrAmt As String

ctrCnt = 11

    Set rsCCRPay = New ADODB.Recordset
    rsCCRPay.Open "SELECT cusnam, userid From CCRPay WHERE refnum = " & Trim(CStr(pRefnum)), _
            gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    If rsCCRPay.BOF <> True And rsCCRPay.EOF <> True Then
        strExporter = rsCCRPay.Fields("cusnam")
        UserName = Trim(UCase(rsCCRPay.Fields("userid") & "")) & Space(8)
    End If
    rsCCRPay.Close
    
Set rsCCRDetail = New ADODB.Recordset
'rsCCRDetail.Open "SELECT * From CCRdtl WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
'        & " order by itmnum", _
'        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
rsCCRDetail.Open "SELECT * From CCRdtl WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
        & " AND ccrnum = " & Trim(CStr(pCCRNum)) & "" _
        & " order by itmnum", _
        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
If rsCCRDetail.BOF <> True And rsCCRDetail.EOF <> True Then
    With rsCCRDetail
        Select Case .Fields("vatcde")
        Case "1"
            remark1 = "0Vat"
        Case "2"
            remark1 = "1%Less"
        Case "3"
            remark1 = "10%Vat"
        Case "4"
            remark1 = "1%Less"
        Case "5"
            remark1 = "6%Vat"
        Case "6"
            remark1 = "1%Less"
        Case Else
            remark1 = " "
        End Select

        If .Fields("guarntycde") = "Y" Then
            remark3 = "U/G"
        Else
            remark3 = " "
        End If
        'UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
        Refn = Format(.Fields("refnum"), "000000")
        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
        CCRf = Format(.Fields("ccrnum"), "000000")
        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
        
'        strRemarks = Mid(.Fields("remark") & "", 1, 15)
        'strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark3, 1, 3)) & Space(7)
        strEntry = .Fields("entnum")
        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
        vslName = .Fields("vslcde") & ""
        
        Printer.Font = "Courier 12cpi"
'        Printer.Font = "Courier"
        Printer.FontSize = 10

        Printer.Print " "

        Printer.Print Space(74) & DateTime
        Printer.Print " "
        Printer.Print " "
        Printer.Print Space(4) & strExporter & Space(13) & vslName & Space(3) & _
                        Trim(Mid(strEntry, 1, 8) & " " & _
                        Mid(strEntry, 9, 8) & " " & _
                        Mid(strEntry, 17, 8) & " " & _
                        Mid(strEntry, 25, 8) & " " & _
                        Mid(strEntry, 33, 8))
        Printer.Print Space(45) & Trim(Mid(strEntry, 41, 8) & " " & _
                        Mid(strEntry, 49, 8) & " " & _
                        Mid(strEntry, 57, 8) & " " & _
                        Mid(strEntry, 65, 8) & " " & _
                        Mid(strEntry, 73, 8))
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        'Printer.Print Space(2) & .Fields("commod")
        Printer.Print Space(32)

        Do While Not .EOF
            strSize = .Fields("cntsze")
            strCtnnum = .Fields("cntnum")
            sRateDescription = Mid(Trim(.Fields("descr")), 1, 25)
            sDays = ""
            docRefNo = .Fields("docRefNo")
            If .Fields("chargetyp") = "IMST" Then
                If IsNull(.Fields("stordys")) = False And CInt(.Fields("stordys")) <> 0 Then
                    sDays = .Fields("stordys")
                End If
            ElseIf .Fields("chargetyp") = "IMRF" Then
                If IsNull(.Fields("rfrhrs")) = False And CInt(.Fields("rfrhrs")) <> 0 Then
                    sDays = .Fields("rfrhrs")
                End If
            End If
            
        
            If .Fields("chargetyp") <> "IMST" And .Fields("chargetyp") <> "EXST" Then
                sValidUntilText = Space(11)
                sValidUntil = ""
                
            Else
                sValidUntilText = "VALID UNTIL"
                sValidUntil = .Fields("enstodttm") & IIf(.Fields("chargetyp") = "IMRF", .Fields("rfrhrs"), "")
            End If
            
            Amount = CDbl(.Fields("amt")) + CDbl(.Fields("dgramt")) + CDbl(.Fields("ovzamt")) + CDbl(.Fields("vatamt")) - CDbl(.Fields("wtax"))
            'Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
            'sharon 05Nov2009 Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(2) & strWgh & Space(26) & Format(CDbl(strArrastre) + CDbl(strWgh), "###,###,###.#0")
            'printing of container numbers, container size, amount, rate code and rate description
            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & sDays & Space(2) & sRateDescription & Space(2) & docRefNo & Space(2) & sValidUntilText & Space(2); sValidUntil & Space(6) & Format(CDbl(Amount), "###,###,###.#0")
            TotalAmount = TotalAmount + Amount
            TotalVatAmount = TotalVatAmount + CDbl(.Fields("vatamt"))
            TotalTaxAmount = TotalTaxAmount + CDbl(.Fields("wtax"))
            strRemarks = Trim(.Fields("remark"))
            ctrCnt = ctrCnt - 1
            .MoveNext
        Loop
        If ctrCnt > 0 Then
            For X = 1 To ctrCnt
                Printer.Print " "
            Next
        End If
    End With

    If TotalVatAmount > 0 Then
        If TotalTaxAmount > 0 Then
            tmpString = "VAT INCLUSIVE LESS W/TAX"
        Else
            tmpString = "VAT INCLUSIVE"
        End If
    Else
        tmpString = "ZERO RATED VAT"
    End If
    'Printer.Print " "  ' 3
    Printer.Print " "  ' 3
    Printer.Print " "  ' 3
    
    'Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & strTArrastre
    Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & Format(CDbl(TotalAmount), "###,###,###.#0")
    tmpString = NumToText(CCur(TotalAmount))
    Word1 = Mid(tmpString, 1, 35)
    If Len(Trim(Mid(tmpString, 35, 1))) <> 0 And Len(Trim(Mid(tmpString, 36, 1))) <> 0 Then
        Word1 = Trim(Word1) & "-"
    End If
    Word2 = Mid(tmpString, 36, 35)
    If Len(Trim(Mid(tmpString, 71, 1))) <> 0 And Len(Trim(Mid(tmpString, 72, 1))) <> 0 Then
        Word2 = Trim(Word2) & "-"
    End If
    Word3 = Mid(tmpString, 72, 35)
    Printer.Print " "
    Printer.Print " "
    Printer.Print Space(46) & Word1
    Printer.Print Space(2) & strRemarks & Space(6) & strRemarkOut & Space(7) & Word2

'        If DomesticMode Then
'            Printer.Print "  DOMESTIC" & Space(36) & Word3
'        Else
'            Printer.Print "  FOREIGN " & Space(36) & Word3
'        End If
    Printer.Print Space(46) & Word3

    Printer.Print " "


    Set rsCCRPay = New ADODB.Recordset
    rsCCRPay.Open "SELECT * From CCRPay WHERE refnum = " & Trim(CStr(pRefnum)), _
            gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    If rsCCRPay.BOF <> True And rsCCRPay.EOF <> True Then
        With rsCCRPay
        'get the Header/footer data
            If IsNull(.Fields("chkamt1")) = False Then
                TotalCheckAmount = TotalCheckAmount + CCur(.Fields("chkamt1"))
            End If
            If IsNull(.Fields("chkamt2")) = False Then
                TotalCheckAmount = TotalCheckAmount + CCur(.Fields("chkamt2"))
            End If
            If IsNull(.Fields("chkamt3")) = False Then
                TotalCheckAmount = TotalCheckAmount + CCur(.Fields("chkamt3"))
            End If
            If IsNull(.Fields("chkamt4")) = False Then
                TotalCheckAmount = TotalCheckAmount + CCur(.Fields("chkamt4"))
            End If
            If IsNull(.Fields("chkamt5")) = False Then
                TotalCheckAmount = TotalCheckAmount + CCur(.Fields("chkamt5"))
            End If
            
            strChqAmt = Format(TotalCheckAmount, "###,###.00")
            strCshAmt = Format(.Fields("cshamt"), "###,###.00")
            
'            tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
'            Printer.Print Space(44) & tmpString
'
            strAdrAmt = Format(.Fields("adramt"), "###,###.00")
            
            UserName = .Fields("userid")
            
            Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
              
            tmpString = strCshAmt & " CS                  "
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(tmpString)
            Printer.Print tmpString
              
            tmpString = strChqAmt & " CK                  "
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(tmpString)
            Printer.Print tmpString
            
            Dim strAdrnum As String
            strAdrnum = " " & CLng(.Fields("adrnum")) & "                               "
            If CLng(Trim(strAdrnum)) = 0 Then strAdrnum = Space(18)
            strAdrnum = Left(strAdrnum, 18)
            tmpString = strAdrAmt & " AD" & strAdrnum
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(tmpString)
            Printer.Print tmpString
            
'            tmpString = strAdrAmt & " AD"
            
'            UserName = .Fields("userid")
'
'            Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
            
            If IsNull(.Fields("chkamt1")) = False Then
                tmp1 = Trim(.Fields("chkno1"))
            Else
                tmp1 = " "
            End If
            
            If IsNull(.Fields("chkamt2")) = False Then
                If Len(tmp1) > 0 Then
                    tmp2 = ", " & Trim(.Fields("chkno2"))
                Else
                    tmp2 = " " & Trim(.Fields("chkno2"))
                End If
            Else
                tmp2 = " "
            End If
            Printer.Print Space(44) & Trim(tmp1) & tmp2

            If IsNull(.Fields("chkamt3")) = False Then
                tmp1 = Trim(.Fields("chkno3"))
            Else
                tmp1 = " "
            End If
            
            If IsNull(.Fields("chkamt4")) = False Then
                If Len(tmp1) > 0 Then
                    tmp2 = ", " & Trim(.Fields("chkno4"))
                Else
                    tmp2 = " " & Trim(.Fields("chkno4"))
                End If
            Else
                tmp2 = " "
            End If

            Printer.Print Space(44) & Trim(tmp1) & tmp2
            
            If IsNull(.Fields("chkamt5")) = False Then
                tmp1 = Trim(.Fields("chkno5"))
            Else
                tmp1 = " "
            End If

            Printer.Print Space(44) & tmp1
            Printer.Print Space(44) & strValidation
            Printer.Print ""
            Printer.Print ""
            Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        
        End With
        rsCCRPay.Close
        Set rsCCRPay = Nothing
    End If

    
    Printer.FontSize = 10
    Printer.EndDoc
Else
MsgBox "Reference/CCR # not found!", vbExclamation, "Error"

End If

rsCCRDetail.Close
Set rsCCRDetail = Nothing

End Sub


