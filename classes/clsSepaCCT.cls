VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSepaCCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/Konstanten
'//GroupHeader
Private Const XMLNS As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
Private Const XMLNSXSI As String = "http://www.w3.org/2001/XMLSchema-instance"
Private Const XSISCHEMELOCATION As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03 pain.001.001.03.xsd"
'//PaymentInformation
'///PaymentMethod ("TRF" = Transfer)
Private Const PMTMTD As String = "TRF"
'///InstructedPriority ("HIGH", "NORM")
Private Const INSTRPRTY As String = "NORM"
'///ChargeBearer
Private Const CHRGBR As String = "SLEV"

'/Datum
Dim datStart As Date

'/XML Objekte
'//GroupHeader Doc, PI, Root und Initiation
Dim xmlDoc As MSXML2.DOMDocument60
Dim xmlPI As MSXML2.IXMLDOMProcessingInstruction
Dim xmlRoot As MSXML2.IXMLDOMNode
Dim xmlNodCstmrCdtTrfInitn As MSXML2.IXMLDOMNode

'//GroupHeader Elemente
Dim xmlNodGrpHdr As MSXML2.IXMLDOMNode
Dim xmlNodMsgId As MSXML2.IXMLDOMNode
Dim xmlNodCreDtTm As MSXML2.IXMLDOMNode
Dim xmlNodNbOfTxs As MSXML2.IXMLDOMNode
Dim xmlNodCtrlSum As MSXML2.IXMLDOMNode
Dim xmlNodInitgPty As MSXML2.IXMLDOMNode
Dim xmlNodInitgPtyNm As MSXML2.IXMLDOMNode

'//PaymentInformation Elemente
Dim xmlNodPmtInf As MSXML2.IXMLDOMNode
Dim xmlNodPmtInfId As MSXML2.IXMLDOMNode
Dim xmlNodPmtMtd As MSXML2.IXMLDOMNode
Dim xmlNodBtchBookg As MSXML2.IXMLDOMNode
Dim xmlNodPmtInfNbOfTxs As MSXML2.IXMLDOMNode
Dim xmlNodPmtInfCtrlSum As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInf As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfInstrPrty As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfSvcLvl As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfSvcLvlCd As MSXML2.IXMLDOMNode
Dim xmlNodReqdExctnDt As MSXML2.IXMLDOMNode
Dim xmlNodDbtr As MSXML2.IXMLDOMNode
Dim xmlNodDbtrNm As MSXML2.IXMLDOMNode
Dim xmlNodDbtrAcct As MSXML2.IXMLDOMNode
Dim xmlNodDbtrAcctId As MSXML2.IXMLDOMNode
Dim xmlNodDbtrAcctIdIBAN As MSXML2.IXMLDOMNode
Dim xmlNodDbtrAgt As MSXML2.IXMLDOMNode
Dim xmlNodDbtrAgtFinInstnId As MSXML2.IXMLDOMNode
Dim xmlNodUltmtDbtr As MSXML2.IXMLDOMNode
Dim xmlNodUltmtDbtrNm As MSXML2.IXMLDOMNode
Dim xmlNodChrgBr As MSXML2.IXMLDOMNode

'/Zähler und Kontrollsumme
Dim intAnzahl As Integer, curSumme As Currency

'/Daten der SEPA-Datei
Dim strMsgId As String, strInitgPtyNm As String, strPmtInfId As String, strBtchBookg As String, strSvcLvl As String, strReqdExctnDt As String
Dim strDbtrNm As String, strDbtrAcctIBAN As String, strDbtrAgtBIC As String, strUltmtDbtrNm As String

'/Weitere Daten
Dim strPfad As String, readyState As Boolean

'/Eigenschaften
Public Property Let MessageID(ByVal vNewValue As String)
    If Len(vNewValue) > 35 Then
        MsgBox "Die Message ID darf maximal 35 Zeichen lang sein.", vbCritical
    Else
        strMsgId = Trim(vNewValue)
    End If
End Property

Public Property Let PaymentID(ByVal vNewValue As String)
    If Len(vNewValue) > 35 Then
        MsgBox "Die Payment ID darf maximal 35 Zeichen lang sein.", vbCritical
    Else
        strPmtInfId = Trim(vNewValue)
    End If
End Property

Public Property Let UrgendPayment(ByVal vNewValue As Boolean)
    If vNewValue Then
        strSvcLvl = "URGP"
    Else
        strSvcLvl = "SEPA"
    End If
End Property

Public Property Let BatchBooking(ByVal vNewValue As Boolean)
    If vNewValue Then
        strBtchBookg = "true"
    Else
        strBtchBookg = "false"
    End If
End Property

Public Property Let ExecutionDate(ByVal vNewValue As Date)
    strReqdExctnDt = Format(vNewValue, "yyyy-mm-dd")
End Property

Public Property Let DebtorName(ByVal vNewValue As String)
    strInitgPtyNm = vNewValue
    strDbtrNm = vNewValue
End Property

Public Property Let DebtorIBAN(ByVal vNewValue As String)
    If Not vNewValue = vbNullString Then
        Dim bank As New clsBankData
        If bank.isValidIBAN(vNewValue) Then
            strDbtrAcctIBAN = vNewValue
        End If
    End If
    
    Set bank = Nothing
End Property

Public Property Let DebtorAgentBIC(ByVal vNewValue As String)
    If vNewValue = vbNullString Then
        Exit Property
    End If
    
    Dim bank As New clsBankData
    
    If bank.isValidBIC(vNewValue) Then
        strDbtrAgtBIC = UCase(vNewValue)
    Else
        MsgBox "Die BIC kann nicht hinzugefügt werden, da der BIC-Code nicht dem allgemeinen Muster entspricht.", vbCritical
    End If
    
    Set bank = Nothing
End Property

Public Property Let UltimateDebtorName(ByVal vNewValue As String)
    strUltmtDbtrNm = vNewValue
End Property

Public Property Let AusgabePfad(ByVal vNewValue As String)
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(vNewValue) Then
        MsgBox "Dieser Pfad existiert nicht.", vbInformation
        Exit Property
    End If
    strPfad = vNewValue
End Property

'/Konstruktor
Private Sub Class_Initialize()
    datStart = Now()
    
    '//GroupHeader Doc, PI und Root
    Set xmlDoc = New MSXML2.DOMDocument60
    
    Set xmlPI = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    Call xmlDoc.InsertBefore(xmlPI, xmlDoc.ChildNodes(0))
    
    Set xmlRoot = xmlDoc.createNode(NODE_ELEMENT, "Document", XMLNS)
    Call xmlDoc.appendChild(xmlRoot)
    
    With xmlDoc.DocumentElement
        Call .setAttribute("xmlns", XMLNS)
        Call .setAttribute("xmlns:xsi", XMLNSXSI)
        Call .setAttribute("xsi:schemaLocation", XSISCHEMELOCATION)
    End With
    
    Set xmlNodCstmrCdtTrfInitn = xmlDoc.createNode(NODE_ELEMENT, "CstmrCdtTrfInitn", XMLNS)
    Call xmlRoot.appendChild(xmlNodCstmrCdtTrfInitn)
End Sub

Public Function check_Values() As Boolean
    If (strMsgId = vbNullString _
        Or strInitgPtyNm = vbNullString _
        Or strPmtInfId = vbNullString _
        Or strSvcLvl = vbNullString _
        Or strReqdExctnDt = vbNullString _
        Or strDbtrNm = vbNullString _
        Or strDbtrAcctIBAN = vbNullString) Then
        MsgBox "Es sind nicht alle erforderlichen Daten vorhanden.", vbCritical
        readyState = False
        check_Values = True
    Else
        readyState = True
        check_Values = False
    End If
End Function

Public Function prepare_sepa_xml() As Boolean
    If Not readyState Then
        MsgBox "Bitte führen Sie zuerst die Methode 'check_Values' durch.", vbInformation
        prepare_sepa_xml = True
        Exit Function
    End If
    
    '/GroupHeader Elemente
    Set xmlNodGrpHdr = xmlDoc.createNode(NODE_ELEMENT, "GrpHdr", XMLNS)
    Call xmlNodCstmrCdtTrfInitn.appendChild(xmlNodGrpHdr)
    
    Set xmlNodMsgId = xmlDoc.createNode(NODE_ELEMENT, "MsgId", XMLNS)
    xmlNodMsgId.Text = strMsgId
    Call xmlNodGrpHdr.appendChild(xmlNodMsgId)
    
    Set xmlNodCreDtTm = xmlDoc.createNode(NODE_ELEMENT, "CreDtTm", XMLNS)
    xmlNodCreDtTm.Text = Format(datStart, "yyyy-mm-dd") & "T" & Format(datStart, "Hh:Nn:Ss")
    Call xmlNodGrpHdr.appendChild(xmlNodCreDtTm)
    
    Set xmlNodNbOfTxs = xmlDoc.createNode(NODE_ELEMENT, "NbOfTxs", XMLNS)
    Call xmlNodGrpHdr.appendChild(xmlNodNbOfTxs)
    
    Set xmlNodCtrlSum = xmlDoc.createNode(NODE_ELEMENT, "CtrlSum", XMLNS)
    Call xmlNodGrpHdr.appendChild(xmlNodCtrlSum)
    
    Set xmlNodInitgPty = xmlDoc.createNode(NODE_ELEMENT, "InitgPty", XMLNS)
    Call xmlNodGrpHdr.appendChild(xmlNodInitgPty)
    
    Set xmlNodInitgPtyNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
    xmlNodInitgPtyNm.Text = strInitgPtyNm
    Call xmlNodInitgPty.appendChild(xmlNodInitgPtyNm)
    
    '//PaymentInformation Elemente
    Set xmlNodPmtInf = xmlDoc.createNode(NODE_ELEMENT, "PmtInf", XMLNS)
    Call xmlNodCstmrCdtTrfInitn.appendChild(xmlNodPmtInf)
    
    Set xmlNodPmtInfId = xmlDoc.createNode(NODE_ELEMENT, "PmtInfId", XMLNS)
    xmlNodPmtInfId.Text = strPmtInfId
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfId)
    
    Set xmlNodPmtMtd = xmlDoc.createNode(NODE_ELEMENT, "PmtMtd", XMLNS)
    xmlNodPmtMtd.Text = PMTMTD
    Call xmlNodPmtInf.appendChild(xmlNodPmtMtd)
    
    If Not strBtchBookg = vbNullString Then
        Set xmlNodBtchBookg = xmlDoc.createNode(NODE_ELEMENT, "BtchBookg", XMLNS)
        xmlNodBtchBookg.Text = strBtchBookg
        Call xmlNodPmtInf.appendChild(xmlNodBtchBookg)
    End If
    
    Set xmlNodPmtInfNbOfTxs = xmlDoc.createNode(NODE_ELEMENT, "NbOfTxs", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfNbOfTxs)
    
    Set xmlNodPmtInfCtrlSum = xmlDoc.createNode(NODE_ELEMENT, "CtrlSum", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfCtrlSum)
    
    Set xmlNodPmtTpInf = xmlDoc.createNode(NODE_ELEMENT, "PmtTpInf", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtTpInf)
    
    Set xmlNodPmtTpInfInstrPrty = xmlDoc.createNode(NODE_ELEMENT, "InstrPrty", XMLNS)
    xmlNodPmtTpInfInstrPrty.Text = INSTRPRTY
    Call xmlNodPmtTpInf.appendChild(xmlNodPmtTpInfInstrPrty)
    
    Set xmlNodPmtTpInfSvcLvl = xmlDoc.createNode(NODE_ELEMENT, "SvcLvl", XMLNS)
    Call xmlNodPmtTpInf.appendChild(xmlNodPmtTpInfSvcLvl)
    
    Set xmlNodPmtTpInfSvcLvlCd = xmlDoc.createNode(NODE_ELEMENT, "Cd", XMLNS)
    xmlNodPmtTpInfSvcLvlCd.Text = strSvcLvl
    Call xmlNodPmtTpInfSvcLvl.appendChild(xmlNodPmtTpInfSvcLvlCd)
    
    Set xmlNodReqdExctnDt = xmlDoc.createNode(NODE_ELEMENT, "ReqdExctnDt", XMLNS)
    xmlNodReqdExctnDt.Text = strReqdExctnDt
    Call xmlNodPmtInf.appendChild(xmlNodReqdExctnDt)
    
    Set xmlNodDbtr = xmlDoc.createNode(NODE_ELEMENT, "Dbtr", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodDbtr)
    
    Set xmlNodDbtrNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
    xmlNodDbtrNm.Text = strDbtrNm
    Call xmlNodDbtr.appendChild(xmlNodDbtrNm)
    
    Set xmlNodDbtrAcct = xmlDoc.createNode(NODE_ELEMENT, "DbtrAcct", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodDbtrAcct)
    
    Set xmlNodDbtrAcctId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    Call xmlNodDbtrAcct.appendChild(xmlNodDbtrAcctId)
    
    Set xmlNodDbtrAcctIdIBAN = xmlDoc.createNode(NODE_ELEMENT, "IBAN", XMLNS)
    xmlNodDbtrAcctIdIBAN.Text = strDbtrAcctIBAN
    Call xmlNodDbtrAcctId.appendChild(xmlNodDbtrAcctIdIBAN)
    
    Set xmlNodDbtrAgt = xmlDoc.createNode(NODE_ELEMENT, "DbtrAgt", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodDbtrAgt)
    
    Set xmlNodDbtrAgtFinInstnId = xmlDoc.createNode(NODE_ELEMENT, "FinInstnId", XMLNS)
    Call xmlNodDbtrAgt.appendChild(xmlNodDbtrAgtFinInstnId)
    
    If strDbtrAgtBIC = vbNullString Then
        Dim xmlNodDbtrAgtFinInstnIdOthr As MSXML2.IXMLDOMNode
        Dim xmlNodDbtrAgtFinInstnIdOthrId As MSXML2.IXMLDOMNode
        
        Set xmlNodDbtrAgtFinInstnIdOthr = xmlDoc.createNode(NODE_ELEMENT, "Othr", XMLNS)
        Call xmlNodDbtrAgtFinInstnId.appendChild(xmlNodDbtrAgtFinInstnIdOthr)
        
        Set xmlNodDbtrAgtFinInstnIdOthrId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
        xmlNodDbtrAgtFinInstnIdOthrId.Text = "NOTPROVIDED"
        Call xmlNodDbtrAgtFinInstnIdOthr.appendChild(xmlNodDbtrAgtFinInstnIdOthrId)
    Else
        Dim xmlNodDbtrAgtFinInstnIdBIC As MSXML2.IXMLDOMNode
        
        Set xmlNodDbtrAgtFinInstnIdBIC = xmlDoc.createNode(NODE_ELEMENT, "BIC", XMLNS)
        xmlNodDbtrAgtFinInstnIdBIC.Text = strDbtrAgtBIC
        Call xmlNodDbtrAgtFinInstnId.appendChild(xmlNodDbtrAgtFinInstnIdBIC)
    End If
    
    If Not strUltmtDbtrNm = vbNullString Then
        Set xmlNodUltmtDbtr = xmlDoc.createNode(NODE_ELEMENT, "UltmtDbtr", XMLNS)
        Call xmlNodPmtInf.appendChild(xmlNodUltmtDbtr)
        
        Set xmlNodUltmtDbtrNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
        xmlNodUltmtDbtrNm.Text = strUltmtDbtrNm
        Call xmlNodUltmtDbtr.appendChild(xmlNodUltmtDbtrNm)
    End If
    
    Set xmlNodChrgBr = xmlDoc.createNode(NODE_ELEMENT, "ChrgBr", XMLNS)
    xmlNodChrgBr.Text = CHRGBR
    Call xmlNodPmtInf.appendChild(xmlNodChrgBr)
End Function

Public Function add_CreditTransferInformation(ByRef CreditInfo As clsSepaCreditInfo) As Boolean
    If check_CreditTransferInformation(CreditInfo) Then
        add_CreditTransferInformation = True
        Exit Function
    End If
    
    Dim xmlNodCdtTrfTxInf As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfPmtId As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfPmtIdEndToEndId As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfAmt As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfAmtInstdAmt As MSXML2.IXMLDOMNode
    Dim xmlAttCdtTrfTxInfAmtInstdAmt As MSXML2.IXMLDOMAttribute
    Dim xmlNodCdtTrfTxInfCdtr As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfCdtrNm As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfCdtrAcct As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfCdtrAcctId As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfCdtrAcctIBAN As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfPurp As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfPurpCd As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfRmtInf As MSXML2.IXMLDOMNode
    Dim xmlNodCdtTrfTxInfRmtInfUstrd As MSXML2.IXMLDOMNode
    
    Set xmlNodCdtTrfTxInf = xmlDoc.createNode(NODE_ELEMENT, "CdtTrfTxInf", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodCdtTrfTxInf)
    
    Set xmlNodCdtTrfTxInfPmtId = xmlDoc.createNode(NODE_ELEMENT, "PmtId", XMLNS)
    Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfPmtId)
    
    Set xmlNodCdtTrfTxInfPmtIdEndToEndId = xmlDoc.createNode(NODE_ELEMENT, "EndToEndId", XMLNS)
    xmlNodCdtTrfTxInfPmtIdEndToEndId.Text = CreditInfo.get_EndToEndID
    Call xmlNodCdtTrfTxInfPmtId.appendChild(xmlNodCdtTrfTxInfPmtIdEndToEndId)
    
    Set xmlNodCdtTrfTxInfAmt = xmlDoc.createNode(NODE_ELEMENT, "Amt", XMLNS)
    Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfAmt)
    
    Set xmlNodCdtTrfTxInfAmtInstdAmt = xmlDoc.createNode(NODE_ELEMENT, "InstdAmt", XMLNS)
    xmlNodCdtTrfTxInfAmtInstdAmt.Text = Replace(CStr(CreditInfo.get_Amount), ",", ".")
    Call xmlNodCdtTrfTxInfAmt.appendChild(xmlNodCdtTrfTxInfAmtInstdAmt)
    
    Set xmlAttCdtTrfTxInfAmtInstdAmt = xmlDoc.createAttribute("Ccy")
    xmlAttCdtTrfTxInfAmtInstdAmt.Value = "EUR"
    Call xmlNodCdtTrfTxInfAmtInstdAmt.Attributes.setNamedItem(xmlAttCdtTrfTxInfAmtInstdAmt)
    
    If Not CreditInfo.get_BIC = vbNullString Then
        Dim xmlNodCdtTrfTxInfCdtrAgt As MSXML2.IXMLDOMNode
        Dim xmlNodCdtTrfTxInfCdtrAgtFinInstnId As MSXML2.IXMLDOMNode
        Dim xmlNodCdtTrfTxInfCdtrAgtFinInstnIdBIC As MSXML2.IXMLDOMNode
        
        Set xmlNodCdtTrfTxInfCdtrAgt = xmlDoc.createNode(NODE_ELEMENT, "CdtrAgt", XMLNS)
        Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfCdtrAgt)
        
        Set xmlNodCdtTrfTxInfCdtrAgtFinInstnId = xmlDoc.createNode(NODE_ELEMENT, "FinInstnId", XMLNS)
        Call xmlNodCdtTrfTxInfCdtrAgt.appendChild(xmlNodCdtTrfTxInfCdtrAgtFinInstnId)
        
        Set xmlNodCdtTrfTxInfCdtrAgtFinInstnIdBIC = xmlDoc.createNode(NODE_ELEMENT, "BIC", XMLNS)
        xmlNodCdtTrfTxInfCdtrAgtFinInstnIdBIC.Text = CreditInfo.get_BIC
        Call xmlNodCdtTrfTxInfCdtrAgtFinInstnId.appendChild(xmlNodCdtTrfTxInfCdtrAgtFinInstnIdBIC)
    End If
    
    Set xmlNodCdtTrfTxInfCdtr = xmlDoc.createNode(NODE_ELEMENT, "Cdtr", XMLNS)
    Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfCdtr)
    
    Set xmlNodCdtTrfTxInfCdtrNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
    xmlNodCdtTrfTxInfCdtrNm.Text = CreditInfo.get_Name
    Call xmlNodCdtTrfTxInfCdtr.appendChild(xmlNodCdtTrfTxInfCdtrNm)
    
    Set xmlNodCdtTrfTxInfCdtrAcct = xmlDoc.createNode(NODE_ELEMENT, "CdtrAcct", XMLNS)
    Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfCdtrAcct)
    
    Set xmlNodCdtTrfTxInfCdtrAcctId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    Call xmlNodCdtTrfTxInfCdtrAcct.appendChild(xmlNodCdtTrfTxInfCdtrAcctId)
    
    Set xmlNodCdtTrfTxInfCdtrAcctIBAN = xmlDoc.createNode(NODE_ELEMENT, "IBAN", XMLNS)
    xmlNodCdtTrfTxInfCdtrAcctIBAN.Text = CreditInfo.get_IBAN
    Call xmlNodCdtTrfTxInfCdtrAcctId.appendChild(xmlNodCdtTrfTxInfCdtrAcctIBAN)
    
    If Not CreditInfo.get_Purpose = vbNullString Then
        Set xmlNodCdtTrfTxInfPurp = xmlDoc.createNode(NODE_ELEMENT, "Purp", XMLNS)
        Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfPurp)
        
        Set xmlNodCdtTrfTxInfPurpCd = xmlDoc.createNode(NODE_ELEMENT, "Cd", XMLNS)
        xmlNodCdtTrfTxInfPurpCd.Text = CreditInfo.get_Purpose
        Call xmlNodCdtTrfTxInfPurp.appendChild(xmlNodCdtTrfTxInfPurpCd)
    End If
    
    Set xmlNodCdtTrfTxInfRmtInf = xmlDoc.createNode(NODE_ELEMENT, "RmtInf", XMLNS)
    Call xmlNodCdtTrfTxInf.appendChild(xmlNodCdtTrfTxInfRmtInf)
    
    Set xmlNodCdtTrfTxInfRmtInfUstrd = xmlDoc.createNode(NODE_ELEMENT, "Ustrd", XMLNS)
    xmlNodCdtTrfTxInfRmtInfUstrd.Text = CreditInfo.get_Verwendungszweck
    Call xmlNodCdtTrfTxInfRmtInf.appendChild(xmlNodCdtTrfTxInfRmtInfUstrd)
    
    intAnzahl = intAnzahl + 1
    curSumme = curSumme + CreditInfo.get_Amount
    CreditInfo.clear
End Function

Private Function check_CreditTransferInformation(ByRef CreditInfo As clsSepaCreditInfo)
    With CreditInfo
        If .get_Amount = 0 Then
            check_CreditTransferInformation = True
        ElseIf .get_EndToEndID = vbNullString Then
            check_CreditTransferInformation = True
        ElseIf .get_IBAN = vbNullString Then
            check_CreditTransferInformation = True
        ElseIf .get_Name = vbNullString Then
            check_CreditTransferInformation = True
        ElseIf .get_Verwendungszweck = vbNullString Then
            check_CreditTransferInformation = True
        End If
    End With
End Function

Public Sub create_sepa_xml()
    If intAnzahl = 0 Or curSumme = 0 Then
        MsgBox "SEPA-XML-Datei kann nicht erstellt werden, da keine Einzelinformationen vorhanden sind.", vbInformation
        Exit Sub
    ElseIf Not readyState Then
        MsgBox "SEPA-XML-Datei kann nicht erstellt werden, da die Grundinformationen nicht vorhanden oder noch nicht überprüft wurden.", vbInformation
        Exit Sub
    End If
    
    xmlNodNbOfTxs.Text = CStr(intAnzahl)
    xmlNodCtrlSum.Text = Replace(CStr(curSumme), ",", ".")
    xmlNodPmtInfNbOfTxs.Text = CStr(intAnzahl)
    xmlNodPmtInfCtrlSum.Text = Replace(CStr(curSumme), ",", ".")
    
    On Error GoTo errorhandler
    If strPfad = vbNullString Then
        strPfad = ThisWorkbook.Path
    End If
    Call xmlDoc.Save(strPfad & "\CCT_" & strMsgId & "_" & strPmtInfId & ".xml")
    On Error GoTo 0
    
    MsgBox "Die SEPA-XML-Datei wurde erfolgreich erstellt." & vbCrLf & _
        "Anzahl Datensätze: " & intAnzahl & vbCrLf & _
        "Summe Datensätze: " & Format(curSumme, "#,##0.00€"), vbInformation, "Datei erstellt"
    
    readyState = False
    clear_variables
    Exit Sub
    
errorhandler:
    MsgBox "Es ist leider ein Fehler beim erstellen der SEPA-XML-Datei aufgetreten." & vbCrLf & _
        "Fehlernummer: " & Err.Number & vbCrLf & _
        "Fehlerbeschreibung: " & Err.Description & _
        "Ausgabepfad: " & strPfad, vbCritical
End Sub

Private Sub clear_variables()
    intAnzahl = 0
    curSumme = 0
    strPfad = vbNullString
    strMsgId = vbNullString
    strInitgPtyNm = vbNullString
    strPmtInfId = vbNullString
    strBtchBookg = vbNullString
    strSvcLvl = vbNullString
    strReqdExctnDt = vbNullString
    strDbtrNm = vbNullString
    strDbtrAcctIBAN = vbNullString
    strDbtrAgtBIC = vbNullString
    strUltmtDbtrNm = vbNullString
End Sub

Private Sub Class_Terminate()
    Set xmlDoc = Nothing
End Sub


