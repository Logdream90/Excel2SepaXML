VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSepaCDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/Konstanten
'//GroupHeader
Private Const XMLNS As String = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"
Private Const XMLNSXSI As String = "http://www.w3.org/2001/XMLSchema-instance"
Private Const XSISCHEMELOCATION As String = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02 pain.008.001.02.xsd"
'//PaymentInformation
'///PaymentMethod
Private Const PMTMTD As String = "DD"
'///ServiceLevel
Private Const SVCLVL As String = "SEPA"
'///ChargeBearer
Private Const CHRGBR As String = "SLEV"

'/Datum
Dim datStart As Date

'/XML Objekte
'//GroupHeader Doc, PI, Root und Initiation
Dim xmlDoc As MSXML2.DOMDocument60
Dim xmlPI As MSXML2.IXMLDOMProcessingInstruction
Dim xmlRoot As MSXML2.IXMLDOMNode
Dim xmlNodCstmrDrctDbtInitn As MSXML2.IXMLDOMNode

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
Dim xmlNodPmtInfNbOfTxs As MSXML2.IXMLDOMNode
Dim xmlNodPmtInfCtrlSum As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInf As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfSvcLvl As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfSvcLvlCd As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfLclInstrm As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfLclInstrmCd As MSXML2.IXMLDOMNode
Dim xmlNodPmtTpInfSeqTp As MSXML2.IXMLDOMNode
Dim xmlNodReqdColltnDt As MSXML2.IXMLDOMNode
Dim xmlNodCdtr As MSXML2.IXMLDOMNode
Dim xmlNodCdtrNm As MSXML2.IXMLDOMNode
Dim xmlNodCdtrAcct As MSXML2.IXMLDOMNode
Dim xmlNodCdtrAcctId As MSXML2.IXMLDOMNode
Dim xmlNodCdtrAcctIdIBAN As MSXML2.IXMLDOMNode
Dim xmlNodCdtrAgt As MSXML2.IXMLDOMNode
Dim xmlNodCdtrAgtFinInstnId As MSXML2.IXMLDOMNode
Dim xmlNodChrgBr As MSXML2.IXMLDOMNode

'//CreditorIdentification
Dim xmlNodCdtrSchmeId As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdId As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdIdPrvtId As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdIdPrvtIdOthr As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdIdPrvtIdOthrId As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNm As MSXML2.IXMLDOMNode
Dim xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNmPrtry As MSXML2.IXMLDOMNode

'/Zähler und Kontrollsumme
Dim intAnzahl As Integer, curSumme As Currency

'/Daten der SEPA-Datei
Dim strMsgId As String, strInitgPtyNm As String, strPmtInfId As String, strPmtTpInfLclInstrmCd As String, strPmtTpInfSeqTp As String
Dim strReqdColltnDt As String, strCdtrNm As String, strCdtrAcctIdIBAN As String, strCdtrAgtBIC As String, strCdtrSchmeIdIdPrvtIdOthrId As String

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

Public Property Let InstrumentCode(ByVal vNewValue As String)
    Select Case vNewValue
        Case "CORE", "B2B":
            strPmtTpInfLclInstrmCd = UCase(vNewValue)
        Case Else:
            MsgBox "InstrumentCode " & vNewValue & " nicht unterstützt."
    End Select
End Property

Public Property Let SequenceType(ByVal vNewValue As String)
    Select Case vNewValue
        'FRST = Erster SEPA Lastschrifteinzug bei wiederkehrende Lastschriften
        'RCUR = Folge SEPA Lastschrifteinzug bei wiederkehrende Lastschriften
        'OOFF = einmalige SEPA Lastschrift
        Case "FRST", "RCUR", "OOFF":
            strPmtTpInfSeqTp = UCase(vNewValue)
        Case Else:
            MsgBox "Sequenztyp " & vNewValue & " nicht unterstützt."
    End Select
End Property

Public Property Let CollectionDate(ByVal vNewValue As Date)
    strReqdColltnDt = Format(vNewValue, "yyyy-mm-dd")
End Property

Public Property Let KreditorName(ByVal vNewValue As String)
    strInitgPtyNm = vNewValue
    strCdtrNm = vNewValue
End Property

Public Property Let KreditorIBAN(ByVal vNewValue As String)
    If Not vNewValue = vbNullString Then
        Dim bank As New clsBankData
        If bank.isValidIBAN(vNewValue) Then
            strCdtrAcctIdIBAN = vNewValue
        End If
    End If
    
    Set bank = Nothing
End Property

Public Property Let KreditorAgentBIC(ByVal vNewValue As String)
    If vNewValue = vbNullString Then
        Exit Property
    End If
    
    Dim bank As New clsBankData
    
    If bank.isValidBIC(vNewValue) Then
        strCdtrAgtBIC = UCase(vNewValue)
    Else
        MsgBox "Die BIC kann nicht hinzugefügt werden, da der BIC-Code nicht dem allgemeinen Muster entspricht.", vbCritical
    End If
    
    Set bank = Nothing
End Property

Public Property Let AusgabePfad(ByVal vNewValue As String)
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(vNewValue) Then
        MsgBox "Dieser Pfad existiert nicht.", vbInformation
        Exit Property
    End If
    strPfad = vNewValue
End Property

Public Property Let KreditorIdentifikation(ByVal vNewValue As String)
    If Len(vNewValue) > 35 Then
        MsgBox "Die Kreditorenidentifikation darf maximal 35 Zeichen lang sein.", vbCritical
    Else
        strCdtrSchmeIdIdPrvtIdOthrId = Trim(vNewValue)
    End If
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
    
    Set xmlNodCstmrDrctDbtInitn = xmlDoc.createNode(NODE_ELEMENT, "CstmrDrctDbtInitn", XMLNS)
    Call xmlRoot.appendChild(xmlNodCstmrDrctDbtInitn)
End Sub

Public Function check_Values() As Boolean
    If (strMsgId = vbNullString _
        Or strInitgPtyNm = vbNullString _
        Or strPmtInfId = vbNullString _
        Or strPmtTpInfLclInstrmCd = vbNullString _
        Or strPmtTpInfSeqTp = vbNullString _
        Or strReqdColltnDt = vbNullString _
        Or strCdtrNm = vbNullString _
        Or strCdtrAcctIdIBAN = vbNullString _
        Or strCdtrSchmeIdIdPrvtIdOthrId = vbNullString) Then
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
    Call xmlNodCstmrDrctDbtInitn.appendChild(xmlNodGrpHdr)
    
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
    Call xmlNodCstmrDrctDbtInitn.appendChild(xmlNodPmtInf)
    
    Set xmlNodPmtInfId = xmlDoc.createNode(NODE_ELEMENT, "PmtInfId", XMLNS)
    xmlNodPmtInfId.Text = strPmtInfId
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfId)
    
    Set xmlNodPmtMtd = xmlDoc.createNode(NODE_ELEMENT, "PmtMtd", XMLNS)
    xmlNodPmtMtd.Text = PMTMTD
    Call xmlNodPmtInf.appendChild(xmlNodPmtMtd)
    
    Set xmlNodPmtInfNbOfTxs = xmlDoc.createNode(NODE_ELEMENT, "NbOfTxs", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfNbOfTxs)
    
    Set xmlNodPmtInfCtrlSum = xmlDoc.createNode(NODE_ELEMENT, "CtrlSum", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtInfCtrlSum)
    
    Set xmlNodPmtTpInf = xmlDoc.createNode(NODE_ELEMENT, "PmtTpInf", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodPmtTpInf)
    
    Set xmlNodPmtTpInfSvcLvl = xmlDoc.createNode(NODE_ELEMENT, "SvcLvl", XMLNS)
    Call xmlNodPmtTpInf.appendChild(xmlNodPmtTpInfSvcLvl)
    
    Set xmlNodPmtTpInfSvcLvlCd = xmlDoc.createNode(NODE_ELEMENT, "Cd", XMLNS)
    xmlNodPmtTpInfSvcLvlCd.Text = SVCLVL
    Call xmlNodPmtTpInfSvcLvl.appendChild(xmlNodPmtTpInfSvcLvlCd)
    
    Set xmlNodPmtTpInfLclInstrm = xmlDoc.createNode(NODE_ELEMENT, "LclInstrm", XMLNS)
    Call xmlNodPmtTpInf.appendChild(xmlNodPmtTpInfLclInstrm)
    
    Set xmlNodPmtTpInfLclInstrmCd = xmlDoc.createNode(NODE_ELEMENT, "Cd", XMLNS)
    xmlNodPmtTpInfLclInstrmCd.Text = strPmtTpInfLclInstrmCd
    Call xmlNodPmtTpInfLclInstrm.appendChild(xmlNodPmtTpInfLclInstrmCd)
    
    Set xmlNodPmtTpInfSeqTp = xmlDoc.createNode(NODE_ELEMENT, "SeqTp", XMLNS)
    xmlNodPmtTpInfSeqTp.Text = strPmtTpInfSeqTp
    Call xmlNodPmtTpInf.appendChild(xmlNodPmtTpInfSeqTp)
    
    Set xmlNodReqdColltnDt = xmlDoc.createNode(NODE_ELEMENT, "ReqdColltnDt", XMLNS)
    xmlNodReqdColltnDt.Text = strReqdColltnDt
    Call xmlNodPmtInf.appendChild(xmlNodReqdColltnDt)
    
    Set xmlNodCdtr = xmlDoc.createNode(NODE_ELEMENT, "Cdtr", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodCdtr)
    
    Set xmlNodCdtrNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
    xmlNodCdtrNm.Text = strCdtrNm
    Call xmlNodCdtr.appendChild(xmlNodCdtrNm)
    
    Set xmlNodCdtrAcct = xmlDoc.createNode(NODE_ELEMENT, "CdtrAcct", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodCdtrAcct)
    
    Set xmlNodCdtrAcctId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    Call xmlNodCdtrAcct.appendChild(xmlNodCdtrAcctId)
    
    Set xmlNodCdtrAcctIdIBAN = xmlDoc.createNode(NODE_ELEMENT, "IBAN", XMLNS)
    xmlNodCdtrAcctIdIBAN.Text = strCdtrAcctIdIBAN
    Call xmlNodCdtrAcctId.appendChild(xmlNodCdtrAcctIdIBAN)
    
    Set xmlNodCdtrAgt = xmlDoc.createNode(NODE_ELEMENT, "CdtrAgt", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodCdtrAgt)
    
    Set xmlNodCdtrAgtFinInstnId = xmlDoc.createNode(NODE_ELEMENT, "FinInstnId", XMLNS)
    Call xmlNodCdtrAgt.appendChild(xmlNodCdtrAgtFinInstnId)
    
    If strCdtrAgtBIC = vbNullString Then
        Dim xmlNodCdtrAgtFinInstnIdOthr As MSXML2.IXMLDOMNode
        Dim xmlNodCdtrAgtFinInstnIdOthrId As MSXML2.IXMLDOMNode
        
        Set xmlNodCdtrAgtFinInstnIdOthr = xmlDoc.createNode(NODE_ELEMENT, "Othr", XMLNS)
        Call xmlNodCdtrAgtFinInstnId.appendChild(xmlNodCdtrAgtFinInstnIdOthr)
        
        Set xmlNodCdtrAgtFinInstnIdOthrId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
        xmlNodCdtrAgtFinInstnIdOthrId.Text = "NOTPROVIDED"
        Call xmlNodCdtrAgtFinInstnIdOthr.appendChild(xmlNodCdtrAgtFinInstnIdOthrId)
    Else
        Dim xmlNodCdtrAgtFinInstnIdBIC As MSXML2.IXMLDOMNode
        
        Set xmlNodCdtrAgtFinInstnIdBIC = xmlDoc.createNode(NODE_ELEMENT, "BIC", XMLNS)
        xmlNodCdtrAgtFinInstnIdBIC.Text = strCdtrAgtBIC
        Call xmlNodCdtrAgtFinInstnId.appendChild(xmlNodCdtrAgtFinInstnIdBIC)
    End If
    
    Set xmlNodChrgBr = xmlDoc.createNode(NODE_ELEMENT, "ChrgBr", XMLNS)
    xmlNodChrgBr.Text = CHRGBR
    Call xmlNodPmtInf.appendChild(xmlNodChrgBr)
    
    '//CreditorIdentification
    Set xmlNodCdtrSchmeId = xmlDoc.createNode(NODE_ELEMENT, "CdtrSchmeId", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodCdtrSchmeId)
    
    Set xmlNodCdtrSchmeIdId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    Call xmlNodCdtrSchmeId.appendChild(xmlNodCdtrSchmeIdId)
    
    Set xmlNodCdtrSchmeIdIdPrvtId = xmlDoc.createNode(NODE_ELEMENT, "PrvtId", XMLNS)
    Call xmlNodCdtrSchmeIdId.appendChild(xmlNodCdtrSchmeIdIdPrvtId)
    
    Set xmlNodCdtrSchmeIdIdPrvtIdOthr = xmlDoc.createNode(NODE_ELEMENT, "Othr", XMLNS)
    Call xmlNodCdtrSchmeIdIdPrvtId.appendChild(xmlNodCdtrSchmeIdIdPrvtIdOthr)
    
    Set xmlNodCdtrSchmeIdIdPrvtIdOthrId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    xmlNodCdtrSchmeIdIdPrvtIdOthrId.Text = strCdtrSchmeIdIdPrvtIdOthrId
    Call xmlNodCdtrSchmeIdIdPrvtIdOthr.appendChild(xmlNodCdtrSchmeIdIdPrvtIdOthrId)
    
    Set xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNm = xmlDoc.createNode(NODE_ELEMENT, "SchmeNm", XMLNS)
    Call xmlNodCdtrSchmeIdIdPrvtIdOthr.appendChild(xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNm)
    
    Set xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNmPrtry = xmlDoc.createNode(NODE_ELEMENT, "Prtry", XMLNS)
    xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNmPrtry.Text = SVCLVL
    Call xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNm.appendChild(xmlNodCdtrSchmeIdIdPrvtIdOthrSchmeNmPrtry)
End Function

Public Function add_DebitTransferInformation(ByVal DebitInfo As clsSepaDebitInfo) As Boolean
    If check_DebitTransferInformation(DebitInfo) Then
        add_DebitTransferInformation = True
        Exit Function
    End If
    
    Dim xmlNodDrctDbtTxInf As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfPmtId As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfPmtIdEndToEndId As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfInstdAmt As MSXML2.IXMLDOMNode
    Dim xmlAttDrctDbtTxInfInstdAmt As MSXML2.IXMLDOMAttribute
    Dim xmlNodDrctDbtTxInfDrctDbtTx As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInf As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfMndtId As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfDtOfSgntr As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrAgt As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrAgtFinInstnId As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtr As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrNm As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrAcct As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrAcctId As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfDbtrAcctIdIBAN As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfRmtInf As MSXML2.IXMLDOMNode
    Dim xmlNodDrctDbtTxInfRmtInfUstrd As MSXML2.IXMLDOMNode
    
    Set xmlNodDrctDbtTxInf = xmlDoc.createNode(NODE_ELEMENT, "DrctDbtTxInf", XMLNS)
    Call xmlNodPmtInf.appendChild(xmlNodDrctDbtTxInf)
    
    Set xmlNodDrctDbtTxInfPmtId = xmlDoc.createNode(NODE_ELEMENT, "PmtId", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfPmtId)
    
    Set xmlNodDrctDbtTxInfPmtIdEndToEndId = xmlDoc.createNode(NODE_ELEMENT, "EndToEndId", XMLNS)
    xmlNodDrctDbtTxInfPmtIdEndToEndId.Text = DebitInfo.get_EndToEndID
    Call xmlNodDrctDbtTxInfPmtId.appendChild(xmlNodDrctDbtTxInfPmtIdEndToEndId)
    
    Set xmlNodDrctDbtTxInfInstdAmt = xmlDoc.createNode(NODE_ELEMENT, "InstdAmt", XMLNS)
    xmlNodDrctDbtTxInfInstdAmt.Text = Replace(CStr(DebitInfo.get_Amount), ",", ".")
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfInstdAmt)
    
    Set xmlAttDrctDbtTxInfInstdAmt = xmlDoc.createAttribute("Ccy")
    xmlAttDrctDbtTxInfInstdAmt.Value = "EUR"
    Call xmlNodDrctDbtTxInfInstdAmt.Attributes.setNamedItem(xmlAttDrctDbtTxInfInstdAmt)
    
    Set xmlNodDrctDbtTxInfDrctDbtTx = xmlDoc.createNode(NODE_ELEMENT, "DrctDbtTx", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfDrctDbtTx)
    
    Set xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInf = xmlDoc.createNode(NODE_ELEMENT, "MndtRltdInf", XMLNS)
    Call xmlNodDrctDbtTxInfDrctDbtTx.appendChild(xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInf)
    
    Set xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfMndtId = xmlDoc.createNode(NODE_ELEMENT, "MndtId", XMLNS)
    xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfMndtId.Text = DebitInfo.get_MandateID
    Call xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInf.appendChild(xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfMndtId)
    
    Set xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfDtOfSgntr = xmlDoc.createNode(NODE_ELEMENT, "DtOfSgntr", XMLNS)
    xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfDtOfSgntr.Text = Format(DebitInfo.get_DateOfSignature, "yyyy-mm-dd")
    Call xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInf.appendChild(xmlNodDrctDbtTxInfDrctDbtTxMndtRltdInfDtOfSgntr)
    
    Set xmlNodDrctDbtTxInfDbtrAgt = xmlDoc.createNode(NODE_ELEMENT, "DbtrAgt", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfDbtrAgt)
    
    Set xmlNodDrctDbtTxInfDbtrAgtFinInstnId = xmlDoc.createNode(NODE_ELEMENT, "FinInstnId", XMLNS)
    Call xmlNodDrctDbtTxInfDbtrAgt.appendChild(xmlNodDrctDbtTxInfDbtrAgtFinInstnId)
    
    If DebitInfo.get_BIC = vbNullString Then
        Dim xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthr As MSXML2.IXMLDOMNode
        Dim xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthrId As MSXML2.IXMLDOMNode
        
        Set xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthr = xmlDoc.createNode(NODE_ELEMENT, "Othr", XMLNS)
        Call xmlNodDrctDbtTxInfDbtrAgtFinInstnId.appendChild(xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthr)

        Set xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthrId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
        xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthrId.Text = "NOTPROVIDED"
        Call xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthr.appendChild(xmlNodDrctDbtTxInfDbtrAgtFinInstnIdOthrId)
    Else
        Dim xmlNodDrctDbtTxInfDbtrAgtFinInstnIdBIC As MSXML2.IXMLDOMNode
        
        Set xmlNodDrctDbtTxInfDbtrAgtFinInstnIdBIC = xmlDoc.createNode(NODE_ELEMENT, "BIC", XMLNS)
        xmlNodDrctDbtTxInfDbtrAgtFinInstnIdBIC.Text = DebitInfo.get_BIC
        Call xmlNodDrctDbtTxInfDbtrAgtFinInstnId.appendChild(xmlNodDrctDbtTxInfDbtrAgtFinInstnIdBIC)
    End If
    
    Set xmlNodDrctDbtTxInfDbtr = xmlDoc.createNode(NODE_ELEMENT, "Dbtr", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfDbtr)
    
    Set xmlNodDrctDbtTxInfDbtrNm = xmlDoc.createNode(NODE_ELEMENT, "Nm", XMLNS)
    xmlNodDrctDbtTxInfDbtrNm.Text = DebitInfo.get_Name
    Call xmlNodDrctDbtTxInfDbtr.appendChild(xmlNodDrctDbtTxInfDbtrNm)
    
    Set xmlNodDrctDbtTxInfDbtrAcct = xmlDoc.createNode(NODE_ELEMENT, "DbtrAcct", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfDbtrAcct)

    Set xmlNodDrctDbtTxInfDbtrAcctId = xmlDoc.createNode(NODE_ELEMENT, "Id", XMLNS)
    Call xmlNodDrctDbtTxInfDbtrAcct.appendChild(xmlNodDrctDbtTxInfDbtrAcctId)
    
    Set xmlNodDrctDbtTxInfDbtrAcctIdIBAN = xmlDoc.createNode(NODE_ELEMENT, "IBAN", XMLNS)
    xmlNodDrctDbtTxInfDbtrAcctIdIBAN.Text = DebitInfo.get_IBAN
    Call xmlNodDrctDbtTxInfDbtrAcctId.appendChild(xmlNodDrctDbtTxInfDbtrAcctIdIBAN)

    Set xmlNodDrctDbtTxInfRmtInf = xmlDoc.createNode(NODE_ELEMENT, "RmtInf", XMLNS)
    Call xmlNodDrctDbtTxInf.appendChild(xmlNodDrctDbtTxInfRmtInf)
    
    Set xmlNodDrctDbtTxInfRmtInfUstrd = xmlDoc.createNode(NODE_ELEMENT, "Ustrd", XMLNS)
    xmlNodDrctDbtTxInfRmtInfUstrd.Text = DebitInfo.get_Verwendungszweck
    Call xmlNodDrctDbtTxInfRmtInf.appendChild(xmlNodDrctDbtTxInfRmtInfUstrd)
    
    intAnzahl = intAnzahl + 1
    curSumme = curSumme + DebitInfo.get_Amount
    DebitInfo.clear
End Function

Private Function check_DebitTransferInformation(ByRef DebitInfo As clsSepaDebitInfo)
    With DebitInfo
        If .get_Amount = 0 Then
            check_DebitTransferInformation = True
        ElseIf .get_EndToEndID = vbNullString Then
            check_DebitTransferInformation = True
        ElseIf .get_IBAN = vbNullString Then
            check_DebitTransferInformation = True
        ElseIf .get_Name = vbNullString Then
            check_DebitTransferInformation = True
        ElseIf .get_Verwendungszweck = vbNullString Then
            check_DebitTransferInformation = True
        ElseIf .get_MandateID = vbNullString Then
            check_DebitTransferInformation = True
        ElseIf .get_DateOfSignature = "00:00:00" Then
            check_DebitTransferInformation = True
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
    Call xmlDoc.Save(strPfad & "\CDD_" & strMsgId & "_" & strPmtInfId & ".xml")
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
    strMsgId = vbNullString
    strInitgPtyNm = vbNullString
    strPmtInfId = vbNullString
    strPmtTpInfLclInstrmCd = vbNullString
    strPmtTpInfSeqTp = vbNullString
    strReqdColltnDt = "00:00:00"
    strCdtrNm = vbNullString
    strCdtrAcctIdIBAN = vbNullString
    strCdtrAgtBIC = vbNullString
    strCdtrSchmeIdIdPrvtIdOthrId = vbNullString
End Sub

Private Sub Class_Terminate()
    Set xmlDoc = Nothing
End Sub


