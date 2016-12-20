VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSepaCreditInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strEndToEndId As String, strInstdAmt As Currency, strFinInstnIdBIC As String, strCdtrNm As String
Dim strCdtrAcctId As String, strPurpCd As String, strRmtInf As String
Dim ErrorLog As clsErrorLog

Private Sub Class_Initialize()
    Set ErrorLog = New clsErrorLog
End Sub

Public Property Get isErrorOccured() As Boolean
    isErrorOccured = ErrorLog.Count
End Property

Public Property Get get_ErrorLog() As String
    get_ErrorLog = ErrorLog.get_Log
End Property

Public Property Get get_EndToEndID() As String
    If strEndToEndId = vbNullString Then
        get_EndToEndID = "NOTPROVIDED"
    Else
        get_EndToEndID = strEndToEndId
    End If
End Property

Public Property Let EndToEndID(ByVal vNewValue As String)
    If Not vNewValue = vbNullString Then
        If Len(vNewValue) > 35 Then
            Call ErrorLog.add("SepaCreditInfo", "Die End-To-End Referenz darf maximal 35 Zeichen lang sein.")
        Else
            strEndToEndId = vNewValue
        End If
    End If
End Property

Public Property Get get_Amount() As Currency
    get_Amount = strInstdAmt
End Property

Public Property Let Amount(ByVal vNewValue As Currency)
    If vNewValue = 0 Then
        Call ErrorLog.add("SepaCreditInfo", "Der Betrag darf nicht 0,00€ sein.")
    ElseIf vNewValue < 0 Then
        Call ErrorLog.add("SepaCreditInfo", "Der Betrag darf nicht kleiner 0,00€ sein.")
    Else
        strInstdAmt = vNewValue
    End If
End Property

Public Property Get get_BIC() As String
    get_BIC = strFinInstnIdBIC
End Property

Public Property Let BIC(ByVal vNewValue As String)
    If vNewValue = vbNullString Then
        Exit Property
    End If
    
    Dim bank As New clsBankData
    
    If Not bank.isValidBIC(vNewValue) Then
        Call ErrorLog.add("SepaCreditInfo", "Die eingegebene BIC entspricht nicht dem allgemeinen Muster.")
    Else
        strFinInstnIdBIC = vNewValue
    End If
    
    Set bank = Nothing
End Property

Public Property Get get_Name() As String
    get_Name = strCdtrNm
End Property

Public Property Let Name(ByVal vNewValue As String)
    If vNewValue = vbNullString Then
        Call ErrorLog.add("SepaCreditInfo", "Der Name darf nicht leer sein.")
    Else
        strCdtrNm = vNewValue
    End If
End Property

Public Property Get get_IBAN() As String
    get_IBAN = strCdtrAcctId
End Property

Public Property Let IBAN(ByVal vNewValue As String)
    Dim bank As New clsBankData
    
    If Not bank.isValidIBAN(vNewValue) Then
        Call ErrorLog.add("SepaCreditInfo", "Die eingegebene IBAN ist ungültig.")
    Else
        strCdtrAcctId = Trim(vNewValue)
    End If
    
    Set bank = Nothing
End Property

Public Property Get get_Purpose() As String
    get_Purpose = strPurpCd
End Property

Public Property Let Purpose(ByVal vNewValue As String)
    strPurpCd = vNewValue
End Property

Public Property Get get_Verwendungszweck() As String
    get_Verwendungszweck = strRmtInf
End Property

Public Property Let Verwendungszweck(ByVal vNewValue As String)
    If vNewValue = vbNullString Then
        Call ErrorLog.add("SepaCreditInfo", "Der Verwendungszweck darf nicht leer sein.")
    ElseIf Len(vNewValue) > 140 Then
        Call ErrorLog.add("SepaCreditInfo", "Der Verwendungszweck darf nicht mehr als 140 Zeichen lang sein.")
    Else
        strRmtInf = vNewValue
    End If
End Property

Public Sub clear()
    strEndToEndId = vbNullString
    strInstdAmt = 0
    strFinInstnIdBIC = vbNullString
    strCdtrNm = vbNullString
    strCdtrAcctId = vbNullString
    strPurpCd = vbNullString
    strRmtInf = vbNullString
    Set ErrorLog = Nothing
    Set ErrorLog = New clsErrorLog
End Sub

Private Sub Class_Terminate()
    Set ErrorLog = Nothing
End Sub
