VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBankData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//REGEX BIC
Private Const REGEXBICPATTERN As String = "[A-Z]{6,6}[A-Z2-9][A-NP-Z0-9]([A-Z0-9]{3,3}){0,1}"
'//REGEX IBAN
Private Const REGEXIBANPATTERN As String = "[A-Z]{2,2}[0-9]{2,2}[a-zA-Z0-9]{1,30}"

Public Function isValidBIC(ByVal BIC As String) As Boolean
    Dim regEx As New RegExp
    regEx.IgnoreCase = True
    regEx.Pattern = REGEXBICPATTERN
    
    isValidBIC = regEx.test(BIC)
    Set regEx = Nothing
End Function

Public Function isValidIBAN(ByVal IBAN As String) As Boolean
    Dim IBANstr As String
    Dim i As Long
    
    IBANstr = Replace(IBAN, " ", "")
    
    Dim regEx As New RegExp
    With regEx
        .IgnoreCase = True
        .Pattern = REGEXIBANPATTERN
        If Not .test(IBAN) Then
            Set regEx = Nothing
            Exit Function
        End If
    End With
    
    Set regEx = Nothing
    
    IBANstr = Mid(IBANstr, 5) & Mid(IBANstr, 1, 4)
    
    For i = 65 To 90
        IBANstr = Replace(IBANstr, Chr(i), i - 64 + 9, , , vbTextCompare)
    Next
    
    isValidIBAN = ModuloBig(IBANstr, 97) = 1
End Function

Private Function ModuloBig(ZahlString As String, Divisor As Long) As Long
    Dim TeilStart As Long
    Dim TeilEnd As Long
    Dim Erg As Long
    Dim Rest As Long
    Dim ErgebnisGanzString As String
    Dim TeilZahl As Long
    
    TeilStart = 1
    TeilEnd = 1
    
    Do While TeilEnd <= Len(ZahlString)
        TeilZahl = Val(Rest & Mid(ZahlString, TeilStart, TeilEnd - TeilStart + 1))
    
        If TeilZahl >= Divisor Then
            Erg = Int(TeilZahl / Divisor)
            Rest = TeilZahl - Erg * Divisor
    
            ErgebnisGanzString = ErgebnisGanzString + CStr(Erg)
    
            TeilStart = TeilEnd + 1
            TeilEnd = TeilStart
        Else
            If ErgebnisGanzString <> "" Then
                ErgebnisGanzString = ErgebnisGanzString & "0"
            End If
            TeilEnd = TeilEnd + 1
        End If
    Loop
    
    If TeilStart <= Len(ZahlString) Then
        Rest = Val(Rest & Mid(ZahlString, TeilStart))
    End If
    
    ' ErgebnisGanzString enthält jetzt das "Ganzzahlige"
    ' Ergebnis ohne Rest, wird z.Z. aber nicht verwendet
    
    ModuloBig = Rest
End Function
