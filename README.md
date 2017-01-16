<h1>Beschreibung</h1>
<p>Excel2SepaXML ist eine Sammlung von Visual Basic (VBA) Klassen, welche den Zweck haben Sepa XML Dateien aus einer Excel Anwendung heraus zu erstellen.<br> 
Zur Zeit sind folgende Dateien erstellbar:</p>
<ul>
  <li>Sepa Überweisung (SCT)</li>
  <li>Sepa Lastschrift (SDD)<br>
    <ul>
      <li>Servicelevel</li>
        <ul>
          <li>CORE</li>
          <li>B2B</li>
         </ul>
      <li>Sequenztypen</li>
        <ul>
          <li>FRST</li>
          <li>RCUR</li>
          <li>OOFF</li>
        </ul>
      </ul>
   </li>
</ul>
<p>Die Beispieldatei im Verzeichnis "example" bietet eine funktionstüchtige Anwendung, welche es ermöglicht Sepa Überweisungen oder Sepa Lastschriften aus einer Liste zu erstellen. Für weitere Informationen besuchen Sie bitte folgende <a href="https://jansesoft.de/blog/index.php/excel2sepaxml/">Website</a>.</p>
<p>im Allgemeinen benötigte Verweise:</p>
<ul>
  <li>Microsoft XML, v6.0</li>
  <li>Microsoft VBScript Regular Expressions 5.5</li>
  <li>Microsoft Scripting Runtime (Nur zur Ausgabe des Fehlerprotokolls im Beispielprogramm)</li>
</ul>
<h2>Sepa Überweisung (SCT) XML erstellen</h2>
<p>Nachfolgend finden Sie Beispielcode aus der Beispielanwendung um eine Sepa XML SCT Datei zu erstellen</p>
<p>&nbsp;</p>
```
Private Sub create_sepa_xml_SCT()
    Dim x As New clsSepaCCT
    Dim c As New clsSepaCreditInfo
    Dim i As Integer: i = 2
    Dim strErrors As String
    
    With x
        .AusgabePfad = Me.lblPfad
        .BatchBooking = Me.chbBatch
        .UrgendPayment = Me.chbUrgend
        .DebtorAgentBIC = Me.lblBIC
        .DebtorIBAN = Me.lblIBAN
        .DebtorName = Me.lblKontoinhaber
        .ExecutionDate = Me.dtpAusführung
        If Me.chbAutoMsgID Then
            .MessageID = Format(Now, "yyyymmdd-HhNnSs")
        Else
            .MessageID = Me.txtMsgID
        End If
        If Me.chbAutoPymtID Then
            .PaymentID = Format(Now, "yyyymmdd-HhNnSs")
        Else
            .PaymentID = Me.txtPymtID
        End If
        If .check_Values Then
            Exit Sub
        End If
        If .prepare_sepa_xml Then
            Exit Sub
        End If
    End With
    
    Do While Not Worksheets("SEPA_Überweisung").Cells(i, 1) = vbNullString
        With c
            .clear
            .Amount = Worksheets("SEPA_Überweisung").Cells(i, 2)
            .BIC = UCase(Worksheets("SEPA_Überweisung").Cells(i, 3))
            .EndToEndID = Worksheets("SEPA_Überweisung").Cells(i, 6)
            .IBAN = Worksheets("SEPA_Überweisung").Cells(i, 4)
            .Name = Worksheets("SEPA_Überweisung").Cells(i, 1)
            .Verwendungszweck = Worksheets("SEPA_Überweisung").Cells(i, 5)
            If .isErrorOccured Then
                strErrors = strErrors & "Fehler in Zeile " & i & ": " & vbNewLine & _
                    .get_ErrorLog
            Else
                Call x.add_CreditTransferInformation(c)
            End If
        End With
        i = i + 1
    Loop
    
    If Not strErrors = vbNullString Then
        If MsgBox("Es sind Fehler während der Erstellung der SEPA-XML Datei aufgetreten, möchten Sie diese in einem Protokoll speichern?", vbQuestion + vbYesNo, "Fehler beim Erstellen") = vbYes Then
            Dim strPfad As String
            
            strPfad = Me.lblPfad
            strPfad = strPfad & "\Fehlerprotokoll.txt"
            
            Dim fso As New FileSystemObject
            Dim stream As TextStream
            
            Set stream = fso.CreateTextFile(strPfad, True, True)
            Call stream.Write(strErrors)
            
            If MsgBox("Das Fehlerprotokoll wurde unter '" & strPfad & "' abgelegt." & vbCrLf & _
                "Möchten Sie die SEPA-XML-Datei trotzdem erstellen?", vbQuestion + vbYesNo) = vbYes Then
                Call x.create_sepa_xml
            End If
            
            Set stream = Nothing
            Set fso = Nothing
        Else
            If MsgBox("Das Fehlerprotokoll wurde verworfen." & vbCrLf & _
                "Möchten Sie die SEPA-XML-Datei trotzdem erstellen?", vbQuestion + vbYesNo) = vbYes Then
                Call x.create_sepa_xml
            End If
        End If
    Else
        Call x.create_sepa_xml
    End If
    
    Set c = Nothing
    Set x = Nothing
End Sub
```
<p>&nbsp;</p>
<h3>Sepa Lastschrift (SDD) XML erstellen</h3>
<p>Nachfolgend finden Sie Beispielcode aus der Beispielanwendung um eine Sepa XML SDD Datei zu erstellen</p>
<p>&nbsp;</p>
```
Private Sub create_sepa_xml_SDD()  
    Dim x As New clsSepaCDD
    Dim c As New clsSepaDebitInfo
    Dim i As Integer: i = 2
    Dim strErrors As String
    
    With x
        .AusgabePfad = Me.lblPfad
        .CollectionDate = Me.dtpAusführung
        .KreditorAgentBIC = Me.lblBIC
        .KreditorIBAN = Me.lblIBAN
        .KreditorIdentifikation = Me.lblGläubigerID
        .KreditorName = Me.lblKontoinhaber
        If Me.chbAutoMsgID Then
            .MessageID = Format(Now, "yyyymmdd-HhNnSs")
        Else
            .MessageID = Me.txtMsgID
        End If
        If Me.chbAutoPymtID Then
            .PaymentID = Format(Now, "yyyymmdd-HhNnSs")
        Else
            .PaymentID = Me.txtPymtID
        End If
        .SequenceType = Me.cmbSequenz
        .InstrumentCode = Me.cmdArt
        If .check_Values Then
            Exit Sub
        End If
        If .prepare_sepa_xml Then
            Exit Sub
        End If
    End With
    
    Do While Not Worksheets("SEPA_Lastschrift").Cells(i, 1) = vbNullString
        With c
            .clear
            .Amount = Worksheets("SEPA_Lastschrift").Cells(i, 2)
            .BIC = Worksheets("SEPA_Lastschrift").Cells(i, 3)
            .DateOfSignature = Worksheets("SEPA_Lastschrift").Cells(i, 8)
            .EndToEndID = Worksheets("SEPA_Lastschrift").Cells(i, 6)
            .IBAN = Worksheets("SEPA_Lastschrift").Cells(i, 4)
            .MandateID = Worksheets("SEPA_Lastschrift").Cells(i, 7)
            .Name = Worksheets("SEPA_Lastschrift").Cells(i, 1)
            .Verwendungszweck = Worksheets("SEPA_Lastschrift").Cells(i, 5)
            If .isErrorOccured Then
                strErrors = strErrors & "Fehler in Zeile " & i & ": " & vbNewLine & _
                    .get_ErrorLog
            Else
                Call x.add_DebitTransferInformation(c)
            End If
        End With
        i = i + 1
    Loop
    
    If Not strErrors = vbNullString Then
        If MsgBox("Es sind Fehler während der Erstellung der SEPA-XML Datei aufgetreten, möchten Sie diese in einem Protokoll speichern?", vbQuestion + vbYesNo, "Fehler beim Erstellen") = vbYes Then
            Dim strPfad As String
            
            strPfad = Me.lblPfad
            strPfad = strPfad & "\Fehlerprotokoll.txt"
            
            Dim fso As New FileSystemObject
            Dim stream As TextStream
            
            Set stream = fso.CreateTextFile(strPfad, True, True)
            Call stream.Write(strErrors)
            
            If MsgBox("Das Fehlerprotokoll wurde unter '" & strPfad & "' abgelegt." & vbCrLf & _
                "Möchten Sie die SEPA-XML-Datei trotzdem erstellen?", vbQuestion + vbYesNo) = vbYes Then
                Call x.create_sepa_xml
            End If
            
            Set stream = Nothing
            Set fso = Nothing
        Else
            If MsgBox("Das Fehlerprotokoll wurde verworfen." & vbCrLf & _
                "Möchten Sie die SEPA-XML-Datei trotzdem erstellen?", vbQuestion + vbYesNo) = vbYes Then
                Call x.create_sepa_xml
            End If
        End If
    Else
        Call x.create_sepa_xml
    End If
    
    Set c = Nothing
    Set x = Nothing
End Sub
```
<p>&nbsp;</p>
