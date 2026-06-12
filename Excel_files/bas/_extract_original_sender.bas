Private Function ExtractOriginalSender(ByVal bodyText As String) As String
    ' Zweck: Extrahiert ursprünglichen Absender aus weitergeleiteten E-Mails.
    ' Sucht nach typischen Weiterleitungs-Markierungen und extrahiert die Original-E-Mail-Adresse.
    ' Rückgabe: E-Mail-Adresse des ursprünglichen Absenders oder leerer String.
    Dim lines() As String
    Dim i As Long, lineText As String
    Dim originalFrom As String
    Dim atPos As Long, startPos As Long, endPos As Long
    Dim ch As String
    
    lines = Split(bodyText, vbLf)
    
    ' Durchsuche Body nach Weiterleitungs-Markierungen
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        
        ' Deutsche Weiterleitungsmarkierungen: "Von:", "Gesendet von:"
        ' Englische: "From:", "Sent by:"
        ' Outlook: "-----Original Message-----" (nachfolgende Von:/From: Zeilen)
        If (InStr(1, lineText, "Von:", vbTextCompare) = 1 Or _
            InStr(1, lineText, "From:", vbTextCompare) = 1 Or _
            InStr(1, lineText, "Gesendet von:", vbTextCompare) = 1 Or _
            InStr(1, lineText, "Sent by:", vbTextCompare) = 1) And _
            InStr(lineText, "@") > 0 Then
            
            ' Extrahiere E-Mail-Adresse aus dieser Zeile
            ' Format: "Von: Name <email@domain.com>" oder "From: email@domain.com"
            atPos = InStr(lineText, "@")
            If atPos > 0 Then
                ' Suche Start der E-Mail (Zeichen vor @)
                startPos = atPos
                Do While startPos > 1
                    ch = Mid$(lineText, startPos - 1, 1)
                    If ch Like "[A-Za-z0-9._+-]" Then
                        startPos = startPos - 1
                    Else
                        Exit Do
                    End If
                Loop
                
                ' Suche Ende der E-Mail (Zeichen nach @)
                endPos = atPos
                Do While endPos < Len(lineText)
                    ch = Mid$(lineText, endPos + 1, 1)
                    If ch Like "[A-Za-z0-9._-]" Then
                        endPos = endPos + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                originalFrom = Mid$(lineText, startPos, endPos - startPos + 1)
                If Len(originalFrom) > 0 And InStr(originalFrom, "@") > 0 Then
                    ExtractOriginalSender = originalFrom
                    Exit Function
                End If
            End If
        End If
        
        ' Alternativ: Suche nach "Am ... schrieb" Pattern (Apple Mail)
        ' Format: "Am 25.02.2026 um 14:30 schrieb Name <email@domain.com>:"
        If InStr(1, lineText, "schrieb", vbTextCompare) > 0 And InStr(lineText, "@") > 0 Then
            atPos = InStr(lineText, "@")
            If atPos > 0 Then
                startPos = atPos
                Do While startPos > 1
                    ch = Mid$(lineText, startPos - 1, 1)
                    If ch Like "[A-Za-z0-9._+-]" Then
                        startPos = startPos - 1
                    Else
                        Exit Do
                    End If
                Loop
                
                endPos = atPos
                Do While endPos < Len(lineText)
                    ch = Mid$(lineText, endPos + 1, 1)
                    If ch Like "[A-Za-z0-9._-]" Then
                        endPos = endPos + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                originalFrom = Mid$(lineText, startPos, endPos - startPos + 1)
                If Len(originalFrom) > 0 And InStr(originalFrom, "@") > 0 Then
                    ExtractOriginalSender = originalFrom
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ExtractOriginalSender = vbNullString
End Function
