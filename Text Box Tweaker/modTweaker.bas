Attribute VB_Name = "modTweaker"
'*******************************************************'
'                                                       '
'   By:         Waleed A. Aly                           '
'   ASL:        [21 M Egypt]                            '
'   eMail:      wa_aly@tdcspace.dk                      '
'   On:         22 Sep, 2002                            '
'                                                       '
'     Please eMail me any Comments and|or Suggestions.  '
'   I hope you like my work and think is usefull !  :)  '
'   I'd love to know how many people are using my Code  '
'   so you can always eMail me if you are goin' to use  '
'   it :)                                               '
'                                      Thanks.          '
'                                                       '
'*******************************************************'

Public Enum eTweakMode
    Normal
    AllLetters
    AllLettersAllCaps
    AllLettersAllSmall
    AlphaNumeric
    AlphaNumericAllCaps
    AlphaNumericAllSmall
    IntegerPositive
    IntegerAllowNegative
    DecimalPositive
    DecimalAllowNegative
    CashPositive
    CashAllowNegative
    PhoneNumber
End Enum

Public Sub Tweak(txt As TextBox, KeyAscii As Integer, Mode As eTweakMode, Optional iDecimalPlaces As Integer = 2, Optional sBannedSet As String, Optional sAllowedSet As String)

    Dim CH As String
    Dim CurPos As Integer
    
    CH = Chr(KeyAscii)
    CurPos = txt.SelStart
    
    'Accept BackSpaces
    If KeyAscii = 8 Then Exit Sub
    'Accept Characters contained within the sAllowedSet string
    If InStr(1, sAllowedSet, CH) <> 0 Then Exit Sub
    'Deny Characters contained within the sBanned string
    If InStr(1, sBannedSet, CH) <> 0 Then GoTo Skip
    
    Select Case Mode
        Case Normal
            Exit Sub
        Case AllLetters
            If IsCAPS(KeyAscii) Or IsSmall(KeyAscii) Then Exit Sub
        Case AllLettersAllCaps
            If IsCAPS(KeyAscii) Then Exit Sub
            If IsSmall(KeyAscii) Then KeyAscii = KeyAscii - 32: Exit Sub
        Case AllLettersAllSmall
            If IsSmall(KeyAscii) Then Exit Sub
            If IsCAPS(KeyAscii) Then KeyAscii = KeyAscii + 32: Exit Sub
        Case AlphaNumeric
            If IsCAPS(KeyAscii) Or IsSmall(KeyAscii) Or IsDigit(KeyAscii) Then Exit Sub
        Case AlphaNumericAllCaps
            If IsCAPS(KeyAscii) Or IsDigit(KeyAscii) Then Exit Sub
            If IsSmall(KeyAscii) Then KeyAscii = KeyAscii - 32: Exit Sub
        Case AlphaNumericAllSmall
            If IsSmall(KeyAscii) Or IsDigit(KeyAscii) Then Exit Sub
            If IsCAPS(KeyAscii) Then KeyAscii = KeyAscii + 32: Exit Sub
        Case IntegerPositive
            If IsDigit(KeyAscii) Then Exit Sub
        Case IntegerAllowNegative
            If IsDigit(KeyAscii) Then Exit Sub
            If CH = "+" Or CH = "-" Then GoTo ToggleSign
        Case DecimalPositive
            If IsDigit(KeyAscii) Then GoTo CheckDecimalPoint
            If CH = "." And InStr(1, txt, ".") = 0 Then Exit Sub
        Case DecimalAllowNegative
            If IsDigit(KeyAscii) Then GoTo CheckDecimalPoint
            If CH = "." And InStr(1, txt, ".") = 0 Then Exit Sub
            If CH = "+" Or CH = "-" Then GoTo ToggleSign
        Case CashPositive
            If IsDigit(KeyAscii) Then GoTo CheckCashDecimalPoint
            If CH = "." And InStr(1, txt, ".") = 0 Then Exit Sub
        Case CashAllowNegative
            If IsDigit(KeyAscii) Then GoTo CheckCashDecimalPoint
            If CH = "." And InStr(1, txt, ".") = 0 Then Exit Sub
            If CH = "+" Or CH = "-" Then GoTo ToggleSign
        Case PhoneNumber
            If IsDigit(KeyAscii) Then Exit Sub
            If CH = "+" Then GoTo ToggleSign
            If CH = "-" And (Not Doubled(txt.Text, "-")) Then Exit Sub
            If CH = " " And (Not Doubled(txt.Text, " ")) Then Exit Sub
    End Select
    
    GoTo Skip

CheckCashDecimalPoint:
    If InStr(1, Left(txt, txt.SelStart), ".") = 0 Then
        Exit Sub
    Else
        If (Len(txt) - InStr(1, txt, ".")) < iDecimalPlaces Then
            If (Len(txt) - InStr(1, txt, ".")) = iDecimalPlaces - 1 Then KeyAscii = Asc(5 * (Val(CH) \ 5))
            Exit Sub
        End If
    End If
    GoTo Skip

CheckDecimalPoint:
    If InStr(1, Left(txt, txt.SelStart), ".") = 0 Then
        Exit Sub
    Else
        If (Len(txt) - InStr(1, txt, ".")) < iDecimalPlaces Then Exit Sub
    End If
    GoTo Skip

ToggleSign:
    If Left(txt, 1) = "+" Or Left(txt, 1) = "-" Then
        txt = CH & Right(txt, Len(txt) - 1)
        txt.SelStart = CurPos
    Else
        txt = CH & txt
        txt.SelStart = CurPos + 1
    End If
    GoTo Skip

Skip:
    KeyAscii = 0
End Sub

Private Function IsCAPS(KeyAscii As Integer) As Boolean
    If KeyAscii > 64 And KeyAscii < 91 Then IsCAPS = True
End Function

Private Function IsSmall(KeyAscii As Integer) As Boolean
    If KeyAscii > 96 And KeyAscii < 123 Then IsSmall = True
End Function

Private Function IsDigit(KeyAscii As Integer) As Boolean
    If KeyAscii > 47 And KeyAscii < 58 Then IsDigit = True
End Function

Private Function Doubled(S1 As String, S2 As String) As Boolean
    If Right(S1, 1) = S2 Then Doubled = True
End Function
