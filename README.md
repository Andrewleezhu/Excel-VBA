# Excel-VBA 阿拉伯数学转英文大写金额的Excel VBA 自定义函数
Public Function SpellNumber(ByVal myNumber, FrontString As String)
Dim Dollars, Cents, Temp
Dim DecimalPlace, Count
ReDim Place(9) As String
 Place(2) = "Thousand "
 Place(3) = "Million "
 Place(4) = "Billion "
 Place(5) = "Trillion "
 myNumber = Trim(Str(myNumber))
 DecimalPlace = InStr(myNumber, ".")
 If DecimalPlace > 0 Then
    Cents = Round((myNumber - Int(myNumber)) * 100, 0) & "/100"
    'Cents = GetTens(Left(Mid(myNumber, DecimalPlace + 1) & "00", 2))
    myNumber = Trim(Left(myNumber, DecimalPlace - 1))
   
 End If
 Count = 1
 Do While myNumber <> ""
    Temp = GetHundreds(Right(myNumber, 3))
    If Temp <> "" Then
      Dollars = Temp & Place(Count) & Dollars
    End If
    If Len(myNumber) > 3 Then
      myNumber = Left(myNumber, Len(myNumber) - 3)
    Else
      myNumber = ""
    End If
    Count = Count + 1
 Loop
 
 Select Case Dollars
    Case ""
      Dollars = FrontString & " Zero "
    Case "one"
      Dollars = FrontString & " One "
    Case Else
      Dollars = FrontString & Dollars
 End Select
 
 Select Case Cents
    Case ""
      Cents = "Only "
    Case "0/100"
      Cents = "Only "
    Case "00/100"
      Cents = "Only "
    Case "one"
      Cents = "and One Cent "
    Case Else
      Cents = "and Cents " & Cents
 End Select
 
 SpellNumber = Dollars & Cents
 'SpellNumber = UCase(SpellNumber)
 
End Function

Private Function GetHundreds(ByVal myNumber)
   Dim result As String
   If Val(myNumber) = 0 Then Exit Function
   myNumber = Right("000" & myNumber, 3)
   If Mid(myNumber, 1, 1) <> "0" Then
     result = GetDigit(Mid(myNumber, 1, 1)) & "Hundred "
   End If
   If Mid(myNumber, 2, 1) <> "0" Then
     result = result & GetTens(Mid(myNumber, 2))
   Else
     result = result & GetDigit(Mid(myNumber, 3))
   End If
   GetHundreds = result
End Function

Private Function GetTens(TensText)
   Dim result As String
   result = ""
   
   If Val(Left(TensText, 1)) = 1 Then
    Select Case Val(TensText)
      Case 10: result = "Ten "
      Case 11: result = "Eleven "
      Case 12: result = "Twelve "
      Case 13: result = "Thirteen "
      Case 14: result = "Fourteen "
      Case 15: result = "Fifteen "
      Case 16: result = "Sixteen "
      Case 17: result = "Seventeen "
      Case 18: result = "Eighteen "
      Case 19: result = "Nineteen "
      Case Else
    End Select
   Else
    Select Case Val(Left(TensText, 1))
      Case 2: result = "Twenty "
      Case 3: result = "Thrity "
      Case 4: result = "Forty "
      Case 5: result = "Fifty "
      Case 6: result = "Sixty "
      Case 7: result = "Seventy "
      Case 8: result = "Eighty "
      Case 9: result = "Ninety "
      Case Else
    End Select
    result = result & GetDigit(Right(TensText, 1))
   End If
   GetTens = result
End Function

Private Function GetDigit(Digit)
    Select Case Val(Digit)
      Case 1: GetDigit = "One "
      Case 2: GetDigit = "Two "
      Case 3: GetDigit = "Three "
      Case 4: GetDigit = "Four "
      Case 5: GetDigit = "Five "
      Case 6: GetDigit = "Six "
      Case 7: GetDigit = "Seven "
      Case 8: GetDigit = "Eight "
      Case 9: GetDigit = "Nine "
      Case Else: GetDigit = " "
    End Select
End Function
