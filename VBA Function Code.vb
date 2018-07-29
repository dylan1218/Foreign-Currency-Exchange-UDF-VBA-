Public Function CurrencyTranslate(DateOfExchange As String, BaseCurrency As String, ExchangedCurrency As String) As String

    Dim ExchangedCurrencyPos As Integer
    Dim EndOfStringPos As Integer
    Dim LengthOfExchangeRate As Integer


' Refer to http://api.fixer.io/ regarding documentation of the API data.
' Dylan Levine
    DateOfExchange = Format(DateOfExchange, "yyyy-mm-dd") 'Formats a date into the format required by the Fixer.IO API

    TextReturn = Application.WorksheetFunction.WebService("http://api.fixer.io/" & DateOfExchange & "?base=" & BaseCurrency & "&symbols=" & ExchangedCurrency)
    ExchangedCurrencyPos = InStr(1, TextReturn, ExchangedCurrency) + 5 'Finds the currency within the returned string from the API, and finds the end of the point where the translation begins
    EndOfStringPos = InStr(1, TextReturn, "}") 'Finsds the numbers of characters at which the string ends
    LengthOfExchangeRate = EndOfStringPos - ExchangedCurrencyPos 'Finds the length of the exchange rate based off of the begining, and ending variables stated above

    If UCase(BaseCurrency) = UCase(ExchangedCurrency) Then
        CurrencyTranslate = "1.00"
        Else
        CurrencyTranslate = Mid(TextReturn, ExchangedCurrencyPos, LengthOfExchangeRate) 'Defines the Function and finds the exchange rate wtihin the given API Json string return
    End If

End Function

