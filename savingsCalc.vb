Sub resetData()
'This makro clears all entries made by user
    Range("B7:B14").ClearContents
End Sub

Sub savings()
'This makro executes the program which counts
'saving potential of a user

Dim remainingAmount As Long
Dim income As Long
Dim housing As Long
Dim others As Long
Dim remaining As Long
Dim alert As String
Dim goal As Long
Dim daily As Long

income = InputBox("Please provide your monthly net income in PLN")
housing = InputBox("Please provide your monthly housing expenses (flat, rent) in PLN")
others = InputBox("Please assess other monthly expenditures in PLN")


Range("B7") = income
Range("B8") = housing
Range("B9") = others
Range("B10") = housing + others
Range("B11") = income - Range("B10")
remaining = Range("B11")

If remaining < 0 Then
    MsgBox "Sorry, it seems you do not earn enough even to pay for the basic expenditures"
Else
    MsgBox "PLN " & Range("B11"), , "Remaining amount"
End If


goal = InputBox("Please enter the amount you wish to save each month in PLN")
Range("B12") = goal
daily = (Range("B11") - Range("B12")) / 30
Range("B14") = daily


MsgBox "PLN " & Range("B14"), , "Daily allowance"

End Sub
