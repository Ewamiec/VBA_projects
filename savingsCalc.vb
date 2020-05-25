Sub resetData()
'This makro clears all entries made by user
    Range("B7:B11").ClearContents
    Range("B14:G14").ClearContents
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

income = InputBox("Please provide your monthly net income in PLN")
housing = InputBox("Please provide your monthly housing expenses (flat, rent) in PLN")
others = InputBox("Please assess other monthly expenditures in PLN")


Range("B7") = income
Range("B8") = housing
Range("B9") = others
Range("B10") = housing + others
Range("B11") = income - Range("B10")
MsgBox "You still have the remaining amount of " & Range("B11"), , " PLN."


End Sub

