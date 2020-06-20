Sub resetData()
'This makro clears all entries made by user
    Range("B7:B17").ClearContents
    Range("E10:E11").ClearContents
End Sub

Sub savings()
'This makro executes the program which counts
'saving potential of a user

Dim income As Long
Dim other_income As Long
Dim housing As Long
Dim tickets As Long
Dim media As Long
Dim streaming As Long
Dim medication As Long
Dim others As Long
Dim remaining As Long
Dim expenses_total As Long
Dim balance As Long
Dim remainingAmount As Long
Dim alert As String
Dim goal As Long
Dim daily As Long
Dim minus As Long


income = InputBox("Please provide your monthly net income in PLN")
other_income = InputBox("Please provide other income you receive in PLN")
housing = InputBox("Please provide your monthly housing expenses (flat, rent) in PLN")
tickets = InputBox("Please provide how much you spend for tickets/vehicles in PLN")
media = InputBox("Please provide the amount you spend for Internet and phone")
streaming = InputBox("Please provide how much you spend for streaming srevices in PLN")
medication = InputBox("Please provide how much you spend for medication in PLN")
others = InputBox("Please assess other monthly expenditures in PLN")
balance = InputBox("Please provide the balance from the previous month (on + or -) in PLN")


Range("B7") = income
Range("B8") = other_income
Range("B9") = housing
Range("B10") = tickets
Range("B11") = media
Range("B12") = streaming
Range("B13") = medication
Range("B14") = others
Range("B15") = WorksheetFunction.Sum(Range("B9:B14"))
Range("B16") = balance
remainingAmount = WorksheetFunction.Sum(Range("B7:B8")) - WorksheetFunction.Sum(Range("B15:B15")) + Range("B16")
Range("B17") = remainingAmount
If remainingAmount < 0 Then
    MsgBox "Sorry, it seems you do not earn enough even to pay for the basic expenditures"
Else
    MsgBox "PLN " & Range("B17"), , "Remaining amount"
End If


goal = InputBox("Please enter the amount you wish to save each month in PLN")
Range("E10") = goal
daily = (Range("B17") - Range("E10")) / 31
Range("E11") = daily


MsgBox "PLN " & Range("E11"), , "Daily allowance"

End Sub


