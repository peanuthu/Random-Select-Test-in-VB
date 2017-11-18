VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Dim cell As Range
'counter counts how many valid cells in the column
Dim counter As Integer

'The loop here decides how many people can work on the day.
For Each cell In Selection
    If IsEmpty(cell) = True Then
        Exit For
    Else
        counter = counter + 1
    End If
Next cell

'The step below takes valid column into an array.
Dim monArr() As String
ReDim monArr(counter)
For i = 0 To (counter - 1)
    monArr(i) = Selection(i)
Next

'This step below makes shifts
' monShift() string stores the people scheduled to work on the day
Dim monShift() As String
If counter <= 3 Then
    ReDim monShift(counter)
    For i = 0 To (counter - 1)
        monShift(i) = monArr(i + 1)
        MsgBox monShift(i)
    Next
    
Else
    ReDim monShift(3)
    Dim randDraw As Integer
    
    For i = 0 To 2
    
        randDraw = Int((counter - 0 + 1) * Rnd + 0)
        
        If monArr(randDraw) <> "" Then
            monShift(i) = monArr(randDraw)
            monArr(randDraw) = ""
            
        Else
            While monArr(randDraw) = ""
                randDraw = Int((counter - 0 + 1) * Rnd + 0)
            Wend 'to end While
            monShift(i) = monArr(randDraw)
        
        End If
        MsgBox monShift(i)

    Next

End If

    
End Sub
