Sub make_text()

Dim rng As Range
Dim txt As String
Dim x As Long

'this simple macro changes the display output of the cell from scieintific notation
'to text. please note that this only works if the value in the cell has not been
'already truncated by the scientific notation.This is upto 1,000,000,000,000,000 [10^15]
'it re writes the content of the cell after changing to text format to reintroduce the value as text.

'to use the macro, simply select the cells with number values you want to change to text, then run the macro.

'CopyRight Manolo Ariza. manolo.ar.cu@gmail.com
'licence GPLv2

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


x = 1
Set rng = Selection
rng.NumberFormat = "@" 'format as text the cell


        Do While x <= rng.Cells.Count 'loop throu all the cells
        
        If IsEmpty(rng.Cells(x).Value) Then
        
        Else
        
        txt = rng.Cells(x).Value2 'copy the value of the cell to a dummy variable
        rng(x).Value2 = txt 'rewrite the value of the cell as text
        
        End If
        
        x = x + 1
        
        Loop
        
        
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub
