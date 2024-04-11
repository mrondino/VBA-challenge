Attribute VB_Name = "Module1"
Sub Ticker()

'Code below creates the headers for the columns and the descriptions
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Code below declares variables for the different functions
Dim Ticker_Name As String

Dim GPITicker As String

Dim GPDTicker As String

Dim GTVTicker As String

Dim Ticker_Total As Double
Ticker_Total = 0

Dim Lastrow As Long
Lastrow = 0

Dim Open_Total As Double
Open_Total = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Percent_Change1 As String

Dim New_Variable As Double
New_Variable = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Greatest_Percent_Increase As Double
Greatest_Percent_Increase = 0

Dim Greatest_Percent_Increase1 As String

Dim Greatest_Percent_Decrease As Double
Greatest_Percent_Decrease = 0

Dim Greatest_Percent_Decrease1 As String


Dim Greatest_Total_Volume As Double
Greatest_Total_Voume = 0

'Code below finds the dynamic Last Row in the code
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'The loop below
For i = 2 To Lastrow

'Code below checks to see if the ticker has changed
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


    Ticker_Name = Cells(i, 1).Value

    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    Open_Total = Open_Total + Cells(i, 3).Value
        
    New_Variable = New_Variable + (Cells(i, 6).Value - Cells(i, 3).Value)
    
    
    Percent_Change = (New_Variable / Open_Total)
    
    Percent_Change1 = FormatPercent(Percent_Change)
        
        If Percent_Change > Greatest_Percent_Increase Then
        Greatest_Percent_Increase = Percent_Change
        GPI_Ticker = Cells(i, 1).Value
        End If
                
        If Percent_Change < Greatest_Percent_Decrease Then
        Greatest_Percent_Decrease = Percent_Change
        GPD_Ticker = Cells(i, 1).Value
        End If
    
        If Ticker_Total > Greatest_Total_Volume Then
        Greatest_Total_Volume = Ticker_Total
        GTV_Ticker = Cells(i, 1).Value
        End If

        
    Range("I" & Summary_Table_Row).Value = Ticker_Name
    Range("L" & Summary_Table_Row).Value = Ticker_Total
    Range("J" & Summary_Table_Row).Value = New_Variable
    
    

    
        If New_Variable > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
    Range("K" & Summary_Table_Row).Value = Percent_Change1
        If Percent_Change > 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    
   Summary_Table_Row = Summary_Table_Row + 1

    Ticker_Total = 0
    New_Variable = 0
    Open_Total = 0
    

'If the original tickers were the same, then the following code runs
    Else

    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    New_Variable = New_Variable + (Cells(i, 6).Value - Cells(i, 3).Value)
    Open_Total = Open_Total + Cells(i, 3).Value

    End If

Next i

Greatest_Percent_Increase1 = FormatPercent(Greatest_Percent_Increase)

Greatest_Percent_Decrease1 = FormatPercent(Greatest_Percent_Decrease)


Cells(2, 17).Value = Greatest_Percent_Increase1
Cells(2, 16).Value = GPI_Ticker
Cells(3, 17).Value = Greatest_Percent_Decrease1
Cells(3, 16).Value = GPD_Ticker
Cells(4, 17).Value = Greatest_Total_Volume
Cells(4, 16).Value = GTV_Ticker


End Sub


