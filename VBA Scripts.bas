Attribute VB_Name = "Module1"
Sub Combine()

    Dim j As Integer
    Dim s As Worksheet

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add
    Sheets(1).Name = "Combined"

    
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    For Each s In ActiveWorkbook.Sheets
        If s.Name <> "Combined" Then
            Application.GoTo Sheets(s.Name).[a1]
            Selection.CurrentRegion.Select
            Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
            Selection.Copy Destination:=Sheets("Combined"). _
              Cells(Rows.Count, 1).End(xlUp)(2)
        
        End If
    Next
End Sub

Sub FinalReport()

    Dim CurrentWs As Worksheet
    Dim Need_Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Need_Summary_Table_Header = True
    Combined_Spreadsheet = True
    
    For Each CurrentWs In Worksheets
    
        ' Set initial variable for holding the ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        ' Set an initial variable for holding the total per ticker name
        Dim Total_Volume As Double
        Total_Volume = 0
        
        ' Set new variables for Moderate Solution Part
        Dim Open_Amt As Double
        Open_Amt = 0
        Dim Close_Amt As Double
        Close_Amt = 0
        Dim Delta_Amt As Double
        Delta_Amt = 0
        Dim Delta_Percent As Double
        Delta_Pct = 0
        Dim Max_Ticker_Nm As String
        Max_Ticker_Nm = " "
        Dim Min_Ticker_Nm As String
        Min_Ticker_Nm = " "
        Dim Max_Pct As Double
        Max_Pct = 0
        Dim MIN_Pct As Double
        MIN_Pct = 0
        Dim Max_Vol_Ticker As String
        Max_Vol_Ticker = " "
        Dim Max_Vol As Double
        Max_Vol = 0
         
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        If Need_Summary_Table_Header Then
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Pct Change"
            CurrentWs.Range("L1").Value = "Total Stock Vol"
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Vol"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            Need_Summary_Table_Header = True
            
        End If
            Open_Amt = CurrentWs.Cells(2, 3).Value
        
        For i = 2 To Lastrow

            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                Close_Amt = CurrentWs.Cells(i, 6).Value
                Delta_Amt = Close_Amt - Open_Amt
                
                If Open_Amt <> 0 Then
                    Delta_Pct = (Delta_Amt / Open_Amt) * 100
                    
                End If
    
                    Total_Ticker_Vol = Total_Ticker_Vol + CurrentWs.Cells(i, 7).Value
                    CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Amt
                    
                If (Delta_Amt > 0) Then
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf (Delta_Amt <= 0) Then
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                    CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Pct) & "%")
                    CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Vol
                    Summary_Table_Row = Summary_Table_Row + 1
                    Delta_Amt = 0
                    Close_Amt = 0
                    Open_Amt = CurrentWs.Cells(i + 1, 3).Value
              
                If (Delta_Pct > Max_Pct) Then
                    Max_Pct = Delta_Pct
                    Max_Ticker_Nm = Ticker_Name
                ElseIf (Delta_Pct < MIN_Pct) Then
                    MIN_Pct = Delta_Pct
                    Min_Ticker_Nm = Ticker_Name
                End If
                       
                If (Total_Ticker_Vol > Max_Vol) Then
                    Max_Vol = Total_Ticker_Vol
                    Max_Vol_Ticker = Ticker_Name
                End If
                    Delta_Pct = 0
                    Total_Ticker_Vol = 0
                Else
                    Total_Ticker_Vol = Total_Ticker_Vol + CurrentWs.Cells(i, 7).Value
                End If
      
        Next i

            If Not COMMAND_SPREADSHEET Then
            
                CurrentWs.Range("Q2").Value = (CStr(Max_Pct) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_Pct) & "%")
                CurrentWs.Range("P2").Value = Max_Ticker_Nm
                CurrentWs.Range("P3").Value = Min_Ticker_Nm
                CurrentWs.Range("Q4").Value = Max_Vol
                CurrentWs.Range("P4").Value = Max_Vol_Ticker
                
            Else
                COMMAND_SPREADSHEET = True
            End If
        
     Next CurrentWs
End Sub


