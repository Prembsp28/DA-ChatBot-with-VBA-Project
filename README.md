# DA-ChatBot-with-VBA
Developed a Sales Data Chatbot using Excel VBA to automate sales data analysis and reporting. The chatbot answers dynamic business questions related to sales, quantity, targets, status, products, regions, and order details. Implemented VBA automation, conditional logic, loops, and dynamic search functionality for instant insights retrieval.

## Features of the ChatBot
1. Dynamic question answering using VBA
2. Search sales details using Order ID
3. Automated data analysis inside Excel
4. Revenue and sales performance tracking
5. Interactive chatbot interface

## Technologies Used
1. Microsoft Excel
2. VBA (Visual Basic for Applications)
3. Excel Formulas
4. Data Analysis Techniques
5. Dashboard Creation
6. Automation

## ChatBot
<a href = "https://github.com/Prembsp28/DA-ChatBot-with-VBA-Project/blob/main/Screenshot%202026-05-16%20131156.png">ChatBot interface</a>

## VBA CODE 

Option Explicit

Sub AskChatBot()

    Dim wsData As Worksheet
    Dim wsBot As Worksheet

    Dim lastRow As Long
    Dim query As String
    Dim response As String

    Dim i As Long
    Dim totalValue As Double
    Dim foundKeyword As String

    Set wsData = ThisWorkbook.Sheets("Raw_Sales_Data")
    Set wsBot = ThisWorkbook.Sheets("ChatBot")

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    query = LCase(Trim(wsBot.Range("E8").Value))

    totalValue = 0
    response = ""
    foundKeyword = ""

    For i = 2 To lastRow

        Dim orderID As String
        Dim region As String
        Dim salesperson As String
        Dim category As String
        Dim product As String

        orderID = LCase(wsData.Cells(i, 1).Value)
        region = LCase(wsData.Cells(i, 3).Value)
        salesperson = LCase(wsData.Cells(i, 4).Value)
        category = LCase(wsData.Cells(i, 5).Value)
        product = LCase(wsData.Cells(i, 6).Value)

 
        If InStr(query, orderID) > 0 Then
            foundKeyword = wsData.Cells(i, 1).Value

        ElseIf InStr(query, region) > 0 Then
            foundKeyword = wsData.Cells(i, 3).Value

        ElseIf InStr(query, salesperson) > 0 Then
            foundKeyword = wsData.Cells(i, 4).Value

        ElseIf InStr(query, category) > 0 Then
            foundKeyword = wsData.Cells(i, 5).Value

        ElseIf InStr(query, product) > 0 Then
            foundKeyword = wsData.Cells(i, 6).Value

        End If

        If foundKeyword <> "" Then

            
            If InStr(query, "order date") > 0 Then

                response = "Order Date of " & foundKeyword & _
                " is : " & Format(wsData.Cells(i, 2).Value, "yyyy-mm-dd")

                Exit For

           
            ElseIf InStr(query, "quantity") > 0 Then

                totalValue = totalValue + wsData.Cells(i, 7).Value
                response = "Quantity of " & foundKeyword & _
                " is : " & totalValue

           
            ElseIf InStr(query, "unit price") > 0 Then

                totalValue = totalValue + wsData.Cells(i, 8).Value
                response = "Unit Price of " & foundKeyword & _
                " is : $ " & totalValue

           
            ElseIf InStr(query, "sales") > 0 Then

                totalValue = totalValue + wsData.Cells(i, 9).Value
                response = "Sales Amount of " & foundKeyword & _
                " is : $ " & totalValue

            
            ElseIf InStr(query, "target") > 0 Then

                totalValue = totalValue + wsData.Cells(i, 10).Value
                response = "Target Amount of " & foundKeyword & _
                " is : $ " & totalValue

            ElseIf InStr(query, "status") > 0 Then

                response = "Status of " & foundKeyword & _
                " is : " & wsData.Cells(i, 11).Value

                Exit For

            End If

        End If

    Next i

    If response = "" Then
        response = "No matching data found."
    End If

    wsBot.Range("E12").Value = response

End Sub


