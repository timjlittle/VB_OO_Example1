Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

'***************************************************
'* Class that holds a list of all the products
'* There is no way to create a new product in this example
'* This will need to be done via the spreadsheet
'* or as a task fo the student
'* A sensible addition would be a "Save" sub which 
'* either updated or created each product
'* 
'* Author Tim Little 2014
'*
'* This code is presented for example purposes only. 
'* Any re-use of this code is at the users discretion.
'* no liability is accepted
'***************************************************
Public Class ProductList
    Public Items As List(Of Product)

    '**********************************************************
    '* Constructor, reads all the products and adds them to a list
    Public Sub New()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim row As Integer
        Dim empty As Boolean = False
        Dim ProductId As Integer
        Dim Description As String
        Dim PriceExVat As Double
        Dim newProduct As Product

        On Error GoTo errHandler

        xlApp = CreateObject("Excel.Application")

        'NB ExcelPath defined in the Globals module
        xlWorkBook = xlApp.Workbooks.Open(ExcelPath)
        'Hard coding strings is poor practice
        xlWorkSheet = xlWorkBook.Worksheets("Products")

        '* This is poor programming because it uses hardcoded
        '* "magic" numbers for the columns
        '* And doesn't check they have suitable values
        '* A beter solution would have more checking, but this obscures the process
        row = 2
        Do
            If xlWorkSheet.Cells(row, 1).ToString <> "" Then
                ProductId = xlWorkSheet.Cells(row, 1).value
                Description = xlWorkSheet.Cells(row, 2).Value
                PriceExVat = xlWorkSheet.Cells(row, 3).Value

                newProduct = New Product(ProductId, Description, PriceExVat)
                Items.Add(newProduct)
            Else
                empty = True
            End If
            row = row + 1
        Loop Until empty


        'Close all the spreadsheet objects to release locks
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Exit Sub

        '* Error handler - make sure all excel objects are closed cleanly
errHandler:

        MsgBox(Err.Description)

        If Not IsNothing(xlWorkSheet) Then
            xlWorkBook.Close()
            releaseObject(xlWorkSheet)
        End If

        If Not IsNothing(xlWorkBook) Then
            releaseObject(xlWorkBook)
        End If

        If Not IsNothing(xlApp) Then
            xlApp.Quit()
            releaseObject(xlApp)
        End If

    End Sub

    '********************************************************************************
    '* Private helper functions
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
