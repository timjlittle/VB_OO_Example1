Imports Microsoft.Office.Interop

'***************************************************
'* Class that holds the details of an individual order item
'* 
'* Author Tim Little 2014
'*
'* This code is presented for example purposes only. 
'* Any re-use of this code is at the users discretion.
'* no liability is accepted
'***************************************************
Public Class OrderItem
    Private myProduct As Product
    Private myQuantity As Integer
    Private myOrderId As Integer
    Private myItemId As Integer
    Private loaded As Boolean = False

    Public Overrides Function ToString() As String
        Dim retStr As String = ""

        If loaded Then
            retStr = myQuantity & " x " & myProduct.Description
        End If

        Return retStr
    End Function

    Public ReadOnly Property CostExVat As Double
        Get
            If loaded Then
                CostExVat = myQuantity * myProduct.PriceExVat
            Else
                Throw New ApplicationException("uninitialised order item")
            End If
        End Get
    End Property

    Public ReadOnly Property ProductId As Integer
        Get
            If loaded Then
                Return myProduct.ProductId
            Else
                Throw New ApplicationException("uninitialised order item")
            End If
        End Get
    End Property

    Public ReadOnly Property Quantity As Integer
        Get
            If loaded Then
                Return myQuantity
            Else
                Throw New ApplicationException("uninitialised order item")
            End If
        End Get
    End Property

    Public ReadOnly Property ProductDescription As String
        Get
            If loaded Then
                Return myProduct.Description
            Else
                Throw New ApplicationException("uninitialised order item")
            End If
        End Get
    End Property


    Public Sub load(ByRef pItemId As Integer)
        On Error GoTo errHandler
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim row As Integer
        Dim empty As Boolean = False
        Dim ProductId As Integer

        xlApp = CreateObject("Excel.Application")

        'NB ExcelPath defined in the Globals module
        xlWorkBook = xlApp.Workbooks.Open(ExcelPath)
        'Hard coding strings is poor practice
        xlWorkSheet = xlWorkBook.Worksheets("OrderItems")

        '* This is poor programming because it uses hardcoded
        '* "magic" numbers for the columns
        '* And doesn't check they have suitable values
        '* A beter solution would have more checking, but this obscures the process
        row = 2
        Do
            If xlWorkSheet.Cells(row, 1).ToString <> "" Then
                myItemId = xlWorkSheet.Cells(row, 1).Value
                myOrderId = xlWorkSheet.Cells(row, 2).Value
                ProductId = xlWorkSheet.Cells(row, 3).Value
                myQuantity = xlWorkSheet.Cells(row, 4).Value
            Else
                empty = True
            End If
            row = row + 1

        Loop Until myItemId = pItemId Or empty

        'Ought to Throw an error if not found
        If empty Then
            Throw New ApplicationException("Invalid order item Id")
        Else
            myProduct = New Product(ProductId)
            loaded = True
        End If

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

    Public Sub Create(ByRef pOrderId As Integer, ByRef pProductId As Integer, ByRef pQuantity As Integer)
        On Error GoTo errHandler
        myProduct = New Product(pProductId)
        myQuantity = pQuantity
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim row As Integer
        Dim empty As Boolean = False
        Dim ProductId As Integer

        xlApp = CreateObject("Excel.Application")

        'NB ExcelPath defined in the Globals module
        xlWorkBook = xlApp.Workbooks.Open(ExcelPath)
        'Hard coding strings is poor practice
        xlWorkSheet = xlWorkBook.Worksheets("OrderItems")

        '* This is poor programming because it uses hardcoded
        '* "magic" numbers for the columns
        '* And doesn't check they have suitable values
        '* A beter solution would have more checking, but this obscures the process
        row = 2
        myItemId = 0
        Do
            If xlWorkSheet.Cells(row, 1).ToString <> "" Then
                If myItemId < xlWorkSheet.Cells(row, 1).Value Then
                    myItemId = xlWorkSheet.Cells(row, 1).Value
                End If

            Else
                empty = True
            End If
            row = row + 1
        Loop Until empty

        myItemId = myItemId + 1
        xlWorkSheet.Cells(row, 1).value = myItemId
        xlWorkSheet.Cells(row, 2).value = pOrderId
        xlWorkSheet.Cells(row, 3).value = pProductId
        xlWorkSheet.Cells(row, 3).value = pQuantity


        loaded = True

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

    Public Sub Remove()
        On Error GoTo errHandler
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim row As Integer
        Dim found As Boolean = False
        Dim empty As Boolean = False

        xlApp = CreateObject("Excel.Application")

        'NB ExcelPath defined in the Globals module
        xlWorkBook = xlApp.Workbooks.Open(ExcelPath)
        'Hard coding strings is poor practice
        xlWorkSheet = xlWorkBook.Worksheets("OrderItems")

        '* This is poor programming because it uses hardcoded
        '* "magic" numbers for the columns
        '* And doesn't check they have suitable values
        '* A beter solution would have more checking, but this obscures the process
        row = 2

        Do
            If xlWorkSheet.Cells(row, 1).value = myItemId Then
                found = True
            ElseIf xlWorkSheet.Cells(row, 1).ToString = "" Then
                empty = True
            Else
                row = row + 1
            End If
        Loop Until found Or empty

        If found Then
            xlWorkSheet.Rows(row).delete()
        End If

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
