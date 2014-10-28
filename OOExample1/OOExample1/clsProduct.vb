Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

'***************************************************
'* Class that holds the details of an individual product
'* There is no way to create a new product in this example
'* This will need to be done via the spreadsheet
'* or as a task fo the student
'* 
'* Author Tim Little 2014
'*
'* This code is presented for example purposes only. 
'* Any re-use of this code is at the users discretion.
'* no liability is accepted
'***************************************************

Public Class Product
    Private myDescription As String
    Private myPriceExVat As Double
    Private myProductId As Integer

    '* Public constructor
    '* This reads the details of the product
    Public Sub New(ByRef pProductId As Integer)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim row As Integer
        Dim empty As Boolean = False

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
                myProductId = xlWorkSheet.Cells(row, 1).value
                myDescription = xlWorkSheet.Cells(row, 2).Value
                myPriceExVat = xlWorkSheet.Cells(row, 3).Value
            Else
                empty = True
            End If
            row = row + 1
        Loop Until myProductId = pProductId Or empty

        'Ought to Throw an error if not found
        If empty Then
            Throw New ApplicationException("Invalid Product Id")
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

    Public Sub New(ByRef pProductId As Integer, ByRef pDesc As String, ByRef pPrice As Double)
        myProductId = pProductId
        myDescription = pDesc
        myPriceExVat = pPrice
    End Sub
    '******************************************************************************
    ' Public properties

    Public ReadOnly Property Description As String
        Get
            Description = myDescription
        End Get
    End Property

    Public ReadOnly Property ProductId As Integer
        Get
            ProductId = myProductId
        End Get
    End Property

    Public ReadOnly Property PriceExVat As Double
        Get
            PriceExVat = myPriceExVat
        End Get
    End Property
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
