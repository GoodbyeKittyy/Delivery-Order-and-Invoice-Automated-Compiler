Attribute VB_Name = "ExtractData"
Option Explicit

Public Sub DataExtraction()

  Dim DestinationTable As Excel.ListObject
  Dim Files As VBA.Collection

  Dim Folder As String

  Set DestinationTable = ActiveWorkbook.Sheets(1).Range("Table1").ListObject
  TableClearTable DestinationTable

  If RetrieveFiles(Files) Then
    If ProcessFiles(DestinationTable, Files) Then
      MsgBox "Task completed (" & Files.Count & " files processed).", vbInformation + vbOKOnly
    Else
      MsgBox "Error while processing files.", vbCritical + vbOKOnly
    End If
  Else
    MsgBox "Error while retrieving files.", vbCritical + vbOKOnly
  End If

  Set DestinationTable = Nothing
  Set Files = Nothing

End Sub

Private Function ExtractData(ByVal CSourceSheet As Excel.Worksheet, ByVal CDestinationTable As Excel.ListObject, ByRef CParsingPoints() As Variant) As Boolean

  Const HEADER_DOCUMENT_NUMBER As Long = 0
  Const HEADER_DOCUMENT_DATE As Long = 1
  Const HEADER_SERIAL_ORDER_NUMBER As Long = 2
  Const COLUMN_PART_NUMBER As Long = 3
  Const COLUMN_CUSTOMER_PURCHASE_ORDER As Long = 4
  Const COLUMN_QUANTITY As Long = 5

  On Local Error GoTo LocalError

  Dim CellCustomerPurchaseOrder As Excel.Range
  Dim CellPartNumber As Excel.Range
  Dim CellQuantity As Excel.Range

  Dim CurrentRowOffset As Long

  Dim HeaderDocumentNumber As String
  Dim HeaderDocumentDate As Date
  Dim HeaderSerialOrderNumber As String
  Dim HeaderSupplierName As String
  Dim ColumnCustomerPurchaseOrder As String
  Dim ColumnPartNumber As String
  Dim ColumnQuantity As Double

  HeaderDocumentNumber = CSourceSheet.Cells.Find(CParsingPoints(HEADER_DOCUMENT_NUMBER), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(, 1).Value
  HeaderDocumentDate = CDate(CSourceSheet.Cells.Find(CParsingPoints(HEADER_DOCUMENT_DATE), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(, 1).Value)
  HeaderSerialOrderNumber = CSourceSheet.Cells.Find(CParsingPoints(HEADER_SERIAL_ORDER_NUMBER), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(, 1).Value
  HeaderSupplierName = ExtractFirstRowText(CSourceSheet)
  Set CellPartNumber = CSourceSheet.Cells.Find(CParsingPoints(COLUMN_PART_NUMBER), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(2)
  Set CellCustomerPurchaseOrder = CSourceSheet.Cells.Find(CParsingPoints(COLUMN_CUSTOMER_PURCHASE_ORDER), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(2)
  Set CellQuantity = CSourceSheet.Cells.Find(CParsingPoints(COLUMN_QUANTITY), , xlValues, xlWhole, xlByRows, xlNext, True, False).Offset(2)
  CurrentRowOffset = 0
  Do
    If Len(Trim(CellPartNumber.Offset(CurrentRowOffset).Value & "")) > 0 Then
      ColumnPartNumber = CellPartNumber.Offset(CurrentRowOffset).Value
      ColumnCustomerPurchaseOrder = CellCustomerPurchaseOrder.Offset(CurrentRowOffset).Value
      ColumnQuantity = CDbl(CellQuantity.Offset(CurrentRowOffset).Value)
      TableAddRow CDestinationTable, Array(HeaderSupplierName, HeaderDocumentNumber, HeaderDocumentDate, HeaderSerialOrderNumber, ColumnCustomerPurchaseOrder, ColumnPartNumber, ColumnQuantity)
      CurrentRowOffset = CurrentRowOffset + 1
    Else
      Exit Do
    End If
  Loop

  Set CellCustomerPurchaseOrder = Nothing
  Set CellPartNumber = Nothing
  Set CellQuantity = Nothing
  Exit Function

LocalError:
  Debug.Print "ExtractData(): Error " & Err.Number & " while extracting data." & vbCrLf & vbTab & Err.Description

End Function
Private Function ExtractFirstRowText(ByVal CSheet As Excel.Worksheet) As String

  Dim Count As Long
  Dim Result As String

  Count = 0
  Do While Count < 255
    Count = Count + 1
    If Len(Trim(CSheet.Cells(1, Count).Value & "")) > 0 Then
      Result = Result & CSheet.Cells(1, Count).Value & ", "
    End If
  Loop

  If Len(Result) > 0 Then
    Result = Left(Result, Len(Result) - 2)
  End If

  ExtractFirstRowText = Result

End Function

Private Function ExtractParsingPoints(ByVal CSourceSheet As Excel.Worksheet, ByRef OParsingPoints() As Variant) As Boolean

  ' Document types.
  ' Why on earth do these cells contain extra spacings??

  Const DOCUMENT_TYPE_DELIVERY_ORDER As String = "  DELIVERY ORDER"
  Const DOCUMENT_TYPE_TAX_INVOICE As String = "TAX   INVOICE"
  ' Type depended.
  Const HEADER_DELIVERY_ORDER_NUMBER As String = "DO No:"
  Const HEADER_DELIVERY_ORDER_DATE As String = "DO Date:"
  Const COLUMN_DELIVERY_ORDER_PART_NUMBER As String = "PART NO / DESCRIPTION"
  Const HEADER_TAX_INVOICE_NUMBER As String = "Invoice No:"
  Const HEADER_TAX_INVOICE_DATE As String = "Invoice Date:"
  Const COLUMN_TAX_INVOICE_PART_NUMBER As String = "PART-NO / DESCRIPTION"
  ' Common.
  Const HEADER_SERIAL_ORDER_NUMBER As String = "S/O No:"
  Const COLUMN_CUSTOMER_PURCHASE_ORDER As String = "CUST-PO"
  Const COLUMN_QUANTITY As String = "QTY"

  On Local Error GoTo LocalError

  Dim CellDocumentType As Excel.Range

  Set CellDocumentType = CSourceSheet.Cells.Find(DOCUMENT_TYPE_DELIVERY_ORDER, , xlValues, xlWhole, xlByRows, xlNext, True, False)
  If Not CellDocumentType Is Nothing Then
    OParsingPoints = Array(HEADER_DELIVERY_ORDER_NUMBER, HEADER_DELIVERY_ORDER_DATE, HEADER_SERIAL_ORDER_NUMBER, COLUMN_DELIVERY_ORDER_PART_NUMBER, COLUMN_CUSTOMER_PURCHASE_ORDER, COLUMN_QUANTITY)
    ExtractParsingPoints = True
  Else
    Set CellDocumentType = CSourceSheet.Cells.Find(DOCUMENT_TYPE_TAX_INVOICE, , xlValues, xlWhole, xlByRows, xlNext, True, False)
    If Not CellDocumentType Is Nothing Then
      OParsingPoints = Array(HEADER_TAX_INVOICE_NUMBER, HEADER_TAX_INVOICE_DATE, HEADER_SERIAL_ORDER_NUMBER, COLUMN_TAX_INVOICE_PART_NUMBER, COLUMN_CUSTOMER_PURCHASE_ORDER, COLUMN_QUANTITY)
      ExtractParsingPoints = True
    Else
      Debug.Print "ExtractParsingPoints(): Unknown document type."
      ExtractParsingPoints = False
    End If
  End If

  Set CellDocumentType = Nothing
  Exit Function

LocalError:
  Debug.Print "ExtractParsingPoints(): Error " & Err.Number & " while extracting parsing points." & vbCrLf & vbTab & Err.Description

End Function

Private Function ProcessFile(ByVal CDestinationTable As Excel.ListObject, ByVal CFile As String) As Boolean

  On Local Error GoTo LocalError

  Dim ExcelApplication As Excel.Application
  Dim FileWorkbook As Excel.Workbook
  Dim SourceSheet As Excel.Worksheet

  Dim File As String
  Dim ParsingPoints() As Variant

  ProcessFile = False
  Debug.Print "Processing file '" & CFile & "'.."
  Set ExcelApplication = ExcelInstance(False)
  Set FileWorkbook = ExcelApplication.Workbooks.Open(CFile, False, True)
  Set SourceSheet = FileWorkbook.Sheets(1)
  ProcessFile = ExtractParsingPoints(SourceSheet, ParsingPoints)
  If ProcessFile Then
    ProcessFile = ExtractData(SourceSheet, CDestinationTable, ParsingPoints)
  End If

  Set SourceSheet = Nothing
  FileWorkbook.Close False
  Set FileWorkbook = Nothing
  Set ExcelApplication = Nothing
  Exit Function

LocalError:
  Debug.Print "ProcessFile(): Error " & Err.Number & " while processing file '" & CFile & "'." & vbCrLf & vbTab & Err.Description
  ExcelInstance True
  Set ExcelApplication = Nothing

End Function

Private Function ProcessFiles(ByVal CDestinationTable As Excel.ListObject, ByVal CFiles As VBA.Collection) As Boolean

  On Local Error GoTo LocalError

  Dim File As Variant

  ProcessFiles = False
  If CFiles.Count = 0 Then
    Debug.Print "Processing " & CFiles.Count & " files."
  Else
    Debug.Print "Processing " & CFiles.Count & " files:"
    For Each File In CFiles
      ProcessFile CDestinationTable, CStr(File)
    Next File

    ExcelInstance True
    Debug.Print "All files processed."
  End If

  ProcessFiles = True
  Exit Function

LocalError:
  Debug.Print "ProcessFiles(): Error " & Err.Number & " while processing files." & vbCrLf & vbTab & Err.Description

End Function

Private Function RetrieveFiles(ByRef OFiles As VBA.Collection) As Boolean

  On Local Error GoTo LocalError

  Dim FileDialog As Object

  Dim File As Variant

  Set OFiles = New VBA.Collection

  RetrieveFiles = False
  Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
  FileDialog.Title = "Select files to extract data"
  FileDialog.ButtonName = "Confirm"
  If (FileDialog.Show = -1) Then
    For Each File In FileDialog.SelectedItems
      OFiles.Add File
    Next File
  End If

  RetrieveFiles = True
  Set FileDialog = Nothing
  Exit Function

LocalError:
  Debug.Print "RetrieveFiles(): Error " & Err.Number & " while retrieving files." & vbCrLf & vbTab & Err.Description

End Function

Private Function ExcelInstance(ByVal CCleanUp As Boolean) As Excel.Application

  Static ExcelApplication As Excel.Application

  If CCleanUp Then
    Set ExcelApplication = Nothing
    Set ExcelInstance = Nothing
  Else
    If ExcelApplication Is Nothing Then
      Set ExcelApplication = New Excel.Application
    End If

    Set ExcelInstance = New Excel.Application
  End If

End Function

Private Sub TableAddRow(ByVal CTable As Excel.ListObject, ByVal CData As Variant)

  Dim NewRow As Excel.ListRow

  Set NewRow = CTable.ListRows.Add(, True)
  If TypeName(CData) = "Range" Then
    NewRow.Range = CData.Value
  Else
    NewRow.Range = CData
  End If

  Set NewRow = Nothing

End Sub

Private Sub TableClearTable(ByVal CTable As Excel.ListObject)

  On Local Error Resume Next

  CTable.DataBodyRange.Rows.Delete

End Sub
