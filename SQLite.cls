VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'_________________________________________________
'John Cagle - SQLite Object
'libraries needed:
'   Microsoft ActiveX Data Objects Library 6.1
'   Microsoft Scripting Runtime
'_________________________________________________

Private ADODBRecordSet As ADODB.Recordset
Private boolIsResultEmpty As Boolean
Private boolIsResultTooLarge As Boolean
Private boolHeaders As Boolean
Private boolIsExistingFile As Boolean
Private boolKeepPPQuery As Boolean
Private strResultPath As String
Private strErrMsg As String
Private strNameArg As String
Private strDatabasePublic As String
Private strTimeStamp As String
Private strFriendlyName As String
Private arrSQLiteExtensions(2) As String
Private arrTableInfo() As String
Private arrGlbHeaders() As String
Private rngListObjectDestination As Range
Private intReturnTypeLimitWhen4GB As Long
Private intSatan As Long
Private dteTimerStart As Date
Public Enum enumReturnType
    [Path_Type] = 1
    [File_Type] = 2
    [Array_Type] = 3
    [Recordset_Type] = 4
    [ListObject_Type] = 5
    [Text_Type] = 6
    [Range_Type] = 7
    [PowerPivotListObject_Type] = 8
End Enum
Private Enum ClassError
    [NoErrors] = 1
    [UnknownError] = 2
    [FileSizeTooLarge] = 3
    [ResultFileDoesNotExsist] = 4
    [ResultFileStringNotSet] = 5
    [RecordSetFieldTooSmall] = 6
    [DBDoesNotExsist] = 7
    [SQLiteNotInstalled] = 8
    [RecordSetNoHeaderError] = 9
    [DBNotFoundinXLStart] = 10
    [DBUnexpectedExtension] = 11
    [FilePathsTooLong] = 12
    [InvalidDrive] = 13
    [DestinationRangeInATable] = 14
    [ExistingFileExtensionNotRecognized] = 15
    [UnidentifiedCustomFormattingCharacter] = 16
End Enum
Private enumClassError As ClassError
Private Sub Class_Initialize()
    Dim fso As New Scripting.FileSystemObject
    boolIsResultEmpty = True
    boolIsResultTooLarge = False
    boolIsExistingFile = False
    Set ADODBRecordSet = New ADODB.Recordset
    enumClassError = NoErrors
    intReturnTypeLimitWhen4GB = 2
    'custom error code (has to be bigger than 513)
    intSatan = 666
    'list expected SQLite Extensions
    arrSQLiteExtensions(0) = ".db3"
    arrSQLiteExtensions(1) = ".db"
    arrSQLiteExtensions(2) = ".sqlite"
    'verify sqlite installation
    If Not fso.FileExists("C:\cygwin\bin\sqlite3.exe") Then
        enumClassError = SQLiteNotInstalled
        ReportError
    End If
    Set fso = Nothing
End Sub
Private Sub Class_Terminate()
    Dim fso As New Scripting.FileSystemObject
    On Error Resume Next
    ADODBRecordSet.Close
    On Error GoTo 0
    Close #1
    If fso.FileExists(strResultPath) And boolIsExistingFile = False Then Kill strResultPath
    Set fso = Nothing
    Set ADODBRecordSet = Nothing
End Sub
Private Property Get IsResultEmpty() As Boolean
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    If strResultPath <> "" Then
        If Not fso.FileExists(strResultPath) Then
            boolIsResultEmpty = True
            enumClassError = ResultFileDoesNotExsist
            ReportError
        End If
        Set fsoFile = fso.GetFile(strResultPath)
        If fsoFile.Size = 0 Then
            boolIsResultEmpty = True
        Else
            boolIsResultEmpty = False
        End If
    Else
        boolIsResultEmpty = True
        enumClassError = ResultFileStringNotSet
        ReportError
    End If
    IsResultEmpty = boolIsResultEmpty
End Property
Private Property Get ReturnPath() As String
    ReturnPath = strResultPath
End Property
Private Property Get ReturnText() As String
    Dim fso As New Scripting.FileSystemObject
    Dim fsoStream As Scripting.TextStream
    Set fsoStream = fso.OpenTextFile(strResultPath, ForReading, False)
    'ReturnText = fsoStream.ReadAll
    ReturnText = fsoStream.ReadLine
    fsoStream.Close
    Set fsoStream = Nothing
    Set fso = Nothing
End Property
Private Property Get ReturnFile() As Scripting.File
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(strResultPath) Then
        Set ReturnFile = fso.GetFile(strResultPath)
    Else
        enumClassError = ResultFileDoesNotExsist
        ReportError
    End If
    Set fso = Nothing
End Property
Private Property Get ReturnArray() As String()
    Dim str As String
    Dim fso As New Scripting.FileSystemObject
    Dim fsoStream As Scripting.TextStream
    Dim arr() As String
    Dim i As Long
    If fso.FileExists(strResultPath) Then
        Set fsoStream = fso.OpenTextFile(strResultPath)
        Do While Not fsoStream.AtEndOfStream
            ReDim Preserve arr(i)
            arr(i) = fsoStream.ReadLine
            i = i + 1
        Loop
        fsoStream.Close
        Set fsoStream = Nothing
        Set fso = Nothing
        ReturnArray = arr()
        'Kill strResultPath
    Else
        enumClassError = ResultFileDoesNotExsist
        ReportError
    End If
End Property
Private Property Get ReturnListObject() As ListObject
    Dim adodbRS As ADODB.Recordset
    Dim adodbField As Object
    Dim qtTempQueryTable As Excel.QueryTable
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim wsTest As Worksheet
    Dim wb As Workbook
    Dim rng As Range
    Dim rngTable As Range
    Dim lngDecimals As Long
    Dim lngCheckPlace As Long
    Dim lngHeader As Long: lngHeader = 0
    Dim strFormatChrs As String
    Dim strFormatChr1 As String
    Dim strFormatChr2 As String
    Dim strFormatChr3 As String
    Dim strDecimal As String
    Dim strChrTest As String
    Dim strName As String
    Dim strCellFormat As String
    Dim strDigits As String
    Dim strTrailingFormatChrs As String
    Dim loCheck As ListObject
    Dim lngHour As Long, lngMinute As Long, lngSecond As Long
    Dim waitTime As Date
        
    'handle omitted destination
    If rngListObjectDestination Is Nothing Then
        Set wb = ThisWorkbook
        Set ws = wb.Sheets.Add(, wb.Sheets(wb.Sheets.Count), 1, xlWorksheet)
        If Len(strFriendlyName) > 25 Then strFriendlyName = Right(strFriendlyName, 25)
        On Error Resume Next
        Set wsTest = wb.Sheets(strFriendlyName)
        On Error GoTo 0
        If wsTest Is Nothing Then
            ws.Name = strFriendlyName
        Else
            err.Raise 666, , "Sheet already exists: " & strFriendlyName
        End If
        Set rngListObjectDestination = ws.Range("A1")
    Else
        Set ws = rngListObjectDestination.Parent
        'check if destination is in a table
        For Each lo In ws.ListObjects
            If Not Intersect(rngListObjectDestination, lo.Range) Is Nothing Then
                enumClassError = DestinationRangeInATable
                Set adodbRS = Nothing
                Set qtTempQueryTable = Nothing
                Set rng = Nothing
                ReportError
            End If
        Next lo
    End If
    
    'get range for table, set a recordset so i can close it and delete the Connection
    Set adodbRS = ReturnRecordset
    If adodbRS.RecordCount > 1000000 Then
        enumClassError = FileSizeTooLarge
        ReportError
    End If
    Set qtTempQueryTable = ws.QueryTables.Add( _
            Connection:=adodbRS, _
            Destination:=rngListObjectDestination)
    With qtTempQueryTable
        .Name = "rng_" & strFriendlyName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    Set rngTable = qtTempQueryTable.ResultRange
    qtTempQueryTable.WorkbookConnection.Delete
    qtTempQueryTable.Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, rngTable, , xlYes, rngListObjectDestination)
    'used to see if selected column titles have custom SQLite formatting characters
    strTrailingFormatChrs = "$%#*"
    For Each adodbField In adodbRS.Fields
        Select Case adodbField.Type
            Case adCurrency
                lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = "$#,##0.00"
            Case adDouble
                'nothing, general number format recognizes numbers
                lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = "General"
                'format number values other than Currency in SQL Query - SQLite Command example: printf(""%.2f"", columnName)
            Case adDBDate
                lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = "yyyymmdd"
            Case adVarChar
                lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = ""
            Case adLongVarChar
                lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = ""
        End Select
        
        
        'check field name for custom SQLite formatting characters
         'check 3 and 2 characters from the end of the string
         'data type character ($ for currency, % for percents, # for numbers)
        If (Len(arrGlbHeaders(lngHeader)) - InStrRev(arrGlbHeaders(lngHeader), "_", , vbTextCompare)) <= 3 And InStr(1, arrGlbHeaders(lngHeader), "_", vbTextCompare) > 0 And InStr(1, strTrailingFormatChrs, Left(Right(arrGlbHeaders(lngHeader), Len(arrGlbHeaders(lngHeader)) - InStrRev(arrGlbHeaders(lngHeader), "_", , vbTextCompare)), 1)) > 0 Then
            strFormatChrs = Right(arrGlbHeaders(lngHeader), Len(arrGlbHeaders(lngHeader)) - InStrRev(arrGlbHeaders(lngHeader), "_", , vbTextCompare))
            strFormatChr1 = Left(strFormatChrs, 1)
            Select Case Len(strFormatChrs)
                Case 1
                    strFormatChr2 = ""
                    strFormatChr3 = ""
                Case 2
                    strFormatChr2 = Right(strFormatChrs, 1)
                    strFormatChr3 = "0"
                Case 3
                    strFormatChr2 = Mid(strFormatChrs, 2, 1)
                    strFormatChr3 = Right(strFormatChrs, 1)
                Case Else
                    enumClassError = UnidentifiedCustomFormattingCharacter
                    ReportError
            End Select
            If (Not IsNumeric(strFormatChr2)) And strFormatChr2 <> "" Then
                enumClassError = UnidentifiedCustomFormattingCharacter
                ReportError
            End If
            If strFormatChr2 <> "" Then lngDecimals = CLng(strFormatChr2)
        Else
            strFormatChrs = ""
        End If

        If strFormatChrs <> "" And InStr(1, strTrailingFormatChrs, strFormatChr1, vbTextCompare) > 0 Then
            'number of decimal places for custom trailing SQLite formatting character
             'can also be a character (ex. "-")
            'get decimal charachter
            If lngDecimals > 0 Then
                strDecimal = "."
            Else
                strDecimal = ""
            End If
            'get decimal zero (or empty string, if there are no decimal places)
            If strFormatChr3 = "0" Then
                strDigits = strDecimal & String(lngDecimals, "0")
            Else
                strDigits = ""
            End If
            'set each column's format and number of decimal places
            Select Case strFormatChr1
                Case "$"
                    'default dollars to 2 decimals
                    If strFormatChr2 = "" Then
                        lngDecimals = 2
                        strDecimal = "."
                    End If
                    strCellFormat = "$* #,##0" & strDecimal & String(lngDecimals, "0") & ";$* \-#,##0" & strDecimal & String(lngDecimals, "0") & ";$* """ & strFormatChr3 & strDigits & """;_(@_)"
                Case "%"
                    strCellFormat = "#,##0" & strDecimal & String(lngDecimals, "0") & "%;\-#,##0" & strDecimal & String(lngDecimals, "0") & "%;""" & strFormatChr3 & strDigits & "%"""
                Case "#"
                    strCellFormat = "#,##0" & strDecimal & String(lngDecimals, "0") & ";\-#,##0" & strDecimal & String(lngDecimals, "0") & ";" & strFormatChr3 & strDigits
                Case "*"
                    strCellFormat = "@"
            End Select
            lo.ListColumns(adodbField.Name).DataBodyRange.NumberFormat = strCellFormat
            lo.ListColumns(adodbField.Name).Name = adodbField.Name
        End If
        lngHeader = lngHeader + 1
    Next adodbField
    adodbRS.Close
    
    On Error Resume Next
    Set loCheck = rngListObjectDestination.Parent.ListObjects("tbl_" & Replace(strFriendlyName, " ", ""))
    If Not loCheck Is Nothing Then
        lngHour = Hour(Now())
        lngMinute = Minute(Now())
        lngSecond = Second(Now()) + 5
        waitTime = TimeSerial(lngHour, lngMinute, lngSecond)
        Application.Wait waitTime
        strTimeStamp = Format(Now(), "yyyymmddhhnnss")
        strFriendlyName = strTimeStamp
    End If
    On Error GoTo 0
    If Len(strFriendlyName) > 20 Then strFriendlyName = Right(strFriendlyName, 20)
    With lo
        '.TableStyle = ""
        .Name = "tbl_" & Replace(strFriendlyName, " ", "")
        .Range.Columns.AutoFit
    End With
    Call ListObjectBorders(lo.Range)
    Set rngListObjectDestination = Nothing
    strNameArg = ""
    Set ReturnListObject = lo
End Property
Private Property Get ReturnRecordset() As ADODB.Recordset
    Dim intFile As Long, i As Long, intRow As Long: intRow = 1
    Dim fso As New Scripting.FileSystemObject
    Dim fsoStream As Scripting.TextStream
    Dim fsoFile As Scripting.File
    Dim arr() As String
    Dim arrHeaders() As String
    Dim strRow As String
    Dim strDelimiter As String
    Dim strExtension As String
    Dim strHeader As String
    Dim d As Date
    Dim s As String
    Dim strFormatChars As String
    Dim strFixedCol As String
    Dim lngAdDataType As Long
    Dim lngAdDataTypeLen As Long
    Dim strTrailingFormatChrs As String: strTrailingFormatChrs = "$%#*"
    
    If enumClassError = FileSizeTooLarge Then
        boolIsResultTooLarge = True
        ReportError
        Exit Property
    End If
    Set fsoFile = ReturnFile
    With ADODBRecordSet
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
    End With
    intFile = 1
    'check file extension
    strExtension = Right(fsoFile.Name, Len(fsoFile.Name) - InStrRev(fsoFile.Name, ".") + 1)
    Select Case strExtension
        Case ".pdl"
            strDelimiter = "|"
        Case ".csv"
            strDelimiter = ","
        Case Else
            enumClassError = ExistingFileExtensionNotRecognized
            ReportError
    End Select
    Set fsoStream = fso.OpenTextFile(fsoFile.Path)
    Do While Not fsoStream.AtEndOfStream
        ReDim Preserve arr(intRow - 1)
        strRow = fsoStream.ReadLine
        arr = Split(strRow, strDelimiter)
        If intRow = 1 Then
            arrHeaders = arr
            arrGlbHeaders = arr
            For i = LBound(arr) To UBound(arr)
                'get custom format chars from column to decide data type
                 'remove customs chars from col name
                  'check if arr contains formatting chars
                If (Len(arr(i)) - InStrRev(arr(i), "_", , vbTextCompare)) <= 3 And InStr(1, arr(i), "_", vbTextCompare) > 0 And InStr(1, strTrailingFormatChrs, Left(Right(arr(i), Len(arr(i)) - InStrRev(arr(i), "_", , vbTextCompare)), 1)) > 0 Then
                    strFormatChars = Right(arr(i), Len(arr(i)) - InStrRev(arr(i), "_", , vbTextCompare))
                    strFixedCol = Left(arr(i), InStrRev(arr(i), "_", , vbTextCompare) - 1)
                Else
                    strFormatChars = ""
                    strFixedCol = arr(i)
                End If
                arrHeaders(i) = strFixedCol
                'choose data type based on custom data type char (at end of column)
                lngAdDataTypeLen = 255
                Select Case Left(strFormatChars, 1)
                    Case "#"
                        lngAdDataType = adDouble
                    Case "$"
                        lngAdDataType = adCurrency
                    Case "@"
                        lngAdDataType = adDBDate
                    Case "%"
                        lngAdDataType = adDouble
                    Case Else
                        lngAdDataType = adVarChar
                        lngAdDataTypeLen = 32000
                End Select
                ADODBRecordSet.Fields.Append strFixedCol, lngAdDataType, lngAdDataTypeLen, adFldMayBeNull
            Next i
            ADODBRecordSet.Open
        Else
            ADODBRecordSet.AddNew
            For i = LBound(arr) To UBound(arr)
                strHeader = arrHeaders(i)
                If (Len(arr(i)) > 255 And ADODBRecordSet.Fields(strHeader).Type <> adVarChar) Or Len(arr(i)) > 32000 Then
                    enumClassError = RecordSetFieldTooSmall
                    ReportError
                End If
                
                'cast arr(i) as correct type when providing values
                Select Case ADODBRecordSet.Fields(strHeader).Type
                    Case adDouble
                        '#
                        If arr(i) = "" Or UCase(arr(i)) = "NULL" Then
                            ADODBRecordSet.Fields(strHeader).Value = Null
                        Else
                            ADODBRecordSet.Fields(strHeader).Value = CDbl(arr(i))
                        End If
                    Case adCurrency
                        '$
                        If arr(i) = "" Or UCase(arr(i)) = "NULL" Then
                            ADODBRecordSet.Fields(strHeader).Value = Null
                        Else
                            ADODBRecordSet.Fields(strHeader).Value = CCur(arr(i))
                        End If
                    Case adDBDate
                        '@
                        If InStr(1, arr(i), "/", vbTextCompare) = 0 Then
                            If IsDate(Mid(arr(i), 5, 2) & "/" & Right(arr(i), 2) & "/" & Left(arr(i), 4)) Then
                                d = Mid(arr(i), 5, 2) & "/" & Right(arr(i), 2) & "/" & Left(arr(i), 4)
                            Else
                                d = CDate(0)
                            End If
                        Else
                            d = CDate(arr(i))
                        End If
                        If arr(i) = "" Or UCase(arr(i)) = "NULL" Then
                            ADODBRecordSet.Fields(strHeader).Value = Null
                        Else
                            ADODBRecordSet.Fields(strHeader).Value = d
                        End If
                    Case adVarChar
                        '_ nothing
                        ADODBRecordSet.Fields(strHeader).Value = arr(i)
                    Case adLongVarChar
                        ADODBRecordSet.Fields(strHeader).Value = arr(i)
                    Case Else
                        err.Raise 666, , "Unknown ADO data type"
                End Select
            Next i
            ADODBRecordSet.Update
        End If
        intRow = intRow + 1
    Loop
    fsoStream.Close
    
    'allow other properties like xlYes/xlNo/xlGuess in ReturnListObject, maybe tablestyle, headers, adFldMayBeNull
    'convert to binding
    'keep classes updated by making it an add-on
    '______________________________________________________________________________________________________
    'allow choice of field data type - adVarChar, adBoolean, adCurrency, adDBDate, adDBTimeStamp, adInteger,
        'this info could be obtained from sqlite schema
        
        'calling returnarray for pragma within recordset will fuck with the timestamp (runsqlite is run again)
            'make optional args for returnarray, if set, use provided table name strresultpath
            'make a global array to hold pragma table_info(TestView); results for every runsqlite that has a table arg
    '______________________________________________________________________________________________________

    Set ReturnRecordset = ADODBRecordSet.Clone
    ADODBRecordSet.Close
End Property
Private Property Get ReturnRange() As Range
    Dim arr() As String
    Dim arr2D() As Variant
    Dim rng As Range
    Dim l As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngRowCount As Long, lngColCount As Long
    Dim boolScreenUpdate As Boolean, boolEnableEvents  As Boolean, varCaclulation As Variant
    
    Dim fso As New Scripting.FileSystemObject
    Dim fsoStream As Scripting.TextStream
    
    With Excel.Application
        boolScreenUpdate = .ScreenUpdating
        boolEnableEvents = .EnableEvents
        varCaclulation = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    arr = ReturnArray
    ReDim arr2D(UBound(arr))
    For l = LBound(arr) To UBound(arr)
        arr2D(l) = Split(arr(l), "|", , vbTextCompare)
    Next l
    lngRowCount = UBound(arr2D): lngColCount = UBound(arr2D(0))
    Set rng = rngListObjectDestination.Resize(lngRowCount + 1, lngColCount + 1)
    For lngRow = 0 To lngRowCount
        For lngCol = 0 To lngColCount
            rng.Cells(lngRow + 1, lngCol + 1).Value = arr2D(lngRow)(lngCol)
        Next lngCol
    Next lngRow
    With Excel.Application
        .ScreenUpdating = boolScreenUpdate
        .EnableEvents = boolEnableEvents
        .Calculation = varCaclulation
    End With
    Set ReturnRange = rng
End Property
Private Property Get ReturnPowerPivotListObject() As ListObject
    Dim lo As ListObject
    Dim q As WorkbookQuery
    Dim strNameGeneral As String
    Dim strQueryPrefix As String: strQueryPrefix = "qry_"
    Dim strTblPrefix As String: strTblPrefix = "tbl_"
    Dim strPPFileName As String
    
    'power pivot check
    If rngListObjectDestination Is Nothing Or strFriendlyName = "" Then err.Raise intSatan, , "no table name or destination provided for power pivot listobject"
    strPPFileName = ReturnFile
    'get table and query names
    strNameGeneral = Replace(strFriendlyName, "-", "_")
    'If Len(strNameGeneral) > 30 Then strNameGeneral = Left(strNameGeneral, 20)
    If Left(strNameGeneral, 4) = strQueryPrefix Or Left(strNameGeneral, 4) = strTblPrefix Then
        strNameGeneral = Right(strNameGeneral, Len(strNameGeneral) - Len(strQueryPrefix))
    End If
    'overwrite previous data
    On Error Resume Next: rngListObjectDestination.Parent.ListObjects(strTblPrefix & strNameGeneral).Delete: rngListObjectDestination.Parent.Parent.Queries(strQueryPrefix & strNameGeneral).Delete: On Error GoTo 0
    'make query
    Set q = rngListObjectDestination.Parent.Parent.Queries.Add(strQueryPrefix & strNameGeneral, "let Source = Csv.Document(File.Contents(""" & strPPFileName & """),null,""|""), " & _
        "#""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers""")
    'make table
    Set lo = rngListObjectDestination.Parent.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=" & rngListObjectDestination.Parent.Parent.FullName & ";Location=" & strQueryPrefix & strNameGeneral & ";Extended Properties=""""", Destination:=rngListObjectDestination)
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & strQueryPrefix & strNameGeneral & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = Replace(strTblPrefix & strNameGeneral, "-", "_")
        .Refresh BackgroundQuery:=False
    End With
    If boolKeepPPQuery = False Then
        q.Delete
        lo.QueryTable.Delete
    End If
    Set ReturnPowerPivotListObject = lo
End Property


Public Function UseExistingFile(strPath As String, lngReturnType As enumReturnType, boolHeadersArg As Boolean, Optional strTable As String, Optional rngListObject As Range, Optional boolKeepPPQueryArg As Boolean) As Variant
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFileResult As Scripting.File
    Dim arrEmpty() As String
    Dim lo As ListObject
    
    'start timer
    dteTimerStart = Now()
    boolIsExistingFile = True
    strResultPath = strPath
    If Not fso.FileExists(strResultPath) Then strResultPath = ThisWorkbook.Path & "\" & strResultPath
    If Not fso.FileExists(strResultPath) Then err.Raise 666, , "no file..."
    Set fsoFileResult = fso.GetFile(strResultPath)
    strTimeStamp = Format(Now(), "yyyymmddhhnnss")
    strFriendlyName = Left(fsoFileResult.Name, InStrRev(fsoFileResult.Name, ".") - 1) & strTimeStamp
    If strTable <> "" Then strFriendlyName = strTable
    If Not rngListObject Is Nothing Then Set rngListObjectDestination = rngListObject
    boolKeepPPQuery = boolKeepPPQueryArg
    'set header choice
    If boolHeadersArg = False Then
        boolHeaders = False
    Else
        boolHeaders = True
    End If
    If boolHeadersArg = False And (lngReturnType = enumReturnType.Recordset_Type Or lngReturnType = enumReturnType.ListObject_Type Or lngReturnType = enumReturnType.PowerPivotListObject_Type) Then
        enumClassError = RecordSetNoHeaderError
        ReportError
        Set UseExistingFile = Nothing
        Set fso = Nothing
        Set fsoFileResult = Nothing
        Exit Function
    End If
    'check if file is empty, return something
    If IsResultEmpty() Then
        boolIsResultEmpty = True
        'handle empty results, return Nothing or empty string/array/file
        Select Case lngReturnType
            Case enumReturnType.Array_Type
                UseExistingFile = arrEmpty
            Case enumReturnType.File_Type
                Set UseExistingFile = Nothing
            Case enumReturnType.ListObject_Type
                Set UseExistingFile = lo
            Case enumReturnType.Path_Type
                UseExistingFile = ""
            Case enumReturnType.Recordset_Type
                Set UseExistingFile = Nothing
            Case enumReturnType.Text_Type
                UseExistingFile = ""
            Case enumReturnType.Range_Type
                Set UseExistingFile = Nothing
            Case enumReturnType.PowerPivotListObject_Type
                Set UseExistingFile = Nothing
        End Select
        GoTo emptyResults
    End If
    'check file size
    If CDbl(fsoFileResult.Size) > CDbl("2000000000") Then
        boolIsResultTooLarge = True
        'report error if file is too large for the ReturnType, only return the path or file
        If lngReturnType <> enumReturnType.Path_Type And lngReturnType <> enumReturnType.File_Type Then
            enumClassError = FileSizeTooLarge
            ReportError
            Exit Function
        End If
    End If
    'choose how to return the results
    Select Case lngReturnType
        Case enumReturnType.Array_Type
            UseExistingFile = ReturnArray
        Case enumReturnType.File_Type
            Set UseExistingFile = ReturnFile
        Case enumReturnType.ListObject_Type
            Set UseExistingFile = ReturnListObject
        Case enumReturnType.Path_Type
            UseExistingFile = ReturnPath
        Case enumReturnType.Recordset_Type
            Set UseExistingFile = ReturnRecordset
        Case enumReturnType.Text_Type
            UseExistingFile = ReturnText
        Case enumReturnType.Range_Type
            Set UseExistingFile = ReturnRange
        Case enumReturnType.PowerPivotListObject_Type
            Set UseExistingFile = ReturnPowerPivotListObject
    End Select
    'print time elapsed
    'Debug.Print "Process took " & Format(Now() - dteTimerStart, "hh:mm:ss")
emptyResults:
    'clean up
    Set fsoFileResult = Nothing
    Set fso = Nothing
End Function
Public Function RunSQLite(strDatabase As String, varSQL As Variant, lngReturnType As enumReturnType, boolHeadersArg As Boolean, Optional varRngOrName As Variant, Optional boolKeepPPQueryArg As Boolean) As Variant
    Dim fso As New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    Dim fsoFileResult As Scripting.File
    Dim fsoFolder As Scripting.Folder
    Dim fsoSQLCmds As Scripting.TextStream
    Dim fsoSQL As Scripting.TextStream
    Dim strSQLCmdFile As String
    Dim arrSQLCmds() As String
    Dim i As Long
    Dim boolExtensionOk As Boolean
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Long: windowStyle = 1
    Dim lo As ListObject
    Dim strShell As String
    Dim strHeaders As String
    Dim strOutput As String
    Dim strDrive As String
    Dim strDatabaseNoExt As String
    Dim arrEmpty() As String
    Dim strSQL As String
    Dim rngTest As Range
    
    'start timer
    dteTimerStart = Now()
    'get strSQl
    If VarType(varSQL) = vbString Then
        'file path
        If fso.FileExists(varSQL) Then
            Set fsoSQL = fso.OpenTextFile(varSQL, ForReading)
            strSQL = fsoSQL.ReadAll
            fsoSQL.Close
            Set fsoSQL = Nothing
        Else
            'sql statement
            strSQL = varSQL
        End If
    Else
        'fso file
        Set fsoSQL = varSQL.OpenAsTextStream
        strSQL = fsoSQL.ReadAll
        fsoSQL.Close
        Set varSQL = Nothing
        Set fsoSQL = Nothing
    End If
    
    'get date for timestamp
    'strTimestamp = Format(Year(Now()), "####") & Format(Month(Now()), "0#") & Format(Day(Now()), "0#") & Format(Hour(Now()), "0#") & Format(Minute(Now()), "0#") & Format(Second(Now()), "0#")
    strTimeStamp = Format(Now(), "yyyymmddhhnnss")
    'strDatabasePublic and other global strings set to ""
    strDatabasePublic = strDatabase
    boolKeepPPQuery = boolKeepPPQueryArg
    'set header choice
    If boolHeadersArg = False Then
        boolHeaders = False
    Else
        boolHeaders = True
    End If
    Select Case boolHeaders
        Case True
            strHeaders = "on"
        Case False
            strHeaders = "off"
    End Select
    If boolHeaders = False And (lngReturnType = enumReturnType.Recordset_Type Or lngReturnType = enumReturnType.ListObject_Type Or lngReturnType = enumReturnType.PowerPivotListObject_Type) Then
        enumClassError = RecordSetNoHeaderError
        ReportError
        Set RunSQLite = Nothing
        Set fso = Nothing
        Set fsoFile = Nothing
        Exit Function
    End If
    'set ListObject destination range
    If Not IsMissing(varRngOrName) Then
        If IsObject(varRngOrName) Then
            On Error Resume Next
            Set rngTest = varRngOrName
            On Error GoTo 0
            If Not rngTest Is Nothing Then Set rngListObjectDestination = varRngOrName
        Else
            strNameArg = varRngOrName
        End If
    End If
    'get database
    If InStr(1, strDatabase, "\", vbTextCompare) = 0 Then
        If InStr(1, ThisWorkbook.Name, "PERSONAL.XLSB", vbTextCompare) > 0 Then
            enumClassError = DBNotFoundinXLStart
            ReportError
            Exit Function
        End If
        strDatabase = ThisWorkbook.Path & "\" & strDatabase
        strDatabasePublic = strDatabase
    Else
        'get file
        Set fsoFile = fso.GetFile(strDatabase)
    End If
    'check if file exists
    If Not fso.FileExists(strDatabase) Then
        enumClassError = DBDoesNotExsist
        ReportError
        Exit Function
    Else
        Set fsoFile = fso.GetFile(strDatabase)
    End If
    'check db3 extension
    boolExtensionOk = False
    For i = LBound(arrSQLiteExtensions) To UBound(arrSQLiteExtensions)
        If InStr(1, Right(strDatabase, Len(strDatabase) - InStrRev(strDatabase, ".") + 1), arrSQLiteExtensions(i), vbTextCompare) > 0 Then
            boolExtensionOk = True
            'get filename without extension
            strDatabaseNoExt = Replace(fsoFile.Name, arrSQLiteExtensions(i), "", , , vbTextCompare)
            i = UBound(arrSQLiteExtensions)
        End If
    Next i
    strFriendlyName = strDatabaseNoExt & strTimeStamp
    'report error if db extension is unrecognized
    If boolExtensionOk = False Then
        enumClassError = DBUnexpectedExtension
        ReportError
    End If
    'setup full filenames
      'SQL commands file - backslashes "\" for normal file usage
    strSQLCmdFile = fsoFile.ParentFolder.Path & "\sqlite3_cmds" & strTimeStamp & ".txt"
      'Output file - add forward slashes ("/") later in array for SQLite commands
    strOutput = fsoFile.ParentFolder.Path & "\SQL_" & strDatabaseNoExt & strTimeStamp & ".pdl"
      'check path lengths
    If Len(strOutput) > 255 Or Len(strSQLCmdFile) > 255 Then
        enumClassError = FilePathsTooLong
        ReportError
    End If
    'get drive letter for command prompt, validate
    strDrive = Left(fsoFile.Path, 2)
    If Right(strDrive, 1) <> ":" Or Asc(Left(strDrive, 1)) < 65 Or Asc(Left(strDrive, 1)) > 90 Then
        enumClassError = InvalidDrive
        ReportError
    End If
    
    'get rid of line breaks, tabs, and duplicate spaces
    strSQL = Replace(strSQL, vbNewLine, "  ", , , vbTextCompare)
    strSQL = Replace(strSQL, vbCrLf, "  ", , , vbTextCompare)
    Do While InStr(1, strSQL, "  ", vbTextCompare) > 0 Or InStr(1, strSQL, Chr(32) & Chr(32), vbTextCompare) > 0
        strSQL = Replace(strSQL, "  ", " ", , , vbTextCompare)
        strSQL = Replace(strSQL, Chr(32) & Chr(32), " ", , , vbTextCompare)
    Loop
    strSQL = Replace(strSQL, "( ", "(", , , vbTextCompare)
    strSQL = Replace(strSQL, " )", ")", , , vbTextCompare)
    
    'ensure SQL statement ends in a semicolon
    If Right(strSQL, 1) <> ";" Then strSQL = strSQL & ";"
    'setup SQL commands
    i = 0: Erase arrSQLCmds
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".mode list": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".headers " & strHeaders: i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".nullval NULL": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".output """ & Replace(strOutput, "\", "/") & """": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = strSQL: i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".exit": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = "": i = i + 1
    'create command file
    Set fsoSQLCmds = fso.CreateTextFile(strSQLCmdFile, True, False)
    'write to command file
    For i = LBound(arrSQLCmds) To UBound(arrSQLCmds)
        fsoSQLCmds.WriteLine arrSQLCmds(i)
    Next i
    fsoSQLCmds.Close
    'run SQL
    strShell = "cmd /c " & strDrive & " & cd """ & fsoFile.ParentFolder.Path & """ & sqlite3 """ & strDatabase & """ "".read " & Replace(fso.GetFileName(strSQLCmdFile), "\", "/") & """"
    wsh.Run strShell, windowStyle, waitOnReturn
    'delete temporary SQL command file
    Kill strSQLCmdFile
    strResultPath = strOutput
    'get result file
    Set fsoFileResult = fso.GetFile(strResultPath)
    'check if file is empty, return something
    If IsResultEmpty Then
        boolIsResultEmpty = True
        'handle empty results, return Nothing or empty string/array/file
        Select Case lngReturnType
            Case enumReturnType.Array_Type
                RunSQLite = arrEmpty
            Case enumReturnType.File_Type
                Set RunSQLite = Nothing
            Case enumReturnType.ListObject_Type
                Set RunSQLite = lo
            Case enumReturnType.Path_Type
                RunSQLite = ""
            Case enumReturnType.Recordset_Type
                Set RunSQLite = Nothing
            Case enumReturnType.Text_Type
                RunSQLite = ""
            Case enumReturnType.Range_Type
                Set RunSQLite = Nothing
            Case enumReturnType.PowerPivotListObject_Type
                Set RunSQLite = Nothing
        End Select
        GoTo emptyResults
    End If
    'check file size
    If CDbl(fsoFileResult.Size) > CDbl("2000000000") Then
        boolIsResultTooLarge = True
        'report error if file is too large for the ReturnType, only return the path or file
        If lngReturnType <> enumReturnType.Path_Type And lngReturnType <> enumReturnType.File_Type Then
            enumClassError = FileSizeTooLarge
            ReportError
            Exit Function
        End If
    End If
    'choose how to return the results
    Select Case lngReturnType
        Case enumReturnType.Array_Type
            RunSQLite = ReturnArray
        Case enumReturnType.File_Type
            Set RunSQLite = ReturnFile
        Case enumReturnType.ListObject_Type
            Set RunSQLite = ReturnListObject
        Case enumReturnType.Path_Type
            RunSQLite = ReturnPath
        Case enumReturnType.Recordset_Type
            Set RunSQLite = ReturnRecordset
        Case enumReturnType.Text_Type
            RunSQLite = ReturnText
        Case enumReturnType.Range_Type
            Set RunSQLite = ReturnRange
        Case enumReturnType.PowerPivotListObject_Type
            Set RunSQLite = ReturnPowerPivotListObject
    End Select
    'delete result file if not needed
    If lngReturnType <> enumReturnType.Path_Type And lngReturnType <> enumReturnType.File_Type And fso.FileExists(strResultPath) Then
        Kill strResultPath
        Set fsoFileResult = Nothing
    End If
emptyResults:
    'clean up
    If boolIsResultEmpty = True And fso.FileExists(strResultPath) Then
        Kill strResultPath
        Set fsoFileResult = Nothing
    End If
    If fso.FileExists(strSQLCmdFile) Then Kill strSQLCmdFile
    Set fsoFile = Nothing
    Set fso = Nothing
    Set fsoSQLCmds = Nothing
    Set fsoFolder = Nothing
    Set rngTest = Nothing
    If strErrMsg <> "" Then MsgBox strErrMsg
    'print time elapsed
    'Debug.Print "Process took " & Format(Now() - dteTimerStart, "hh:mm:ss")
End Function

Public Sub ImportListObjectToDB3(lo As ListObject, strDb3 As String, Optional strTblName As String)
    Dim fso As New Scripting.FileSystemObject
    Dim fsoSQLCmds As Scripting.TextStream
    Dim arrSQLCmds() As String
    Dim strPath As String, strTimeStamp As String, strSQLCmdFile As String, strDrive As String, strPDL As String, strShell As String
    Dim i As Integer
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Long: windowStyle = 1
    
    If strTblName = "" Then strTblName = lo.Name
    strPath = lo.Parent.Parent.Path
    strTimeStamp = Format(Now(), "yyyymmddhhnnss")
    strSQLCmdFile = strPath & "\sqlite3_cmds" & strTimeStamp & ".txt"
    strDrive = Left(strPath, 2)
    
    strPDL = Replace(ExportLOToPDL(lo), "\", "/")
    
    i = 0: Erase arrSQLCmds
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".mode list": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".headers on": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".nullval NULL": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = "drop table if exists " & strTblName & ";": i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".import " & strPDL & " " & strTblName: i = i + 1
    ReDim Preserve arrSQLCmds(i): arrSQLCmds(i) = ".exit": i = i + 1
    'create command file
    Set fsoSQLCmds = fso.CreateTextFile(strSQLCmdFile, True, False)
    'write to command file
    For i = LBound(arrSQLCmds) To UBound(arrSQLCmds)
        fsoSQLCmds.WriteLine arrSQLCmds(i)
    Next i
    fsoSQLCmds.Close
    'run SQL
    strShell = "cmd /c " & strDrive & " & cd """ & strPath & """ & sqlite3 """ & strDb3 & """ "".read " & Replace(fso.GetFileName(strSQLCmdFile), "\", "/") & """"
    
    wsh.Run strShell, windowStyle, waitOnReturn
    'delete temporary SQL command file
    Kill strSQLCmdFile
    Kill strPDL
End Sub
Private Function ExportLOToPDL(lo As ListObject, Optional strPDL As String) As String
    Dim numRows As Long, numCols As Integer
    Dim r As Long, c As Integer
    Dim data
    Dim ExpRng As Range
    Dim s As String

    Set ExpRng = lo.Range
    numCols = ExpRng.Columns.Count
    numRows = ExpRng.Rows.Count
    If strPDL = "" Then strPDL = lo.Parent.Parent.Path & "\" & lo.Name & ".pdl"
    Open strPDL For Output As #1
        For r = 1 To numRows
            For c = 1 To numCols
                If IsEmpty(ExpRng.Cells(r, c)) Then
                    data = ""
                Else
                    data = ExpRng.Cells(r, c).Value
                    'Conversion only needed if doing CSV file
                    'If IsNumeric(Data) Then Data = Val(Data)
                    If IsError(data) Then data = Val(data)
                End If
                If c <> numCols Then
                    Print #1, data & "|";
                Else
                    Print #1, data & ""
                End If
            Next c
        Next r
    Close #1
    ExportLOToPDL = strPDL
End Function
Private Sub ListObjectBorders(r As Range)
    r.Borders(xlDiagonalDown).LineStyle = xlNone
    r.Borders(xlDiagonalUp).LineStyle = xlNone
    With r.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Private Sub ReportError(Optional strErrNumber As String, Optional strErrDescription As String)
    Select Case enumClassError
        Case ClassError.FileSizeTooLarge
            err.Raise intSatan, , "SQLite Error: File size is greater than 2GB. Results can only be returned as a Path or File."
        Case ClassError.ResultFileDoesNotExsist
            err.Raise intSatan, , "SQLite Error: Results file (strResultPath) is not set or it does not exist."
        Case ClassError.RecordSetFieldTooSmall
            err.Raise intSatan, , "SQLite Error: RecordSet field data type is not big enough."
        Case ClassError.DBDoesNotExsist
            err.Raise intSatan, , "SQLite Error: Database not found: " & strDatabasePublic
        Case SQLiteNotInstalled
            err.Raise intSatan, , "SQLite Error: SQLite3.exe not found (C:\cygwin\bin\sqlite3.exe)."
        Case RecordSetNoHeaderError
            err.Raise intSatan, , "SQLite Error: To return a RecordSet/ListObject, headers must be turned on."
        Case DBNotFoundinXLStart
            err.Raise intSatan, , "SQLite Error: The SQLite object is being created from the Personal Workbook. Please provide the entire database path."
        Case DBUnexpectedExtension
            err.Raise intSatan, , "SQLite Error: Database file has unexpected extension: " & strDatabasePublic
        Case FilePathsTooLong
            err.Raise intSatan, , "SQLite Error: File paths are too long."
        Case InvalidDrive
            err.Raise intSatan, , "SQLite Error: Invalid drive: " & strDatabasePublic
        Case DestinationRangeInATable
            err.Raise intSatan, , "SQLite Error: Destination range is already in a table.(" & rngListObjectDestination.Parent.Name & " " & rngListObjectDestination.Address & ")"
        Case ExistingFileExtensionNotRecognized
            err.Raise intSatan, , "SQLite Error: File Extension not recognized.: " & strResultPath
        Case UnidentifiedCustomFormattingCharacter
            err.Raise intSatan, , "SQLite Error: Unidentified custom formatting character placement."
        Case ClassError.UnknownError
            'return actual err.Number and err.Description
            err.Raise err.Number, , err.Description
            err.Raise strErrNumber, , strErrDescription
    End Select
    'reset error enum
    err.Clear
    enumClassError = NoErrors
End Sub



