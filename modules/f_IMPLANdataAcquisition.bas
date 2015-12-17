Attribute VB_Name = "f_IMPLANdataAcquisition"
Dim myfile As String
'==========================================================
'  procedures to acquire the data from an IMPLAN data file
'==========================================================
Sub retrieveImplanData()
'lastModified 20151107
    Dim sql As String, myq As String, msg As String, msg0 As String
'define alert; if datasheet is present, add an alert that it will be erased
    msg = "This assumes you have created your (aggregated) I x I matrix in IMPLAN." & vbCr & _
            "See FAQ for more information if this is unclear..."
    If WorksheetExists("DataSheet") Then msg = msg & vbCrLf & vbCrLf & _
            "This will also erase your current existing SAM!" & vbCrLf & "CANCEL and save it somewhere first if you want to keep it..."
    If MsgBox(msg, vbCritical + vbOKCancel) = vbCancel Then Exit Sub
'select IMPLAN export file; wait with deleting datasheet until there is a working IMPLAN file....
    ChDir ActiveWorkbook.Path
    myfile = Application.GetOpenFilename( _
        fileFilter:="IMPLAN (*.impdb; *.iap),*.impdb;*.iap", _
        title:="Please Select the datafiles you wish to update from...")
    If (myfile = "" Or myfile = "False") Then _
        GoTo exit0 'exit without clearing datasaheet
    Application.ScreenUpdating = False
'wayfinding depends on the IMPLAN version -- IMPLAN Pro 2.0 [.iap] became available in May 1999,[5] and IMPLAN Version 3.0 was released on November 3, 2009[impdb]
    Select Case Right(myfile, 4)
        Case Is = ".iap": 'older version
            ixi = "Regional SAM Balances IxI Industry Detail" 'table
            typecodes = "type codes" 'table
            typecode = "type code" 'field name
            typeDescr = "type" 'field
            employment = "SAEmployment" 'table
            Ipayments = "Institution Payments"
            Ireceipts = "Institution Receipts"
            Icode = "Industry code"
        Case Is = "mpdb"
            ixi = "RegionalSAMBalancesIxIIndustryDetail"
            typecodes = "TypeCodesAll"
            typecode = "typecode"
            typeDescr = "type"
            employment = "StudyAreaEmployment"
            Ipayments = "InstitutionPayments"
            Ireceipts = "InstitutionReceipts"
            Icode = "IndustryCode"
    End Select
    sql = "select [" & ixi & "].* from [" & ixi & "]"
    If checkIImatrix(myfile, sql) = False Then
        MsgBox "I could not find any IxI data in your file; make sure to run the IxI matrix in IMPLAN..."
        GoTo exit0
    End If
'now clear datasheet, if existent
    If WorksheetExists("DataSheet") Then _
        ClearMatrices (True)
    msg = "Acquiring data: "
'execute remote queries in the ACCESS database (=IMPLAN export file)
    For i = 1 To 4
        Select Case i
        Case 1: myq = "U-query"
            sql = "SELECT [" & ixi & "].[" & Ipayments & "], [" & ixi & "].[" & Ireceipts & "], 1000*[Value] AS Kvalue, [" & typecodes & "].Description, [" & typecodes & "].[" & typecode & "] " & _
                    " FROM [" & ixi & "] INNER JOIN [" & typecodes & "] ON [" & ixi & "].[" & Ireceipts & "] = [" & typecodes & "].[" & typecode & "] WHERE (((1000*[Value])<>0));"
        Case Is = 2: myq = "V-Query"
            sql = "TRANSFORM Sum([U-query].Kvalue) AS SumOfKvalue" & _
                " SELECT [U-query].[" & Ireceipts & "] FROM [U-query] GROUP BY [U-query].[" & Ireceipts & "]" & _
                " PIVOT [U-query].[" & Ipayments & "];"
        Case Is = 3: myq = "W-query"
            sql = "SELECT  [" & typecodes & "].Description, [" & typecodes & "]." & typeDescr & " as [type],[V-Query].* FROM [V-Query] INNER JOIN [" & typecodes & "] ON [V-Query].[" & Ireceipts & "] = [" & typecodes & "].[" & typecode & "] ORDER BY [" & typecodes & "].[" & typecode & "];"
        Case Is = 4: myq = "Z-Empl"
            sql = "SELECT [" & employment & "].Employment" & _
            " FROM [" & employment & "] INNER JOIN [" & typecodes & "] ON [" & employment & "].[" & Icode & "] = [" & typecodes & "].[" & typecode & "]" & _
            " WHERE ((([" & employment & "].Employment)<>0)) ORDER BY [" & employment & "].[" & Icode & "];"
        End Select
        msg = msg & " - " & myq & vbCrLf
        showinfo msg
        createQDF myfile, sql, myq
    Next
        msg = msg & " - Importing SAM data" & vbCrLf
        showinfo msg
'import the data (recordsets)
    If CopyFromRecordSet(myfile, "W-Query") = False Then Exit Sub
        msg = msg & " - Importing Employment data" & vbCrLf
        showinfo msg
    If CopyFromRecordSet(myfile, "Z-empl") = False Then Exit Sub
'cleanup
    cleanupImport
    Application.ScreenUpdating = True
    showinfo "", True
    If MsgBox("Import complete." & vbCr & _
        "Would you like to continue and create your Matrices?", vbYesNo) = vbYes _
    Then: createSam
exit0:
    Application.ScreenUpdating = True
    showinfo "", True
    Exit Sub
End Sub
Function checkIImatrix(myfile, sql)
'lastModified 20100518
'===make sure we can actually retrieve the data; i.e. the tables are present
    checkIImatrix = False
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cnstr As String
    Set cn = New ADODB.Connection
    cnstr = "Driver={Microsoft Access Driver (*.mdb)};" & _
                     "Dbq=" & myfile & ";"
    cn.Open cnstr
    Set rs = New ADODB.Recordset
    rs.Open sql, cn
    If Not rs.EOF Then checkIImatrix = True
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Function
Function createQDF(myfile, strsql, qname)
'lastModified 20151107
    On Error Resume Next
    Dim db As DAO.Database
    Set db = OpenDatabase(myfile)
    sql = "drop table [" & qname & "]"
        db.Execute sql, dbFailOnError
    Set qdf = db.CreateQueryDef(qname, strsql)
    qdf.sql = sql
    qdf.Execute
    Set qdf = db.CreateQueryDef(qname, strsql)
    db.Close
    Set db = Nothing
End Function

Function CopyFromRecordSet(myfile, myq)
'lastModified 20151107
    On Error GoTo relink
'sometimes it just needs a little time to think, so we'll let it loop for up to 100 times...
    tryconnect = 0
relink:
    tryconnect = tryconnect + 1
    If tryconnect > 100 Then
        MsgBox "Somehow I cannot lock onto the file; try restarting excel"
        GoTo error
    End If
    Dim db As DAO.Database
    Set db = OpenDatabase(myfile)
    On Error GoTo 0 'stop error trapping
    Dim intColIndex As Integer
    Dim rs As DAO.Recordset
    Select Case myq
    Case Is = "W-Query":
        Set TargetRange = Sheets("SAM>>").Cells(1, 1)
        sql = "select " & myq & ".* from [" & myq & "]"
        Set rs = db.OpenRecordset(sql)
    ' write field names
        For intColIndex = 0 To rs.Fields.Count - 1
             TargetRange.Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
        Next
    ' write recordset
        TargetRange.Offset(1, 0).CopyFromRecordSet rs
    Case Is = "Z-empl"
        Set TargetRange = Sheets("inputEMPL").Cells(1, 3)
        Set rs = db.OpenRecordset("select * from [" & myq & "]")
        TargetRange.Offset(1, 0).CopyFromRecordSet rs
    End Select
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    CopyFromRecordSet = True
    Exit Function
error:
    CopyFromRecordSet = False
End Function

Sub cleanupImport()
'===lastModified 20100518===========================
'bring sector descriptions from SAM to the inputEMPL
    Sheets("SAM>>").Columns("B:C").Copy
    Sheets("inputEmpl").Visible = True
    Sheets("inputEmpl").Select
    Columns("A:B").Select
'add type "Industry" if blanc
    ActiveSheet.Paste
    If Cells(2, 1) = "" Then _
        Range(Cells(2, 1), Cells(Range("A:A").Find(what:="factors", After:=Cells(2, 1), SearchDirection:=1).Row - 1, 1)) = "industry"
'add data label
    Cells(1, 3) = "Gross Employment"
'cleanup SAM
    Sheets("SAM>>").Select
    Columns("B:C").Select
    Selection.delete Shift:=xlToLeft
    Sheets("tools").Select
End Sub
