VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmExportMSAccessMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'***********************************************************************************************************
'Export Access Metadata
'   based on Export Access Metadata work by Hal Beresford
'   and modified in 2012
'
' Functionality added in this version
'   1) Automation in adding necessary References that are used by this form and prioritize them
'   2) Retrieving and associating a label, that is likely related to a given field, based on "fuzzy" string
'       match
'           - Removing two fields from previous version: fieldCaption and fieldRowSource, because they don't
'               fit our use cases
'           - Adding label description to fieldDescription column
'           - Adding labelMatchScore column for the score of doing a "fuzzy" string match (We can look it as
'               a confidence value for how confident is it to select the label for a field name
'           - Adding matchedLabelIdx column for the index number of the matched label
'   3) Printing out all the possible labels from the Access forms
'           - labelIdx column going from 0 to n
'           - formName column is the name of the form where this label is shown in the Access form
'           - labelName column is the name of the label
'           - labelCaption column is the label caption
'
'
'   version 1.0.0
'   since 9/17/2012 by shw
'***********************************************************************************************************

Private Sub cmdExportMetadata_Click()
    loadNecessaryReferences
    exportAccessMetadata
End Sub


Public Sub getAllLabelsFromAllForms()
On Error GoTo ErrorHandler:
    Dim afrm As AccessObject
    Dim frm As Form
    Dim openFrmError As Boolean
    ReDim labels(1) As Control
    
    For Each afrm In CurrentProject.AllForms
        openFrmError = False
        If Not afrm.IsLoaded Then
            DoCmd.OpenForm afrm.Name, acDesign, , , , acHidden
        End If
        
        If Not openFrmError And afrm.Name <> "frmExportMSAccessMetadata" Then
            Set frm = Forms(afrm.Name)
            Call getAllLabelsInForm(frm)
        End If
    Next afrm
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 7784 Then
        openFrmError = True
        Resume Next
    Else
        MsgBox "Sub getAllLabelsFromAllForms Error. Code: " & Err.Number & ", Description: " & Err.Description
    End If
    
    Exit Sub
End Sub

Public Sub getAllLabelsInForm(frm As Form)
On Error GoTo thisError
    Dim ctrl As Control
    
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is Label Then
            If Len(ctrl.Caption) > 0 Or Len(ctrl.Name) > 0 Then
                ReDim Preserve labels(labelIdx) As Control
                Set labels(labelIdx) = ctrl
                labelIdx = labelIdx + 1
            End If
        ElseIf ctrl.ControlType = acSubform Then
            Call getAllLabelsInForm(ctrl.Form)
        Else
        End If
    Next ctrl
    Exit Sub
    
thisError:
    Debug.Print "Sub getAllLabelsInForm Error. Code: " & Err.Number & ", Description: " & Err.Description
    Exit Sub
End Sub

'***********************************************************************************************************
' Utilized Fuzzy scoring algorithm at
'   http://code.google.com/p/fast-vba-fuzzy-scoring-algorithm/source/browse/trunk/Fuzzy1
'
'   Ultimately, the algorithm is fast and easy to use. For a better performed string match, algorithm like
'       Needleman-Wunsch Algorithm or some sequence matching algorithm should be implemented and used.
'
'   Version 1.0.0
'   since 9/17/2012 by shw
'***********************************************************************************************************

Private Function fuzzySearchLabel(dbTblColName As String, returnObj() As String)
On Error GoTo thisError
    Dim inStrIdx As Integer
    Dim ctrl As Control
    Dim tmpMatchScore As Single
    Dim tmpNameMatchScore As Single
    Dim maxMatchScore As Single
    Dim matchLabel As String
    Dim matchLabelIdx As Integer
    matchLabel = ""
    Dim lbnd As Integer
    lbnd = LBound(modFuzzy.labels)
    Dim ubnd As Integer
    ubnd = UBound(modFuzzy.labels)
    
    For inStrIdx = lbnd To ubnd
        Set ctrl = modFuzzy.labels(inStrIdx)
        tmpMatchScore = modFuzzy.fuzzy(dbTblColName, ctrl.Caption)
        tmpNameMatchScore = modFuzzy.fuzzy(dbTblColName, ctrl.Name)
        If maxMatchScore < tmpMatchScore Then
            maxMatchScore = tmpMatchScore
            matchLabel = ctrl.Caption
            matchLabelIdx = inStrIdx
        End If
        If maxMatchScore < tmpNameMatchScore Then
            maxMatchScore = tmpNameMatchScore
            matchLabel = ctrl.Caption
            matchLabelIdx = inStrIdx
        End If
    Next inStrIdx
    
'    Dim returnObj(2) As String
    returnObj(0) = CStr(matchLabelIdx)
    returnObj(1) = matchLabel
    returnObj(2) = CStr(maxMatchScore)
    
'    fuzzySearchLabel = returnObj
    
thisExit:
    Exit Function
    
thisError:
    MsgBox "Function fuzzySearchLabel Error. Code: " & Err.Number & ". description: " & Err.Description
    GoTo thisExit
    
End Function

'***********************************************************************************************************
'exportAccessMetadata()
'November 11, 2010
'
'Functions: 1) Chooses table name not already in use.  Example name: ztblMetadata_MSAccessToExcel_2010_11_11,
'               with possible "#" characters at the end.
'           2) Creates the metadata table.
'           3) Queries table MSysObjects, populates the metadata table with names of database tables.
'           4) Queries database tables, populates the metadata table with each table's variables.
'           5) Exports the metadata table to format of user's choice.
'           6) Deletes the metadata table from MS Access database.
'
'Output:    your database's metadata in Excel or other user-chosen format
'
'History:
'Original code found at http://www.access-programmers.co.uk/forums/showthread.php?t=99194.  Author raskew.
'Hal Beresford later modified that code to fix bugs and add functionality.
'***********************************************************************************************************

Private Sub exportAccessMetadata()
On Error GoTo thisErr

    'Declare variables
    Dim db As DAO.Database      'database
        'FIRST RECORDSET
    Dim rs As Recordset         'recordset object
    Dim strSQL As String        'SQL String
        'SECOND RECORDSET
    Dim rs2 As Recordset        'recordset object
    Dim metadataTable As String 'Name of metadata table to be created
    Dim strSQL2 As String       'SQL String
        'TABLES & FIELDS
    Dim td As TableDef          'table object
    Dim tName As String         'table name
    Dim fld As Field            'field object
    Dim fType As String         'field type
    Dim i As Long               'integer to store table number
    Dim guessFieldLabel As String   'guessing the label of a field
    Dim ctrl As Control         ' Control object for label
    
    
    Call getAllLabelsFromAllForms

    'Initialize Variables
    Set db = CurrentDb
    strSQL = "SELECT MSysObjects.Name, MSysObjects.Type From MsysObjects WHERE " & _
                "((MSysObjects.Type)=1 OR (MSysObjects.Type)=6) " & _
                "ORDER BY MSysObjects.Name;"
    Set rs = db.OpenRecordset(strSQL)
    metadataTable = "ztblMetadata_MSAccessToExcel_" & Year(Now()) & "_" & Month(Now()) & "_" & Day(Now())
    i = 0
    
    'For new metadata table, find table name that is not already in use
    i = objectExists(metadataTable)
    Do While i = 1
        metadataTable = metadataTable & "#"
        i = objectExists(metadataTable)
    Loop
    
    'Create new metadata table
    ' Original SQL
    ' db.Execute "CREATE TABLE " & metadataTable & "(fieldName TEXT (55), tableName TEXT (55), fieldType TEXT (20), fieldSize Long,  fieldDescription TEXT (255), fieldCaption TEXT (255), fieldRowSource MEMO);"
        
    db.Execute "CREATE TABLE " & metadataTable & "(" & _
        "fieldName TEXT (55), " & _
        "tableName TEXT (55), " & _
        "fieldType TEXT (20), " & _
        "fieldSize LONG, " & _
        "fieldDescription TEXT (255), " & _
        "labelMatchScore DOUBLE, " & _
        "matchedLabelIdx LONG, " & _
        "labelIdx LONG, " & _
        "formName TEXT (255), " & _
        "labelName TEXT (255), " & _
        "labelCaption TEXT (255)" & _
    ");"
On Error Resume Next
    'Initialize rs2 Variable
    Set rs2 = db.OpenRecordset(metadataTable, dbOpenDynaset)
    
    'Cycle through system tables and figure out pkeys and number of records
    Dim fldIdx As DAO.Index
    Dim pkeyFld As DAO.Field
    Dim pkeyFldName As String
    
    rs.MoveFirst
    Do While Not rs.EOF
        'For tables that are not system tables
        If Left(rs!Name, 4) <> "MSys" And Left(rs!Name, 1) <> "~" Then
            'Identify the table number
            tName = rs!Name
            i = findTdfNumber(tName)
            If i <> -999 Then
                pkeyFldName = ""
                'Cycle through all fields in the table, populating ztblMetadata along the way
                Set td = db.TableDefs(i)
                
                For Each fldIdx In td.Indexes
                    If fldIdx.Primary Then
                        For Each pkeyFld In fldIdx.Fields
                            pkeyFldName = pkeyFldName & pkeyFld.Name & ", "
                        Next pkeyFld
                    End If
                 Next fldIdx
                    
                rs2.AddNew
                    rs2!FieldName = pkeyFldName
                    rs2!tableName = tName
                    rs2!FieldSize = td.RecordCount
                rs2.Update
            End If
        End If
        rs.MoveNext
    Loop
    
    
    'Cycle through system table and read and write db tables metadata
    Dim returnGuessObj(2) As String
    rs.MoveFirst
    Do While Not rs.EOF
        'For tables that are not system tables
        If Left(rs!Name, 4) <> "MSys" And Left(rs!Name, 1) <> "~" Then
            'Identify the table number
            tName = rs!Name
            i = findTdfNumber(tName)
            If i <> -999 Then
                'Cycle through all fields in the table, populating ztblMetadata along the way
                Set td = db.TableDefs(i)
                For Each fld In td.Fields
                    fType = acFieldType(fld.Type)
                    
                    fuzzySearchLabel fld.Name, returnGuessObj
                    rs2.AddNew
                        rs2!FieldName = fld.Name
                        rs2!tableName = tName
                        rs2!fieldType = fType
                        rs2!FieldSize = fld.Size
                        rs2!fielddescription = returnGuessObj(1)
                        If Not IsNumeric(returnGuessObj(2)) Or Len(returnGuessObj(2)) < 1 Then
                            rs2!labelMatchScore = 0
                        Else
                            rs2!labelMatchScore = CDbl(returnGuessObj(2))
                        End If
                        rs2!matchedLabelIdx = CLng(returnGuessObj(0))
                        rs2!fieldCaption = fld.Properties("caption")
                        rs2!fieldRowSource = fld.Properties("rowsource")
                    rs2.Update
                 Next fld
            End If
        End If
        rs.MoveNext
    Loop
    
    Dim lbnd As Integer
    Dim ubnd As Integer
    lbnd = LBound(modFuzzy.labels)
    ubnd = UBound(modFuzzy.labels)
    Dim parentCtrlFormName As String
    
    For i = lbnd To ubnd
        Set ctrl = modFuzzy.labels(i)
        parentCtrlFormName = ctrl.Properties.Parent.FormName
        rs2.AddNew
            rs2!labelIdx = CLng(i)
            rs2!FormName = parentCtrlFormName
            rs2!labelName = ctrl.Name
            rs2!labelCaption = ctrl.Caption
        rs2.Update
    Next i
    
    
    'Export new metadata table
    DoCmd.OutputTo acOutputTable, metadataTable
    
    'Close recordset and database objects
    rs.Close
    rs2.Close
    db.Close
    
    'Delete new metadata table
    DoCmd.DeleteObject acTable, metadataTable

thisExit:
    Exit Sub

thisErr:
    'Alert user of error
    MsgBox "Operation not completed.  Error Code: " & Err.Description
    'Close recordset and database objects
    rs.Close
    db.Close
    GoTo thisExit

End Sub


'*************************************************************************************
'findTdfNumber(tblName As String) As Long
'by Hal
'November 10, 2010
'
'Returns the table number of a particular table name within the current database.
'Returns -999 if table name not found.
'
'*************************************************************************************

Private Function findTdfNumber(tblName As String) As Long
On Error GoTo thisErr

    'variable declaration
    Dim db As DAO.Database
    Dim tdf As TableDef
    Dim i, num As Long
    
    'initialization
    Set db = CurrentDb
    Set tdf = db.TableDefs(0)
    num = CurrentDb.TableDefs.Count
    
    'find table number
    Do While Not tblName = tdf.Name
        i = i + 1
        Set tdf = db.TableDefs(i)
    Loop


thisExit:
    'close database object
    db.Close
    'return item number
    findTdfNumber = i
    Exit Function
    
thisErr:
    i = -999
    GoTo thisExit
    
End Function


'***********************************************************************************************************
'acFieldType(intType As Integer) As String
'November 10, 2010
'
'Returns text name of field type
'
'History:
'Original code found at http://www.access-programmers.co.uk/forums/showthread.php?t=99194.  Author raskew.
'Hal later modified the code.
'***********************************************************************************************************

Private Function acFieldType(intType As Integer) As String
    Dim varResult As String

    Select Case intType
        Case dbBoolean
            varResult = "Yes/No"            '1.  aka "Boolean"
        Case dbByte
            varResult = "Byte"              '2
        Case dbInteger
            varResult = "Integer"           '3
        Case dbLong
            varResult = "Long Integer"      '4.  aka "Long"
        Case dbCurrency
            varResult = "Currency"          '5
        Case dbSingle
            varResult = "Single"            '6
        Case dbDouble
            varResult = "Double"            '7
        Case dbDate
            varResult = "Date/Time"         '8
        Case dbBinary
            varResult = "Binary"            '9
        Case dbText
            varResult = "Text"              '10
        Case dbLongBinary
            varResult = "Long Binary"       '11  aka "LongBinary"
        Case dbMemo
            varResult = "Memo"              '12
        Case dbGUID
            varResult = "GUID"              '15
    End Select
    
    acFieldType = varResult
End Function

'*************************************************************************************
'objectExists(strObjectName As String)
'by Hal
'October 7, 2010
'
'Function takes the name of an object.  Returns 1 if the object exists, 0 if not.
'
'*************************************************************************************
Private Function objectExists(strObjectName As String)
On Error Resume Next

    'declare variables
    Dim db As DAO.Database
    Dim rs As Recordset

    'initialize variables
    Set db = CurrentDb
    Set rs = db.OpenRecordset("MSysObjects", dbReadOnly)
    
    'see if table exists
    rs.FindFirst "Name = '" & strObjectName & "'"
    If rs.NoMatch Then
        objectExists = 0
    Else
        objectExists = 1
    End If

    'close recordset
    rs.Close
    db.Close
    
End Function

Private Sub loadVBExtRef()
    Dim i As Integer
    Dim vbideRefFound As Boolean
    vbideRefFound = False

    For i = 1 To Application.VBE.ActiveVBProject.References.Count
        If Application.VBE.ActiveVBProject.References(i).Name = "VBIDE" Then
            vbideRefFound = True
        End If
    Next i
    
    If Not vbideRefFound Then
        Debug.Print "Adding VBIDE Ref. References count: " & Application.VBE.ActiveVBProject.References.Count
        Application.VBE.ActiveVBProject.References.AddFromGuid Guid:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    End If
    
End Sub

Private Sub removeReferences(refs() As Reference)
On Error GoTo ErrorHandler:
    Dim i As Integer
    Dim ref As Reference
    
    For i = 0 To UBound(refs)
        Set ref = refs(i)
        Debug.Print "ref(" & i & ")/" & UBound(refs) & ": " & ref.Name
        If ref.Name <> "VBIDE" Then
            References.Remove ref
        End If
    Next i
    
    Debug.Print "Sub removeReferences(refs) ENDS"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Sub removeReferences(refs) Error. Code: " & Err.Number & ", Description: " & Err.Description
    Resume Next
End Sub

Public Sub loadNecessaryReferences()
On Error GoTo ErrorHandler:
    Dim i As Integer
    Dim trRefIdx As Integer
    Dim daoRefFound As Boolean
    Dim trailingRefs() As Reference
    Dim refCopy() As String
    
    trRefIdx = -1
    daoRefFound = False
    
    ' Add VBIDE ref because it is required for retreiving Access metadata
    Call loadVBExtRef
        
    ReDim trailingRefs(0) As Reference
    ReDim refCopy(0) As String
    
    For i = 1 To References.Count
        If References(i).Name = "stdole" Then
            trRefIdx = 0
        Else
            If References(i).Name = "DAO" Then
                daoRefFound = True
            Else
                If trRefIdx > -1 And References(i).Name <> "VBIDE" Then
                    ReDim Preserve trailingRefs(trRefIdx) As Reference
                    Set trailingRefs(trRefIdx) = References(i)
                    trRefIdx = trRefIdx + 1
                End If
            End If
        End If
'        Debug.Print Application.VBE.ActiveVBProject.References(i).Name & "; " & Application.VBE.ActiveVBProject.References(i).Description & "; " & Application.VBE.ActiveVBProject.References(i).Guid & "; " & Application.VBE.ActiveVBProject.References(i).Major & "; " & Application.VBE.ActiveVBProject.References(i).Minor
    Next i
    
    If UBound(trailingRefs) > 0 Then
        ReDim refCopy(UBound(trailingRefs), 2)
        For i = 0 To UBound(trailingRefs)
            refCopy(i, 0) = trailingRefs(i).Guid
            refCopy(i, 1) = CStr(trailingRefs(i).Major)
            refCopy(i, 2) = CStr(trailingRefs(i).Minor)
        Next i
    
        Call removeReferences(trailingRefs)
    End If
    
    ' Adding DAO ref "explicitly", because we want to guarrantte the priority
    If Not daoRefFound Then
        References.AddFromGuid Guid:="{00025E01-0000-0000-C000-000000000046}", Major:=5, Minor:=0
    End If
    
    If UBound(refCopy) > 0 Then
        For i = 0 To UBound(refCopy)
            References.AddFromGuid refCopy(i, 0), CLng(refCopy(i, 1)), CLng(refCopy(i, 2))
        Next i
    End If
        
    Debug.Print "Final References count: " & Application.VBE.ActiveVBProject.References.Count
    Exit Sub
    
ErrorHandler:
    MsgBox "Sub loadReferences Error. Code: " & Err.Number & ", Description: " & Err.Description
    Exit Sub

End Sub

