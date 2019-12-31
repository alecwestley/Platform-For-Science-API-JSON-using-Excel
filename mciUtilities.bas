Attribute VB_Name = "mciUtilities"
'++
'   Title:          mciUtils
'
'   Purpose:        Helper functionality to keep the code simpler/shorter
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       Subs and functions that are not specific to this applicaiton
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        06Nov2014
'
'   Modified:       10MAY2016
'
'   Copyright:
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Version, Note, who
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 0.1 - Complete rework
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Option Explicit

Public Function ciDoGetToken() As String
'++
' get the security token to use for passing info back and forth
'--
Dim sToken As String
Dim sExpSampleType As String
Dim sBaseUrl As String

On Error GoTo ciDoGetToken_Error

sToken = gjSessionID

If sToken = "" Then
    frmci_Login.plAppTitle = zAppTitle
    frmci_Login.plAppVersion = zAppVersion
    frmci_Login.Show
    'launch the login form and get the token back
    sToken = frmci_Login.pgToken
    sExpSampleType = frmci_Login.pgExpSampleType
End If

gjSessionID = sToken

If sExpSampleType = "" Then
    gExpSampleType = mExpSampleType
Else
    gExpSampleType = sExpSampleType
End If

ciDoGetToken = gjSessionID

ciDoGetToken_Exit:
    Exit Function

ciDoGetToken_Error:
    Call ciError("ciDoGetToken - " & Err.Description, Err.Number)
    Resume ciDoGetToken_Exit
End Function


Public Sub ciDoLogIn()
'++
' Call the SKD Logout
'--

Dim sJSON As String
Dim sString As String
Dim sJsonLogout As String

'On Error GoTo ciDoLogIn_Error
On Error Resume Next

'Get the json
'gjSessionID = ""

Call ciDoGetToken

gBaseUrl = fmciaOptions.pgServerUrl

ciDoLogIn_Exit:
    On Error Resume Next
    Exit Sub

ciDoLogIn_Error:
    ciError Err.Description, Err.Number, "ciDoLogIn"
    Resume ciDoLogIn_Exit
    
End Sub


Public Sub ciError(vText As String, _
                   Optional vError As Long, _
                   Optional vFromName As String, _
                   Optional vIsNotErr As Boolean)
'++
' Handler for err text, simples
'--
Dim sPad As String
Dim sName As String
Dim oType As VbMsgBoxStyle

On Error Resume Next

If vError <> 0 Then
    sPad = " " & Err.Number
End If

If vFromName <> "" Then
    sName = "This error came from " & vFromName & " - "
Else
    sName = vFromName
End If

If vIsNotErr Then
    oType = vbCritical
Else
    oType = vbInformation
End If

If Not vIsNotErr Then
    Select Case vError
        Case 0
            sPad = " - cancelled opperation" & sPad
        
    End Select
End If
    MsgBox sName & vText & sPad, , zAppTitle
End Sub


Public Function ciaGetColumLetter(theAddress As String) As String
'++
' Get the column letter (more than A ..)
'--
On Error GoTo ciaGetColumLetter_Error
' Because .Address is $<columnletter>$<rownumber>, drop the first

' character and the characters after the column letter(s).

ciaGetColumLetter = Mid(theAddress, InStr(theAddress, "$") + 1, InStr(2, theAddress, "$") - 2)
ciaGetColumLetter_Exit:
    Exit Function

ciaGetColumLetter_Error:
    'Debug.Print Err.Description
    'Stop
    Resume ciaGetColumLetter_Exit
End Function

Public Sub ciaClearDebug()
'++
' Debug print only the line
'--
On Error Resume Next

Application.SendKeys "^g ^a {DEL}"

End Sub

Public Sub ciaDoWriteFile(vText As String, _
                       vFileName As String)
'++
' Write the file out
'--
Dim sFilePath As String
Dim lLastCol As Long
Dim lLastRow As Long
Dim sString As String
Dim lRow As Long
Dim lCol As Long


On Error GoTo ciaDoWriteFile_Error

sFilePath = Application.DefaultFilePath & "\" & vFileName '& ".txt"

Open sFilePath For Output As #1

sString = ""

Write #1, vText

ciaDoWriteFile_Exit:
    On Error Resume Next
    Close 1
    Exit Sub

ciaDoWriteFile_Error:
    'Debug.Print Err.Description
    Resume ciaDoWriteFile_Exit
End Sub

Public Sub DoHideColumn(vColName As String)
'++
' Hide the given column
' Find the column and hide it
'--
Dim oRange As Range
Dim theCell As String

On Error GoTo DoHideColumn_Error

theCell = ciaFindCell(vColName)

Set oRange = Range(theCell)

oRange.Select

Selection.EntireColumn.Hidden = True

DoHideColumn_Exit:
    On Error Resume Next
    Set oRange = Nothing
    Exit Sub

DoHideColumn_Error:
    'Debug.Print "DoHideColumn - " & Err.Description
    'Stop 'Here
    Resume DoHideColumn_Exit

End Sub


Public Sub DoHideExpTempHeadder()
'++
' Hide the standard Experiment template headder columns from the 'user'
' see the constants module for the list
'--
Dim sExpCol2 As String
Dim sExpCol3 As String
Dim sExpCol4 As String

On Error GoTo DoHideExpTempHeadder_Error


sExpCol2 = gExpSample.AttributeFromName(zExpShtCol2)
sExpCol3 = gExpSample.AttributeFromName(zExpShtCol3)
sExpCol4 = gExpSample.AttributeFromName(zExpShtCol4)


DoHideColumn (sExpCol2)
DoHideColumn (sExpCol3)
DoHideColumn (sExpCol4)

DoHideExpTempHeadder_Exit:
    On Error Resume Next

    Exit Sub

DoHideExpTempHeadder_Error:
    'Debug.Print "DoHideExpTempHeadder_Error - " & Err.Description
    Resume DoHideExpTempHeadder_Exit

End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


Public Function ciSendHTTPRequest(vJson As String, _
                         Optional vRequestType As String) As String
'++
' Interface to whatever HTTP blurb functionality that actually works!
'--
'
'all the tech here could change. So not a long term solution

Dim sResource As String
Dim oHttp As MSXML2.XMLHTTP60
Dim sResponse As String

On Error GoTo ciSendHTTPRequest_Error

Set oHttp = New MSXML2.XMLHTTP60

sResource = vJson

'Debug.Print sResource

oHttp.Open "POST", gBaseUrl & "/sdk" & ";jsessionid=" & gjSessionID, False

oHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

oHttp.sEnd "json=" & sResource

sResponse = oHttp.ResponseText

'Debug.Print sResponse

ciSendHTTPRequest = sResponse

ciSendHTTPRequest_Exit:
    On Error Resume Next
    Set oHttp = Nothing
'    'Debug.Print "----------------------------- Start - " & Now()
'    'Debug.Print sResponse
'    'Debug.Print "----------------------------- End   - " & Now()
'    Stop
    Exit Function
    
ciSendHTTPRequest_Error:
    ciError Err.Description, Err.Number
    Resume ciSendHTTPRequest_Exit
End Function
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Function ciCreateNewContainer(vContainerType As String, _
                                     Optional vProjectID As String, _
                                     Optional vlocationId As String, _
                                     Optional theToken As String) As String
'++
' Get a new container based on an Entity Type Name
'--
Dim sExpID As String
Dim sJSON As String
Dim sResponse As String
Dim sString As String
Dim oEntity As cciEntity

On Error GoTo ciCreateNewContainer_Error

'get the experiment creation json

Set oEntity = New cciEntity

Let oEntity.sdkCmd = "create"
Let oEntity.AddSuperType = "CONTAINER"
Let oEntity.AddEntityType = vContainerType
sJSON = oEntity.JSONCreate

sResponse = ciSendHTTPRequest(sJSON)

'get the new id from the response
ciCreateNewContainer = ciGetThingFromJson(sResponse, "barcode")

ciCreateNewContainer_Exit:
    On Error Resume Next
    Set oEntity = Nothing
    Exit Function

ciCreateNewContainer_Error:
    ciError "Get new Experiment" & Err.Description, Err.Number, "ciCreateNewContainer"
    Resume ciCreateNewContainer_Exit
    
End Function



Public Function ciGetThingFromJson(vJson As String, _
                                   vThing As String, _
                          Optional vLastBarcode As Boolean, _
                          Optional vStartPos) As String
'++
' Get the barcode for the sample lot from the json
'--
Dim sThing As String
Dim lStartThingName As Long
Dim lStartThing As Long
Dim lPosEnd As Long
Dim sString As String
Dim lLenJSON As Long

On Error GoTo ciGetThingFromJson_Error

If vLastBarcode Then
    lLenJSON = Len(vJson) - 100
Else
    lLenJSON = 1
End If

lStartThingName = InStr(lLenJSON, vJson, vThing)


If lStartThingName > 0 Then
    
    lStartThing = lStartThingName + Len(vThing) + 4 'that is plus ": "
     
    lPosEnd = InStr(lStartThing, vJson, ",") - 1 ' that is - ",
    
    sString = Left(vJson, lPosEnd - 1)
    
    sThing = Right(sString, lPosEnd - lStartThing + 1)
    
    ciGetThingFromJson = sThing
Else
    ciGetThingFromJson = "error"
End If

ciGetThingFromJson_Exit:
    Exit Function

ciGetThingFromJson_Error:
    ciError Err.Description, Err.Number, "ciGetThingFromJson"
    Resume ciGetThingFromJson_Exit

End Function
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Function ciGetAttributeFromJson(vJson As String, _
                                       vThing As String) As String
'++
' Get the attribute for the sample lot from the json
' from this -> "CONC": {"stringData": "55.23"
' so pad = 19
'--
Dim sThing As String
Dim lStartThingName As Long
Dim lStartThing As Long
Dim lPosEnd As Long
Dim sString As String
Dim lPad As Long
Dim sTemp As String

On Error GoTo ciGetAttributeFromJson_Error

lPad = 17
sTemp = vThing & """" & ":{" & """" & "stringData" & """"

lStartThingName = InStr(vJson, sTemp)

If lStartThingName > 0 Then
    
    lStartThing = lStartThingName + Len(vThing) + lPad
     
    lPosEnd = InStr(lStartThing, vJson, "}") - 1 ' that is - ",
    
    sString = Left(vJson, lPosEnd - 1)
    
    sThing = Right(sString, lPosEnd - lStartThing)

    ciGetAttributeFromJson = sThing
Else
    ciGetAttributeFromJson = "error"
End If

ciGetAttributeFromJson_Exit:
    Exit Function

ciGetAttributeFromJson_Error:
    ciError Err.Description, Err.Number, "ciGetAttributeFromJson"
    Resume ciGetAttributeFromJson_Exit

End Function
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


Private Function pStripOutSessionID(vString As String) As String
'--
' Strip out the session id from the stuff looking for "JSESSIONID="
'++
Dim lPosStart As Long
Dim lPosEnd As Long
Dim lLen As Long
Dim sString As String
Dim sString2 As String

On Error GoTo pStripOutSessionID_Error

lPosStart = InStr(vString, "jsessionid")
lPosEnd = InStr(lPosStart, vString, "}")
lLen = lPosEnd - lPosStart

sString = Left(vString, lPosEnd - 2)

'take the  session id - len of  session id
sString2 = Right(sString, lLen - 14)

pStripOutSessionID = sString2

pStripOutSessionID_Exit:
    
    Exit Function

pStripOutSessionID_Error:
'    'Debug.Print "------------------------- Start -> "; Now()
'    'Debug.Print err.Description
'    'Debug.Print "------------------------- End   -> "; Now()
    Resume pStripOutSessionID_Exit
End Function




'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ciaFindCell(theWord As String, _
                   Optional theWorksheet As Worksheet, _
                   Optional vPrompt As Boolean) As String
'++
' Locate a cell in the sheet by contents, return the address
' Assuming we grab the used range and not searching the whole sheet, time etc
'--
Dim wkSht As Worksheet
Dim oRange As Range
Dim theCell As Range

On Error Resume Next
Set wkSht = theWorksheet

On Error GoTo ciaFindCell_Error

If wkSht Is Nothing Then
    Set wkSht = ActiveSheet
End If

Set oRange = wkSht.UsedRange

'Come back - no use for hidden sheets.
'but it breaks index update
oRange.Select

Set theCell = oRange.Find(theWord, , , lookat:=xlWhole)

ciaFindCell = theCell.Address

ciaFindCell_Exit:
    On Error Resume Next
    Set theCell = Nothing
    Set oRange = Nothing
    Set wkSht = Nothing
    Exit Function

ciaFindCell_Error:
    'Debug.Print ciaFindCell & ", " & ciaFindCell & "err-" & Err.Description
    If vPrompt Then
        Call ciError("This worksheet appears to be missing a column." & _
                Chr(13) & Chr(13) & _
                "Unable to find - " & theWord & _
                Chr(13) & Chr(13) & _
                "All valid data can still be sent back to Core", , , True)
    End If
    Resume ciaFindCell_Exit
End Function


Public Function ciaGetUsedRange(theStart As Range, _
                                theString As String, _
                                vLastCol As String) As Range
'++
' given a cell return the used range, cells that have data
'--
Dim sTemp As String
Dim sLastCol As String
Dim oTopRow As Range
Dim oUsedCol As Range

On Error GoTo ciaGetUsedRange_Error
  
If vLastCol <> "" Then
    sLastCol = vLastCol
ElseIf vLastCol <> "" Then
    vLastCol = ""
Else
   
End If

Set oTopRow = ciaGetUsedCol(theStart, "")

If theString = "" Then
    sTemp = mIndexColumn
Else
    sTemp = theString
End If

Set oUsedCol = ciaGetUsedRow(theStart, theString)

Set ciaGetUsedRange = Range(theStart, Cells(oUsedCol.Row, oTopRow.Column))

ciaGetUsedRange_Exit:
    On Error Resume Next
    Set oTopRow = Nothing
    Set oUsedCol = Nothing
    Exit Function

ciaGetUsedRange_Error:
    'Debug.Print Err.Description & " - ciaGetUsedRange_Error"
    Resume ciaGetUsedRange_Exit
End Function

Public Function ciaGetUsedCol(theStart As Range, _
                              theString As String) As Range
'++
' given a cell return the used range, cells that have data
' Come back here to rework
'--

On Error GoTo ciaGetUsedCol_Error
  Dim oRange As Range

If theString = "" Then
    Set oRange = Cells(1, theStart.End(xlToRight).Column)
    'Set oRange = Cells(theStart.SpecialCells(xlCellTypeLastCell).Row, 1)
Else
    Set oRange = Cells(theStart.Row, _
    Cells.Find(what:=theString, SearchOrder:=xlByColumns, _
    SearchDirection:=xlNext, LookIn:=xlValues).Column)
    
End If

Set ciaGetUsedCol = oRange

ciaGetUsedCol_Exit:
    On Error Resume Next
    Set oRange = Nothing
    
    Exit Function

ciaGetUsedCol_Error:
    'Debug.Print Err.Description
    Resume ciaGetUsedCol_Exit
End Function

Public Function ciaGetUsedRow(theStart As Range, _
                              theString As String) As Range
'++
' given a cell return the used range, cells that have data
'--
Dim oRange As Range
Dim theCell As Range

On Error GoTo ciaGetUsedRow_Error

If theString = "" Then
    theString = " "
End If

If theString = " " Then
    Set oRange = Cells(theStart.SpecialCells(xlCellTypeLastCell).Row, 1)
Else
    Set oRange = Cells(theStart.Row, _
        Cells.Find(what:=theString, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, LookIn:=xlValues).Column)
End If

'come back - why set oRange twice?
 
'Set oRange = Cells(theStart.End(xlDown).Row, theStart.Column)
    
'Set ciaGetUsedRow = oRange(oRange.Row, oRange.Column)
Set ciaGetUsedRow = oRange(oRange.Row, oRange.Column)

ciaGetUsedRow_Exit:
    On Error Resume Next
    Set oRange = Nothing
    Exit Function

ciaGetUsedRow_Error:
    'Stop
    'Debug.Print Err.Description & " - ciaGetUsedRow_Error"
    Resume ciaGetUsedRow_Exit
End Function


'"===
Public Function ciaGetUsedRowNum(theStart As Range, _
                                 theString As String) As Long
'++
' given a cell return the used range, cells that have data
'--
Dim oRange As Range
Dim theCell As Range

On Error GoTo ciaGetUsedRowNum_Error

If theString = "" Then
    theString = " "
End If

If theString = " " Then
    Set oRange = Cells(theStart.SpecialCells(xlCellTypeLastCell).Row, 1)
Else

Set oRange = Cells(theStart.Row, _
    Cells.Find(what:=theString, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, LookIn:=xlValues).Column)
End If

  
ciaGetUsedRowNum = oRange.Row

ciaGetUsedRowNum_Exit:
    On Error Resume Next
    Set oRange = Nothing
    Exit Function

ciaGetUsedRowNum_Error:
    'Stop
    'Debug.Print Err.Description
    Resume ciaGetUsedRowNum_Exit
End Function

'------
Public Sub ciaDoMoveColumnToEnd(theColumn As String)
'++
' Move the experiment ID columns from the downloaded poistions to away
' from the users area of interest.
'--

Dim lLastCol As Long
Dim sString As String
On Error GoTo ciaDoMoveColumnToEnd_Error

'get the last used column
lLastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
sString = ActiveSheet.Cells(1, lLastCol + 1).End(xlToLeft).Address
sString = ciaGetColumLetter(sString)

'move this column beyond it (to the right)
Application.CutCopyMode = False ' don't want an existing operation to interfere
Columns(sString).Insert XlDirection.xlToRight
Columns(sString).Value = Columns(theColumn).Value ' this would be one greater if to the right of F
Columns(theColumn).Delete

ciaDoMoveColumnToEnd_Exit:
    On Error Resume Next
   ' Set wkSht = Nothing
    Application.CutCopyMode = True
    Exit Sub
    
ciaDoMoveColumnToEnd_Error:
    'Debug.Print Err.Description
    Resume ciaDoMoveColumnToEnd_Exit

End Sub

Public Sub ciaDoMoveColumns()
'++
' Move the experiment ID columns from the downloaded poistions to away
' from the users area of interest.
'--
Dim wkSht As Worksheet
Dim sExSamBar As String 'EXPT_SAMPLE_BARCODE
Dim sSamTyp As String 'SAMPLE_TYPE_REF
Dim sSamBarRef As String 'SAMPLE_BARCODE_REF



On Error GoTo ciaDoMoveColumns_Error

Set wkSht = ActiveSheet

'find the addresses for the current experiment references
sExSamBar = ciaFindCell("EXPT_SAMPLE_BARCODE", wkSht)
sSamTyp = ciaFindCell("SAMPLE_TYPE_REF", wkSht)
sSamBarRef = ciaFindCell("SAMPLE_BARCODE_REF", wkSht)

'move them to the end
ciaDoMoveColumnToEnd (ciaGetColumLetter(sSamBarRef))
ciaDoMoveColumnToEnd (ciaGetColumLetter(sSamTyp))
ciaDoMoveColumnToEnd (ciaGetColumLetter(sExSamBar))

ciaDoMoveColumns_Exit:
    On Error Resume Next
    Set wkSht = Nothing
    Exit Sub
    
ciaDoMoveColumns_Error:
    'Debug.Print Err.Description
    Resume ciaDoMoveColumns_Exit

End Sub

Public Sub ciaDoReadOptions()
'++
' Show the dialog and collect option information
'--
On Error GoTo ciaDoReadOptions_Error

Application.StatusBar = "Reading options"

gBaseUrl = fmciaOptions.pgServerUrl

ciaDoReadOptions_Exit:
    On Error Resume Next
    Application.StatusBar = ""
    Exit Sub
    
ciaDoReadOptions_Error:
    'Stop 'Here
    ciError Err.Description, Err.Number, "ciaOptions"
    Resume ciaDoReadOptions_Exit
End Sub



Public Sub ciaOptions()
'++
' Show the dialog and collect option information
'--
On Error GoTo ciaOptions_Error
Application.StatusBar = "Show Options"

fmciaOptions.Show

ciaOptions_Exit:
    On Error Resume Next
    Application.StatusBar = ""
    Exit Sub
    
ciaOptions_Error:
    ciError Err.Description, Err.Number, "ciaOptions"
    Resume ciaOptions_Exit
End Sub

Public Sub DoClearRange(theRange As Range)
'++
' clear the contents of the cell
'--
Dim lCount As Long

On Error GoTo DoClearRange_Error

For lCount = 1 To theRange.Rows.Count
    theRange.Cells(lCount, 1).Value = ""
Next

DoClearRange_Exit:
    On Error Resume Next
    
    Exit Sub

DoClearRange_Error:
    'Debug.Print "DoClearRange - " & Err.Description
    Resume DoClearRange_Exit
    
End Sub


Public Sub DoClearRangeFormula(theRange As Range)
'++
' clear the contents of the cell
'--
Dim lCount As Long

On Error GoTo DoClearRangeFormula_Error

For lCount = 1 To theRange.Rows.Count
    If Left(theRange.Cells(lCount, 1).Formula, 1) = "=" Then
        theRange.Cells(lCount, 1).Formula = ""
    End If
Next

DoClearRangeFormula_Exit:
    On Error Resume Next
    
    Exit Sub

DoClearRangeFormula_Error:
    'Debug.Print "DoClearRangeFormula - " & Err.Description
    Resume DoClearRangeFormula_Exit
    
End Sub

'------
Public Sub ciaDoMoveColumnTo(theColumn As String, _
                    Optional theTargetColumn As String, _
                    Optional theLeft As Boolean)
'++
' Move the theColumn to the left of theTargetColumn
'--
Dim sTargetCol As String
Dim lLastCol As Long
Dim sString As String
Dim sTagetColLetter As String

On Error GoTo ciaDoMoveColumnTo_Error

sString = ciaFindCell(theTargetColumn)

sString = ciaGetColumLetter(sString)

'move this column beyond it (to the right)
Application.CutCopyMode = False ' don't want an existing operation to interfere

If theLeft Then
    Columns(sString).Insert XlDirection.xlToLeft
Else
    Columns(sString).Insert XlDirection.xlToRight
End If

sTagetColLetter = ciaFindCell(theColumn)

sTagetColLetter = ciaGetColumLetter(sTagetColLetter)

Columns(sString).Value = Columns(sTagetColLetter).Value ' this would be one greater if to the right of F

Columns(sTagetColLetter).Delete

ciaDoMoveColumnTo_Exit:
    On Error Resume Next
   ' Set wkSht = Nothing
    Application.CutCopyMode = True
    Exit Sub
    
ciaDoMoveColumnTo_Error:
    'Debug.Print Err.Description
    Resume ciaDoMoveColumnTo_Exit

End Sub
