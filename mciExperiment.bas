Attribute VB_Name = "mciExperiment"
Option Explicit
'++
'   Title:          mciExperiment
'
'   Purpose:        Providing functionality in support of the Experiment data
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       Given a defined set of columns pass this data to the server
'                   Assuming supported by mciReadMe, mciConstants,
'                   mciUtilities, mciInterface
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        10MAY2016
'
'   Modified by:    10MAY2016
'
'   Copyright:
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Version, Note, who
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 0.1 - starter version

Public Sub ciaDoSendExperimentData()
'++
' This will be replace when experiment template download write out
' coulmn names not attribute names.
' Send the updated data back to the platform this is taylored to
' pooling exp as the headders are attribute names
' For each row in the data post it back
' *beware if some attributes are null then they will be set to null
'--
Dim wkSht As Worksheet
Dim indexStart As String
Dim sampleStart As String
Dim headerEnd As String
Dim theRange As Range
Dim theSamples As Range
Dim theRow As Range
Dim lCol As Long
Dim lRow As Long
Dim sJSON As String
Dim oEntity As cciEntity
Dim sResponse As String
Dim sResponseTest As String
Dim oColumnMap As New cciEntity
Dim sTrans As String
Dim sTemp As String
        
On Error GoTo ciaDoSendExperimentData_Error

Set wkSht = ActiveSheet

gCancelEscape = False

Application.StatusBar = "Reading options.."

Application.StatusBar = "Working - there maybe a pause for network traffic..."

sampleStart = ciaFindCell(mSampleColumn)

headerEnd = ciaGetUsedCol(Cells(1, 1), "")

Set theSamples = wkSht.Range(sampleStart)

'Set theRange = ciaGetUsedRange(theSamples, "", headerEnd)
Set theRange = Selection 'ciaGetUsedRange(theSamples, "", headerEnd)

'FOR each row in the Range
'   get the JSON (jay-saawn)
'   post the data
'
'Le boom!

If gjSessionID = "" Then
    Call ciDoGetToken
    If gjSessionID = "" Then
        Call ciError("Error with Login", , , False)
        Exit Sub
    End If
End If


'set up Entitiy
Set oEntity = New cciEntity
Let oEntity.sdkCmd = "update-experiment-sample-data"
Let oEntity.AddEntityType = gExpSampleType
Let oEntity.AddSuperType = "EXPERIMENT SAMPLE"

lRow = 1
For Each theRow In theRange.Rows
    If gCancelEscape Then
        Exit For
    End If
    'clear the attribute listing
    oEntity.ClearAttributes
    
    lRow = lRow + 1
    
    Application.StatusBar = "Working - " & _
                     wkSht.Cells(lRow, 1).Text & _
                     " - Excel 'not responding' may show if the network is slow. " & _
                     "Please let it run a few sec more, thanks...."
    
    lCol = 1
    
    'Add the attribute values and the barcode
    Do While lCol <= theRange.Rows.Columns.Count And Not gCancelEscape
        'add the map transpose here
        sTemp = wkSht.Cells(1, lCol).Text
        
        Call oEntity.AddAttribute(sTemp, wkSht.Cells(lRow, lCol).Text)

        'get the barcode to pass to the entity
        If wkSht.Cells(1, lCol).Text = "EXPT_SAMPLE_BARCODE" Then
            Let oEntity.AddEntityBarcode = wkSht.Cells(lRow, lCol).Text
        End If
        lCol = lCol + 1
    Loop

    sJSON = oEntity.JSONExpSam
    
    sResponse = ciSendHTTPRequest(sJSON)
    
    'remove the following lines in Production code! or add a test
    sResponseTest = ciGetThingFromJson(sResponse, "success", True, 1)
    
    If sResponseTest = "error" Then
        If MsgBox("Unfortunately we're unable to update the record." & _
                     Chr(13) & _
                     Chr(13) & _
                    "Click OK to abort sending data." & _
                    Chr(13) & _
                    Chr(13) & _
                    "When contacting support, the server responded:" & _
                    Chr(13) & _
                    Chr(13) & _
                    sResponse, _
                                vbOKCancel, zAppTitle) = vbOK Then
            gCancelEscape = True
            
            Exit Sub
        End If
    End If
     
    With wkSht.Cells(lRow, 1).Interior
        If sResponseTest = "ru" Then
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599963377788629
            .PatternTintAndShade = 0
        Else
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End If
    End With
Next

Application.ScreenUpdating = False

wkSht.Cells(1, 1).Select


ciaDoSendExperimentData_Exit:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    Set theSamples = Nothing
    Set theRange = Nothing
    Set wkSht = Nothing
    Exit Sub
    
ciaDoSendExperimentData_Error:
    'Debug.Print "ciaDoSendExperimentData - " & Err.Description
    Resume ciaDoSendExperimentData_Exit
End Sub
