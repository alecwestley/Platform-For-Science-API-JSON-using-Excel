Attribute VB_Name = "mciConstants"
'++
'   Title:          mciConstants
'
'   Purpose:        Hold application wide constants
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       Please read mciaReadMe first.
'                   Constans used throughout the application, should be prefixed
'                   by a zed or zee (z) kept here for quick access for debug.
'                   All test values can be changed here as the code will look to
'                   here.
'                   *******  No variables/values are in the code  *******
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Modified:       15OCT2017
'
'   Created:        06Nov2014
'
'   Copyright:
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Version, who, When, Note
' 1.1.0 Alec 10MAY2016 Much revision of pooling + refactored code for deadwood
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Connection
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Option Explicit
Public Const zAppTitle = "Platform for Science - Desktop / Excel"
Public Const zAppVersion = "2.1.0"
Public gBaseUrl As String       'Passed around by varions actions - user update
Public gjSessionID As String    ' needed for server chatter
Public gExpSampleType As String 'Can be removed (with depend) note!
Public Const zReadMeSheet = "wkshtCiaReadMe"
Public gCancelEscape As Boolean

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Experiment Template sheet
' Standard columns to hide (or move)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const zExpShtCol1 = "EXPT_SAMPLE_BARCODE" ' maybe later
Public Const zExpShtCol2 = "SAMPLE_TYPE_REF"
Public Const zExpShtCol3 = "SAMPLE_BARCODE_REF"
Public Const zExpShtCol4 = "SAMPLE_NAME_REF"
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Experiment Template sheet - Ends
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' POOLING - Starts
' These are 'hard coded' solution based constants that will not change
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const mPoolColumn = "ciaPoolName"
Public Const mIndexColumn = "Index Barcode"
Public Const mSampleColumn = "EXPT_SAMPLE_BARCODE" '"SAMPLE_NAME_REF"
Public Const mLastHeaderColumn = "Comments" '"ci_bulkTransfer_targetContainer"
Public Const mTheI5Index = "ci_index"
Public Const mTheI7Index = "ci_index2"
Public Const mtheVolToTake = "ciaMassToTake"
Public Const mtheCalcColumn = "ci_bulkTransfer_amount"
Public Const mtheAmount = "ci_Cell_amount"
Public Const mtheConc = "ci_CellContents_concentrationNanoMolar"
Public Const mExpSampleType = "QPCR QS STEP SAMPLE"
Public Const mMapESam = "ES_INDEX"
Public Const mMapCont = "DESTINATION_CONTAINER"
Public Const mMapDest = "DESTINATION_CELL"
Public Const mPoolContainer = "ci_bulkTransfer_targetContainer"
Public Const mPoolContType = "AMPLICON POOLING TUBE"
Public Const zServer1 = "https://experience52.platformforscience.com"
Public Const zServer2 = "https://<add me 2>"
Public Const zServer3 = "https://<add me 3>"
Public Const zAccount1_1 = "CHANGE ME 0"
Public Const zAccount1_2 = "CHANGE ME 1"
Public Const zAccount2_1 = "CHANGE ME 2"
Public Const zAccount3_1 = "CHANGE ME 3"
Public Const zPoolCont1 = "TUBE"
Public Const zPoolCont2 = "LOBIND TUBE"
Public Const zPoolCont3 = "POOL TUBE"
Public Const zCiDiluentVol = ""
Public Const zAttMapCol = 3 ' Start the headder translation here (col 2 & 4)
Public Const zNameNew = "Name New" 'New name to replace the old name
Public Const zNameOld = "Name Old"


Public gTheI5Index As String
Public gTheI7Index As String
Public gPoolContType As String
Public gExpSample As cciEntity ' remapping col headders (temp)
Public gEntityType As String


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Pooling - Ends
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

