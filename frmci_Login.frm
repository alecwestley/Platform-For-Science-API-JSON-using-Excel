VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmci_Login 
   Caption         =   "Core Platform login"
   ClientHeight    =   1505
   ClientLeft      =   49
   ClientTop       =   322
   ClientWidth     =   5061
   OleObjectBlob   =   "frmci_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmci_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Option Explicit
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'++
'   Title:          frmLogin
'
'   Purpose:        Connecting to data Sources
'
'   Version:        0.5
'
'   Abstract:       Connects to a platform and returns the connection and
'                   handles any errors.
'                   The order of code is: Const>Work>Properties>Form>Utilities
'                   You, dear reader, will want the Work section.
'                   This form is intended to work as a selfcontained unit so
'                   there maybe some duplication of code.
'                   Also see mci_aReadMe for more information
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        14/03/2005
'
'   Modified by:    09OCT2015 Sanity Check and tidy, no function change
'                   07Nov2014 Added url and login see note on mBaseUrl
'                   06Nov2014 Alec simplified for simple login;
'                   17JUL2008 Modified Alec
'
'   Copyright:      No sir.
'
'   Depends on:     See later.
'
'--
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Constants - used only in this Form
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'These can be changed to lang or other msgs
Private Const zDebug = "Debug const or msg Please change"
Private Const zErrEmptyControl = "Please add the correct info to "
Private Const zErrConnection = " - Connection error"
Private Const zfUserName = "Change Me"
Private Const zfPassword = "Change Me"
Private Const zfAppTitle = "You forgot to set the App Title on initialze"
Private Const zfAppVersion = "You forgot to set the version on initialze"
Private Const zReadMeSheet = "wkshtCiaReadMe"

Private mAppTitle As String
Private mAppVersion As String
Private mIsValid As Boolean    'used to show/hide/control this form
Private mCancel As Boolean      'True = Cancel button pressed
Private mUserName As String
Private mPassword As String
Private mAccount As String
Private mgSessionId As String
Private mBaseUrl As String
Private mErrorMsg As String
Private mExpSampleType As String



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Work - Functionality
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Sub pDoGetConnection()
'++
' Setup the connection and get back the securty token to use going forward
'--
Dim sJsonStr As String
Dim sHeadText As String
Dim oHttp As MSXML2.XMLHTTP60

On Error GoTo pDoGetConnection_Error

'Set cBro = New cBrowser
Set oHttp = New MSXML2.XMLHTTP60

'get the josn
sJsonStr = pGetJSON
 
'Url encode it
sJsonStr = "json=" & sJsonStr

oHttp.Open "POST", gBaseUrl & "/sdklogin", False
oHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

oHttp.sEnd sJsonStr

sHeadText = pStripOutSessionID(oHttp.ResponseText)

'add here a check on session id?

If oHttp.Status = 200 And sHeadText <> "" Then
    mIsValid = True
    mgSessionId = sHeadText
Else
  '  mErrorMsg = oHttp.ResponseText
   ' MsgBox mErrorMsg
End If

    
pDoGetConnection_Exit:
    On Error Resume Next
    Set oHttp = Nothing
    Exit Sub
    
pDoGetConnection_Error:
    'Debug.Print Err.Description & ", " & Err.Number
    Call pError("Connection - " & Err.Description, Err.Number)
    Resume pDoGetConnection_Exit
    
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public properties - Let
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Let plDatabase(vBaseUrl As String)
mBaseUrl = vBaseUrl
End Property

Public Property Let plUserName(sUserName As String)
mUserName = sUserName
End Property

Public Property Let plPassWord(sPassWord As String)
mPassword = sPassWord
End Property
Public Property Let plAccount(sAccount As String)
mAccount = sAccount
End Property

Public Property Let plBaseUrl(sBaseUrl As String)
mBaseUrl = sBaseUrl
End Property

Public Property Let plAppTitle(sAppTitle As String)
mAppTitle = sAppTitle
End Property
Public Property Let plAppVersion(sAppVersion As String)
mAppVersion = sAppVersion
End Property

Public Property Let plExpSampleType(sExpSampleType As String)
mExpSampleType = sExpSampleType
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public properties - Get
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Property Get pgToken() As String
pgToken = mgSessionId
End Property

Public Property Get pgIsValid() As Boolean
pgIsValid = mIsValid
End Property

Public Property Get pgExpSampleType() As String
plExpSampleType = mExpSampleType
End Property

Private Sub txtPassword_Change()
If txtPassword = "" Then
    mIsValid = False
End If
End Sub

Private Sub txtUserName_Change()
mUserName = txtUserName.Value
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Form level
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub UserForm_Initialize()
'++
' Setup default options if not set on startup
'--
On Error GoTo UserForm_Initialize_Error

'read the saved user name
Call pReadOptions

If mUserName = "" Then
    mUserName = "Change Me"
End If
If mPassword = "" Then
    mPassword = "password"
End If
If mAppTitle = "" Then
    mAppTitle = zfAppTitle
End If

If mAppVersion = "" Then
    mAppVersion = zfAppVersion
End If

If mAccount = "" Then
    mAccount = ""
End If

If gBaseUrl = "" Then
    'show the options form to gather the data
    'this writes the data to the sheet
    If mCancel = False Then
        fmciaOptions.Show
        pReadOptions
        mBaseUrl = ""
    End If
End If
UserForm_Initialize_Exit:
    Exit Sub
    
UserForm_Initialize_Error:
    Resume UserForm_Initialize_Exit

End Sub

Private Sub UserForm_Activate()
'++
' Setup for showing the form
'--
'
Call pSetUp
Call UserForm_Initialize
Application.StatusBar = "Logged out"
End Sub

Private Sub Form_Unload(Cancel As Integer)

pCleanUp

Unload Me
 
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Form controls
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdOK_Click()

On Error GoTo cmdOK_Click_Error
'used to escape the options form
mCancel = False

mUserName = Me.txtUserName.Text
mPassword = Me.txtPassword.Text

'Add calls here
pDoGetConnection


If mIsValid Then
    Call pWriteOptions
    Application.StatusBar = "Connected to " & gBaseUrl & " - " & mAccount & _
                            " as " & mUserName
  Me.Hide
  
Else
    Me.lblVersion.Caption = "Error connecting please try again or cancel"
End If


cmdOK_Click_Exit:
  Exit Sub
    
cmdOK_Click_Error:
  Call pError("OK - " & Err.Description, Err.Number)
  Resume cmdOK_Click_Exit
  
End Sub

Private Sub cmdCancel_Click()

mIsValid = False

'used to escape the options for show
mCancel = True

pCleanUp

Unload Me

End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private utility tools
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Sub pCleanUp()

Dim lcnt As Long

'shocking miss use of power
On Error Resume Next

Call pWriteOptions

For lcnt = 0 To Me.Controls.Count
  If Me.Controls(lcnt).Name Like "txt*" Then
   Me.Controls(lcnt).Text = ""
  End If
Next

cmdCancel.SetFocus

End Sub

Private Function pCheckControls() As Boolean
Dim lcnt As Long

'shocking miss use of power
On Error GoTo pCheckControls_Error
pCheckControls = False

For lcnt = 1 To Me.Controls.Count - 1
  If Me.Controls(lcnt).Name Like "txt*" Then
   If Me.Controls(lcnt).Text = "" Then
     MsgBox zErrEmptyControl & Me.Controls(lcnt).Name, vbCritical, mAppTitle
     Exit Function
   End If
  End If
Next

pCheckControls = True

pCheckControls_Exit:
  Exit Function

pCheckControls_Error:
  Call pError("check Controls - " & Err.Description, Err.Number)
  Resume pCheckControls_Exit

End Function

Private Sub pSetupLogin()
'++
' Setup the text on the login tab
'--

On Error GoTo pSetupLogin_Error

Me.txtUserName = mUserName
Me.txtPassword = mPassword

pSetupLogin_Exit:
   Exit Sub

pSetupLogin_Error:
   pError "Login - " & Err.Description, Err.Number
   Resume pSetupLogin_Exit
End Sub

Private Sub pSetUp()

On Error GoTo pSetUp_Error
pReadOptions
Me.Caption = mAppTitle
Me.lblVersion.Caption = mAppVersion
mBaseUrl = mBaseUrl

Me.txtUserName = mUserName
Me.txtPassword = mPassword

'
pSetUp_Exit:
    Exit Sub

pSetUp_Error:
    pError "Form Setup - " & Err.Description, Err.Number
    Resume pSetUp_Exit
    
End Sub

Private Sub pError(sText As String, Optional lError As Long)
'++
' Handler for err text, simples
'--
Dim sPad As String

If lError <> 0 Then
    sPad = " " & Err.Number
End If

On Error Resume Next
    MsgBox Trim(sText) & sPad, vbCritical, mAppTitle
End Sub

Private Function pGetJSON() As String
'++
' Return the json as a string as it's not very VBA friendly with all the "
'--
Dim sPre As String
Dim sSep As String
Dim sAcc As String
Dim sSuf As String
Dim sString As String

On Error GoTo pGetJSON_Error

'create the json string for the login use the with to give the correct levels

sPre = "{""request"":{""data"":{""lims_userName"":"""
sSep = """,""lims_password"":"""
sAcc = """,""accountRef"":{""entityId"":"""",""barcode"":"""",""name"":"""
sSuf = """}},""typeParam"":""*"",""sdkCmd"":""sdk-login""}}"

sString = sPre & mUserName & sSep & mPassword & sAcc & mAccount & sSuf

pGetJSON = sString

pGetJSON_Exit:
    On Error Resume Next
    Exit Function
    
pGetJSON_Error:
    Call pError("Get Json - " & Err.des, Err.num)
    Resume pGetJSON_Exit
End Function


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

pStripOutSessionID = pStripOutEmployeeID(sString2)

pStripOutSessionID_Exit:
    
    Exit Function

pStripOutSessionID_Error:
    'Debug.Print Err.Description
    Resume pStripOutSessionID_Exit
End Function

Private Function pStripOutEmployeeID(vString As String) As String
'--
' Strip out the session id from the stuff looking for "employeeId="
'++
Dim lPosStart As Long
Dim lPosEnd As Long
Dim lLen As Long
Dim sString As String
Dim sString2 As String

On Error GoTo pStripOutEmployeeID_Error

lPosStart = InStr(vString, ",")


sString = Left(vString, lPosStart - 2) '-2 for the ",

pStripOutEmployeeID = sString

pStripOutEmployeeID_Exit:
    
    Exit Function

pStripOutEmployeeID_Error:
    'Debug.Print Err.Description
    Resume pStripOutEmployeeID_Exit
End Function


Public Sub DoLogOut()
'++
' Call the SKD Logout
'--

Dim sJSON As String
Dim sString As String
Dim sJsonLogout As String

On Error GoTo DoLogOut_Error

'Get the json
sJSON = pGetJsonLogOut

sString = mBaseUrl & "/sdklogin" & ";jessionid=" & mgSessionId & "json=" & sJSON

DoLogOut_Exit:
    On Error Resume Next
    Exit Sub

DoLogOut_Error:
    Resume DoLogOut_Exit
    
End Sub
Private Function pGetJsonLogOut() As String
'++
' Return the json as a string as it's not very VBA friendly with all the "
'--

Dim oEntity As cciEntity


On Error GoTo pGetJsonLogOut_Error

Set oEntity = New cciEntity

Let oEntity.sdkCmd = "sdk-logout"

pGetJsonLogOut = oEntity.JSON


pGetJsonLogOut_Exit:
    On Error Resume Next
    'Set jObj = Nothing
    Set oEntity = Nothing
    Exit Function
    
pGetJsonLogOut_Error:
   ' Stop
    Call ciError("Get Json - " & Err.des, Err.num)
    Resume pGetJsonLogOut_Exit
End Function

Private Sub pReadOptions()

Dim wkSht As Worksheet

On Error GoTo pReadOptions_Error

Set wkSht = ThisWorkbook.Worksheets(zReadMeSheet)

mUserName = wkSht.Cells(4, 2).Value
gBaseUrl = wkSht.Cells(1, 2).Value
mAccount = wkSht.Cells(6, 2).Value
gEntityType = wkSht.Cells(7, 2).Value

pReadOptions_Exit:
    On Error Resume Next
    Set wkSht = Nothing
    Exit Sub
    
pReadOptions_Error:
   ' Stop
    Resume pReadOptions_Exit
End Sub

Private Sub pWriteOptions()

Dim wkSht As Worksheet

On Error GoTo pWriteOptions_Error

Set wkSht = ThisWorkbook.Worksheets(zReadMeSheet)

'only update the username if added
If mUserName <> "" Then
    wkSht.Cells(4, 2).Value = mUserName
    wkSht.Cells(1, 2).Value = gBaseUrl
End If

pWriteOptions_Exit:
    On Error Resume Next
    Set wkSht = Nothing
    Exit Sub
    
pWriteOptions_Error:
    Resume pWriteOptions_Exit
End Sub
