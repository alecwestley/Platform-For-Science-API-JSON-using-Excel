VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmciaOptions 
   Caption         =   "Options"
   ClientHeight    =   2037
   ClientLeft      =   21
   ClientTop       =   322
   ClientWidth     =   5859
   OleObjectBlob   =   "fmciaOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmciaOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Option Explicit
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'++
'   Title:          fmciaOptions
'
'   Purpose:        Gathering and updating Options for the template
'
'   Version:        0.1
'
'   Abstract:       Manage connection url, index columns and other bits
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        21/04/2016
'
'   Modified by:    Alec 16OCT2017
'
'   Copyright:      No sir.
'
'   Depends on:     See later.
'
'--
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Constants - used only in this Form
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const zFormTitle = "Options"

Private mValid As Boolean
Private mServerUrl As String    'Starting out with this to check against
Private mAccount As String      'tennent or account
Private mUserName As String     'to remember
Private mEntityName As String

Public Property Get pgServerUrl() As String
pgServerUrl = mServerUrl
End Property
Public Property Get pgAccount() As String
pgAccount = mAccount
End Property


Public Property Get pgUserName() As String
pgUserName = mUserName
End Property


Private Sub cmbTennant_Change()
'++
' Update the list form the choices
'--

mAccount = cmbTennant.Value
End Sub

Private Sub cmbUrl_Change()
'++
' On change the Url update the cmbAccount/Tennent values
'--
If cmbUrl.Value <> mServerUrl Then
    pUpdateAccount
    mServerUrl = cmbUrl.Value
End If
Me.cmbTennant.Value = mAccount
End Sub


Private Sub cmdAdvanced_Click()
'++
' Check the login and permissions (eventually)
' show the mapping worksheet
'--
Dim wkSht As Worksheet

On Error GoTo cmdAdvanced_Click_Error

Set wkSht = ThisWorkbook.Worksheets(zReadMeSheet)

wkSht.Activate


cmdAdvanced_Click_Exit:
    On Error Resume Next
    Set wkSht = Nothing
  Exit Sub
    
cmdAdvanced_Click_Error:
    Debug.Print "cmdAdvanced_Click - " & Err.Description
  Resume cmdAdvanced_Click_Exit
End Sub

Private Sub cmdCancel_Click()

mValid = False

pCleanUp

Unload Me

End Sub

Private Sub cmdOK_Click()

On Error GoTo cmdOK_Click_Error

mValid = True
'Add calls here
Call pWriteOptions

gCancelEscape = True


Call Form_Unload(0)
  
cmdOK_Click_Exit:
  Exit Sub
    
cmdOK_Click_Error:

  Resume cmdOK_Click_Exit
End Sub

Private Sub txtEntityType_Change()
mEntityName = txtEntityType.Value
End Sub

Private Sub txtUserName_Change()
mUserName = txtUserName.Value
End Sub

Private Sub UserForm_Initialize()
'++
' Setup default options if not set on startup
'--
On Error GoTo UserForm_Initialize_Error

Me.Caption = zAppTitle & " - Options"

'get the existing values
Call pReadOptions

pUpdateAccount

Me.cmbUrl.Value = mServerUrl

Me.txtUserName = mUserName

Me.txtEntityType = mEntityName

'replace with range and load please
Me.cmbUrl.AddItem zServer1
Me.cmbUrl.AddItem zServer2
Me.cmbUrl.AddItem zServer3

UserForm_Initialize_Exit:
    Exit Sub
    
UserForm_Initialize_Error:
    Resume UserForm_Initialize_Exit

End Sub


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



Private Sub Form_Unload(Cancel As Integer)

pCleanUp

Unload Me
 
End Sub

Private Sub pReadOptions()

Dim wkSht As Worksheet

On Error GoTo pReadOptions_Error

Set wkSht = ThisWorkbook.Worksheets(zReadMeSheet)

gBaseUrl = mServerUrl

mServerUrl = wkSht.Cells(1, 2).Value

mUserName = wkSht.Cells(4, 2).Value
mAccount = wkSht.Cells(6, 2).Value
mEntityName = wkSht.Cells(7, 2).Value

pReadOptions_Exit:
    On Error Resume Next
    Set wkSht = Nothing
    Exit Sub
    
pReadOptions_Error:
    'Debug.Print Err.Description
    Resume pReadOptions_Exit
End Sub

Private Sub pWriteOptions()

Dim wkSht As Worksheet

On Error GoTo pWriteOptions_Error

If mValid Then
    Set wkSht = ThisWorkbook.Worksheets(zReadMeSheet)
    wkSht.Cells(1, 2).Value = mServerUrl
    wkSht.Cells(4, 2).Value = mUserName
    wkSht.Cells(5, 2).Value = mPoolContType
    wkSht.Cells(6, 2).Value = mAccount
    wkSht.Cells(7, 2).Value = mEntityName

End If

ThisWorkbook.Save

pWriteOptions_Exit:
    On Error Resume Next
    Set wkSht = Nothing
    Exit Sub
    
pWriteOptions_Error:
   ' Stop
    Resume pWriteOptions_Exit
End Sub


Private Sub pUpdateAccount()
'++
' Update the accounts based on the list
'--
Dim lCount As Integer

On Error GoTo pUpdateAccount_Error


For lCount = 1 To Me.cmbTennant.ListCount
    Me.cmbTennant.RemoveItem lCount
Next lCount

Select Case Me.cmbUrl
    
    Case zServer1
        Me.cmbTennant.Value = zAccount1_1
        Me.cmbTennant.AddItem zAccount1_1
        Me.cmbTennant.AddItem zAccount1_2
   
    Case zServer2
        Me.cmbTennant.Value = zAccount2_1
        Me.cmbTennant.AddItem zAccount2_1
        
    Case zServer3
        Me.cmbTennant.AddItem zAccount3_1
        Me.cmbTennant.Clear
End Select

pUpdateAccount_Exit:
    On Error Resume Next
    Exit Sub
    
pUpdateAccount_Error:
    If Err.Number = -2147024809 Then
         Resume Next
    End If
    Resume pUpdateAccount_Exit
End Sub

