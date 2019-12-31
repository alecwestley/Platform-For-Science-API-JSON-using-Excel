VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cciAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'++
'   Title:          cciAttribute
'
'   Purpose:        Manage Enitty attributes makeing JSON interaction simpler.
'
'   App Version:    see mciConstants
'
'   Module Version: 0.2
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       The JSON translation in VBA is not great. internet based
'                   code provided in upto version 1.0 is sub-optima, Opensource
'                   is often fast n dirty especially if it's not well commented!
'
'                   This class supports an entity class(cciEntity) providing the
'                   the attribute value pairs SJON as it's the only really hard
'                   section of the JSON to manage.
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        19Mar2016
'
'   Modified by:    09Nov2016
'
'   Copyright:      No
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Added fix for API and name spaces se Public Property Get Name() As String
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private cName As String        'The attribute name.
Private cNameAlt As String     'Used to manage different attribute to Excel name
                                ' as of today (10May2016) the exp template gives
                                ' configured names not column names for user
Private cValue As String       'The attribute value, all handled as strings
Private cMandatory As Boolean  '

Private Const theStart = "{"
Private Const theQuote = """"
Private Const theSep = ":"
Private Const theEnd = "}"

Public Property Get JSON() As String
'++
' Return the value of the JSON
'--

JSON = pGetJSON

End Property

Public Property Get JSON_Simple() As String
'++
' Return the value of the JSON
'--

JSON_Simple = pGetJSON_Simple

End Property
Public Property Get JSON_expSam() As String
'++
' Return the value of the JSON
'--

JSON_expSam = pGetJSON_expSam

End Property

Public Property Get Name() As String
'++
' Return the value of the JSON
'--
'Enable this line if API problems -> JSON no like spaces
Name = UCase(Replace(cName, " ", "_", 1))
Name = cName
End Property

Public Property Get NameAlt() As String
'++
' Return the value of the JSON
'--

NameAlt = cNameAlt

End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Sub AddAttribute(vName As String, _
                        vValue As String, _
                        vMandatory As Boolean, _
               Optional vNameAlt As String)
'++
' Return the value of the JSON
'--
cName = vName
cValue = vValue
cMandatory = vMandatory
cNameAlt = vNameAlt
End Sub

Public Sub Class_Terminate()
'++
' Tidy up
'--
On Error Resume Next


End Sub

Private Function pGetJSON() As String
'++
' Generate the JSON
' ****come back here**** Add the lower level value pairs.
' We're assuming here that we're not having multiple choices per cell.
' This may well have to change but v1.0 is not specified.
'--
'                "CIAMANGANESE": {
'                    "mandatory": 0,
'                    "stringData": ""
'                },
Dim theString As String

On Error GoTo pGetJSON_Error

If cMandatory Then
    theString = """" & cName & """" & theSep & theStart & """mandatory"":1,""stringData"": """ & cValue & """" & theEnd
Else
    theString = """" & cName & """" & theSep & theStart & """mandatory"":0,""stringData"": """ & cValue & """" & theEnd
End If

pGetJSON = theString

pGetJSON_Exit:
    Exit Function
    
pGetJSON_Error:
    'Debug.Print Err.Description
    Resume pGetJSON_Exit

End Function

Private Function pGetJSON_Simple() As String
'++
' Generate the JSON
' ****come back here**** Add the lower level value pairs.
' We're assuming here that we're not having multiple choices per cell.
' This may well have to change but v1.0 is not specified.
'--
'                "<String>":"<value>",
Dim theString As String

On Error GoTo pGetJSON_Simple_Error

theString = """" & cName & """" & theSep & """" & cValue & """"

pGetJSON_Simple = theString

pGetJSON_Simple_Exit:
    Exit Function
    
pGetJSON_Simple_Error:
    'Debug.Print Err.Description
    Resume pGetJSON_Simple_Exit

End Function


Private Function pGetJSON_expSam() As String
'++
' Generate the JSON
' ****come back here**** Add the lower level value pairs.
' We're assuming here that we're not having multiple choices per cell.
' This may well have to change but v1.0 is not specified.
'--
'                "CIAMANGANESE": {
'                    "mandatory": 0,
'                    "stringData": ""
'                },
Dim theString As String

On Error GoTo pGetJSON_expSam_Error
    
theString = """" & UCase(cName) & """" & theSep & theStart & """stringData"": """ & cValue & """" & theEnd

pGetJSON_expSam = theString

pGetJSON_expSam_Exit:
    Exit Function
    
pGetJSON_expSam_Error:
    'Debug.Print Err.Description
    Resume pGetJSON_expSam_Exit

End Function
