VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cciEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'++
'   Title:          cciEntity
'
'   Purpose:        Manage mapping Platform entities makeing JSON interaction
'                   simpler.
'
'   App Version:    see mciConstants
'
'   Module Version: 0.1
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       The JSON translation in VBA is not great. internet based
'                   code provided in upto version 1.0 is not great, but is open-
'                   source.
'                   This class will be supported by an attribute class to use in
'                   collection to map out the JSON.
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        18Mar2016
'
'   Modified by:
'
'   Copyright:      No
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Hint - Watch out for COntrollled vocab on Experiments - you get success
'        returned when an error will wipe the data...
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


Private cSuperType As String
Private cEntityType As String
Private cSdkCmd As String
Private cHeadder As String
Private cFooter As String
Private cEntityName As String
Private cEntityID As String
Private cEntityBarcode As String
Private cFileAtt As String
Private cLocationId As String
Private cAttributes As Collection
Private cAssociations As Collection

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Return data to the caller
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Property Get SuperType() As String
'++
' Return the value held as the SuperType
'--
SuperType = cSuperType
End Property

Public Property Get EntityType() As String
'++
' Return the value held as the EntityType
'--
EntityType = cEntityType
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Add values - using Let, we're not using objects externally so no set
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Property Let AddSuperType(vSuperType As String)
'++
' Setup the SuperType to use
'--
cSuperType = vSuperType
End Property

Public Property Let AddEntityType(vEntityType As String)
'++
' Setup the EntityType to use
'--
cEntityType = vEntityType
End Property

Public Property Let sdkCmd(vCommand As String)
'++
' Setup the SDK command type, useful commands are:
' get, set, update,create, update-experiment-sample-data
'--
cSdkCmd = vCommand
End Property

Public Property Let AddEntityName(vEntityName As String)
'++
' Add the entity name of the value pair, not compulsary
'--
cEntityName = vEntityName
End Property

Public Property Let AddEntityID(vEntityID As String)
'++
' Add the entity ID of the value pair, not compulsary
'--
cEntityID = vEntityID
End Property

Public Property Let AddEntityBarcode(vEntityBarcode As String)
'++
' Add the entity Barcode value of the value pair, best to use this
'--
cEntityBarcode = vEntityBarcode
End Property

Public Property Let AddlocationId(vlocationId As String)
'++
' Add the Location ID value of the value pair, not compulsary
'--
cLocationId = vlocationId
End Property

Public Property Let FileAtt(vFullName As String)
cFileAtt = vFullName
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Associations - not complusary
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get AssociationList() As String
'++
' Return as string of the list of associations
' Maybe useful for debug, what associations have you got?
'--
AssociationList = pGetAssociations
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Attruibutes - generally needed ;-)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


Public Property Get AttributeList() As String
'++
' Return the list of attributes we hold, could be nice for debugging code
'--
AttributeList = pGetAttributes

End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Worker subs and functions
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Public subs and functions
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub AddAttribute(vName As String, _
                        vValue As String, _
               Optional vMandatory As Boolean, _
               Optional vAltName As String)
'++
' Add an Attribute
'--
Call pDoAddAttribute(vName, _
                     vValue, _
                     vMandatory, _
                     vAltName)

End Sub

Public Sub AddAssociation(vAssociationName As String)
'++
' Add an association to the list of associations, not compulsary
'--
Call pDoAddAssociation(vAssociationName)
End Sub

Public Function JSON() As String
JSON = pDoMakeJSON(cSuperType, cEntityType, cEntityID, cSdkCmd)
End Function
Public Function JSONExpSam() As String
JSONExpSam = pDoMakeJSONExpSam(cSuperType, cEntityType, cEntityBarcode, cSdkCmd, True)
End Function

Public Function JSONCreate() As String
JSONCreate = pDoMakeJSONCreate(cSuperType, cEntityType, cEntityID, cSdkCmd)
End Function

Public Function AttributeFromName(theName As String) As String
AttributeFromName = pGetNiceName(theName)
End Function

Public Function NameFromAttribute(theName As String) As String
NameFromAttribute = pGetAttributeName(theName)
End Function

Public Function ClearAttributes()
'++
' Clear the current attributes for reuse
'--
Call ClassTermHelper(cAttributes)
End Function

Public Sub Class_Terminate()
'++
' Tidy up
'--
On Error Resume Next
Call ClassTermHelper(cAttributes)
Call ClassTermHelper(cAssociations)

End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Private subs and functions
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClassTermHelper(vCollection As Collection)
'++
' Clear the objects
'--
Dim lCount As Long
Dim lColCount As Long

On Error GoTo ClassTermHelper_Error

lColCount = vCollection.Count

If Not vCollection Is Nothing Then
    lCount = 1
    
    Do While lCount <= lColCount
        vCollection.Remove (1)
        lCount = lCount + 1
        
    Loop
End If

ClassTermHelper_Exit:
    Exit Sub
    
ClassTermHelper_Error:

    Resume ClassTermHelper_Exit

End Sub


Private Sub pDoAddAssociation(vName As String)
Dim cAssociation As cciAssociation
Dim sJsoName As String
Dim vValue As String 'Come back here
Dim vMandatory As Boolean
Dim vAltName As String


On Error GoTo pDoAddAssociation_Error

Set cAssociation = New cciAssociation

sJsoName = UCase(Replace(vName, " ", "_", 1))

With cAssociation
    .AddAssociation vName, vValue, vMandatory, vAltName
End With


pDoAddAssociation_Exit:
    On Error Resume Next
    Set cAssociation = Nothing
    Exit Sub
    
pDoAddAssociation_Error:
    Resume pDoAddAssociation_Exit
    

End Sub


Private Function pGetAssociations() As String

End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Attirbutes
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Sub pDoAddAttribute(vName As String, _
                            vValue As String, _
                   Optional vMandatory As Boolean, _
                   Optional vAltName As String)
                   
Dim cAttribute As cciAttribute
Dim sJsoName As String

'                "CIAMANGANESE": {
'                    "mandatory": 0,
'                    "stringData": ""
'                },

On Error GoTo pDoAddAttribute_Error

Set cAttribute = New cciAttribute

sJsoName = UCase(Replace(vName, " ", "_", 1))

With cAttribute
    .AddAttribute vName, vValue, vMandatory, vAltName
End With

'if we don't already have a colleciton of attributes make one
If cAttributes Is Nothing Then
    Set cAttributes = New Collection
End If

cAttributes.Add cAttribute, vName

pDoAddAttribute_Exit:
    On Error Resume Next
    'remove me
    'Set cAttribute = Nothing
    Exit Sub
pDoAddAttribute_Error:
    'Debug.Print Err.Description
    Resume pDoAddAttribute_Exit
End Sub


Private Function pGetAttributes() As String

End Function

Private Sub pDoClearObjects()
'Call this to clean out the attributes collection

'For each cciAttribute in the collection
'   set it = nothing

End Sub
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' JSON
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function pDoMakeJSON(vSuperType As String, _
                             vEntityType As String, _
                             vEntityID As String, _
                             vSdkCmd As String)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Build the JSON
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Dim sJsonAtttributes As String
Dim sJsonPrefix As String
Dim sJsonSuffix As String

Dim lCount As Long

On Error GoTo pDoMakeJSON_Error
            
sJsonPrefix = "{""request"": {""sdkCmd"": """ & vSdkCmd & """," & _
              """data"": {""associations"": {},""superTypeName"": """ & vSuperType & """," & _
              """locationId"": 0,""values"": {"
              
              
              
sJsonSuffix = "},""name"": """",""entityTypeName"": """ & vEntityType & """," & _
             """projectIds"": """", ""barcode"": """",""sampleRef"": {""name"": """"," & _
             " ""entityId"": """",""barcode"": """"}},""responseOptions"":" & _
             " [""CONTEXT_GET"",""MESSAGE_LEVEL_WARN""],""typeParam"": """ & vEntityType & """,""logicOptions"": [""EXECUTE_TRIGGERS""]}}"


lCount = 1
sJsonAtttributes = ""

If Not cAttributes Is Nothing Then

    Do While lCount <= cAttributes.Count
        sJsonAtttributes = sJsonAtttributes + cAttributes(lCount).JSON & ","
        lCount = lCount + 1
    Loop
    
    
    sJsonAtttributes = Left(sJsonAtttributes, Len(sJsonAtttributes) - 1)
End If
pDoMakeJSON = sJsonPrefix & sJsonAtttributes & sJsonSuffix

pDoMakeJSON_Exit:
    On Error Resume Next
    
    Exit Function

pDoMakeJSON_Error:
    'Debug.Print Err.Description
    'Stop
    Resume pDoMakeJSON_Exit
End Function
Private Function pDoMakeJSONExpSam(vSuperType As String, _
                             vEntityType As String, _
                             vEntityID As String, _
                             vSdkCmd As String, _
                    Optional vExTrigger As Boolean)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Build the JSON
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Dim sJsonAtttributes As String
Dim sJsonPrefix As String
Dim sJsonSuffix As String
Dim sPostBarcode As String
Dim sExTrigger As String

Dim lCount As Long

On Error GoTo pDoMakeJSONExpSam_Error
sJsonPrefix = "{""request"":{""logicOptions"":[""EXECUTE_TRIGGERS"" ],""data"":{""entityTypeName"":" & gEntityType & ",""projectIds"":"""",""experimentData"":{"
          
'sJsonPrefix = "{""request"":{""logicOptions"":["""" ],""data"":{""entityTypeName"":""POOLING RUN SAMPLE"",""projectIds"":"""",""experimentData"":{"
                             
sJsonSuffix = "},""name"":"""",""sampleRef"":{""name"":"""",""entityID"":"""",""barcode"":""""},""associations"":{""{"":""}""},""locationID"":0,""superTypeName"":""EXPERIMENT SAMPLE"",""barcode"":""" & vEntityID & """},""typeParam"":" & gEntityType & ",""sdkCmd"":" & vSdkCmd & ",""responseOptions"":[""CONTEXT_GET"",""MESSAGE_LEVEL_WARN""]}}"

lCount = 1

sJsonAtttributes = ""

Do While lCount <= cAttributes.Count
    sJsonAtttributes = sJsonAtttributes + cAttributes(lCount).JSON_expSam & ","
    lCount = lCount + 1
Loop

sJsonAtttributes = Left(sJsonAtttributes, Len(sJsonAtttributes) - 1)

pDoMakeJSONExpSam = sJsonPrefix & sJsonAtttributes & sJsonSuffix '& sPostBarcode

pDoMakeJSONExpSam_Exit:
    On Error Resume Next
    
    Exit Function

pDoMakeJSONExpSam_Error:
    'Debug.Print Err.Description
    Resume pDoMakeJSONExpSam_Exit
End Function
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' JSON
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function pDoMakeJSONCreate(vSuperType As String, _
                             vEntityType As String, _
                             vEntityID As String, _
                             vSdkCmd As String)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Build the JSON
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Dim sJsonAtttributes As String
Dim sJsonPrefix As String

On Error GoTo pDoMakeJSONCreate_Error
            
sJsonPrefix = "{""request"": {""sdkCmd"": """ & vSdkCmd & """," & _
              """data"": {""associations"": {},""superTypeName"": """ & vSuperType & """," & _
              """locationId"": 0,""values"": """",""name"": """",""entityTypeName"": """ & vEntityType & """," & _
             """projectIds"": """", ""barcode"": """"},""responseOptions"":" & _
             " [""CONTEXT_GET"",""MESSAGE_LEVEL_WARN""],""typeParam"": """ & vEntityType & """,""logicOptions"": [""EXECUTE_TRIGGERS""]}}"


pDoMakeJSONCreate = sJsonPrefix

pDoMakeJSONCreate_Exit:
    On Error Resume Next
    
    Exit Function

pDoMakeJSONCreate_Error:
    'Debug.Print Err.Description
    Resume pDoMakeJSONCreate_Exit
End Function

Private Function pGetNiceName(theName As String) As String
'++
' Given the "nice" name return the attribute name
'--
Dim lCount As Long
Dim sName As String

On Error GoTo pGetNiceName_Error

sName = theName

lCount = 1

Do While lCount <= cAttributes.Count
   'Debug.Print "The Name = " & theName & "|Alt = " & cAttributes.Item(lCount).NameAlt & "|Att = " & cAttributes.Item(lCount).Name
    If cAttributes.Item(lCount).Name = theName Then
        sName = cAttributes.Item(lCount).NameAlt
        Exit Do
    End If
    lCount = lCount + 1
Loop

pGetNiceName = sName

pGetNiceName_Exit:
    Exit Function
    
pGetNiceName_Error:
    'Debug.Print "pGetNiceName - " & Err.Description
    Resume pGetNiceName_Exit

End Function


Private Function pGetAttributeName(theName As String) As String
'++
' Given the "nice" name return the attribute name
'--
Dim lCount As Long
Dim sName As String

On Error GoTo pGetAttributeName_Error

sName = theName

lCount = 1

Do While lCount <= cAttributes.Count

   'Debug.Print "|" & theName & "|" & Chr$(13) & "|" & _
   cAttributes.Item(lCount).NameAlt & "|" & Chr$(13) & "|" & _
   cAttributes.Item(lCount).Name & "|" & Chr$(13) & _
   "-------"
    
    'If StrComp(cAttributes.Item(lCount).NameAlt, theName, CompareMethod.TextCompare) Then
    If cAttributes.Item(lCount).NameAlt = theName Then
        sName = cAttributes.Item(lCount).Name
        'Debug.Print sName & Chr$(13) & "-------"
        Exit Do
    End If
    lCount = lCount + 1
Loop

pGetAttributeName = sName

pGetAttributeName_Exit:
    Exit Function
    
pGetAttributeName_Error:
    'Debug.Print "Core Entity, GetAttributeName - " & Err.Description
    Resume pGetAttributeName_Exit

End Function

Private Function pErrTxt() As String

End Function

