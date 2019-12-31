Attribute VB_Name = "mciInterface"
'++
'   Title:          mciInterface
'
'   Purpose:        Separate out the public interface for the Ribbon calls
'
'   App Version:    0.2
'
'   Module Version: 0.2
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       Simple location for the stub functions supporting the Ribbon
'                   All the real nuts and bolts are in mci_Work
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Created:        06Nov2014
'
'   Modified by:    10MAY2016
'
'   Copyright:
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Version, who, Note
' 1.0 Alec No comments, it will change with use
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Option Explicit


Sub ci_Login(control As IRibbonControl)
'++
'Callback for ciLogin onAction
'--
gjSessionID = ""
Call ciDoLogIn
End Sub


Sub ci_SendDataToCore(control As IRibbonControl)
'++
'Callback for cia_SendDataToCore onAction
'--
Call ciaDoSendExperimentData
End Sub


Sub ci_Options(control As IRibbonControl)
'++
'Callback for cia_SendDataToCore onAction
'--
Call ciaOptions
End Sub





