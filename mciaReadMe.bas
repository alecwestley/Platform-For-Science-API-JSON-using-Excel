Attribute VB_Name = "mciaReadMe"
'++
'   Title:          mciaReadMe
'
'   Purpose:        Providing information and how to for this Template or addin
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   Abstract:       This Template provides functionality linking Excel to
'                   Platform for Science or Core LIMS
'
'                   Current Versions use the JSON API not  OData
'
'                   The target audience is those **NEW** to Visual Basic for
'                   Applications (VBA). Coding is styled for readability.
'                   This is intended as a template, start point or utility.
'
'                   For more information contact alec.westely@hotmail.com/
'                   alec.westley@thermofisher.com
'                   This module is for notes, version information and scope of
'                   functionality for this template workbook.
'
'                   **** No hardcoded values are in the code ****
'
'                   For any 'hardcoded' values look in mciConstants
'
'                   This code is "free" but the hours figuring it out/debugging
'                   it are not ;-)
'
'                   Suggest this is kept as a template not an installed Add-in
'                   for security and installation headaches.
'
'                   See end for ToDo List
'
'   Designer:       Alec
'
'   Author:         Alec
'
'   Modified:       20DEC2019
'
'   Created:        06Nov2014
'
'
'   Copyright:
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Version, Note, who - Notes HERE are for the App version
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 1.1.1 20DEC2019 Clean and tidy for publish
' 1.1   15OCT2017 Verison 1.1 Platform for Science release with base functions
' 1.0   10MAY2016 Refactored code. All previous versions are depricated.
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Functionality
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'An Excel addin that gives the user a Platform for Science.
'
'The Addin will be available from the users main excel template
'
'Communication to the Core LIMS/Platform for Science is via the API using JSON
'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Conventions
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' In general names hint to what's going on:
'   Do     - Do something, no response unless an error
'   Get    - Return a thing based on the criteria provided
'   Add    - Add a given thing(s) to a thing
'   Create - Make a new thing based on criteria provided
'   JSON   - Provide a built JSON string as VBA does not like ",' etc
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' Conventions - Ends
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
