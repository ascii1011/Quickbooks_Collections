Attribute VB_Name = "Notes"
''Notes''

'When adding a new Form, then do the following
'1. Create a global int var in "global_vars" module like so "Public sFrmFormNameHere As Integer"
    'this var will be an indicator whether or not the form is being displayed already
    'If it is displayed then it is considered active
    'A. Add sFrmFormNameHere to frmMain.prcInitFormVars and set var ( sFrmFormNameHere = 0)
    'B. Add sFrmFormNameHere to frmMain.prcClearWindows and fill with the following
            'If sFrmFormNameHere <> 0 Then
            '    Unload FrmFormNameHere
            'End If
    'C. If there are things that need to be updated in the form on a time interval while Collections is running
        'then also put sFrmFormNameHere into frmMain.prcUpdateFormSec in the following format
            'FrmFormNameHere page
            'If sFrmFormNameHere = 1 Then
            '    fill in with what needs to be refreshed
            'End If
    
'2. Create a procedure like "frmMain.prcShowFrmFormNameHere" and fill in with the following format
        'Sub prcShowFrmFormNameHere()
        '    If sFrmFormNameHere = 0 Then
        '        Load FrmFormNameHere
        '        FrmFormNameHere.Show
        '    Else
        '        MsgBox "The 'Form Name' window is already open.  If you feel you have reached this message in error, then go to 'Tools' at the top and click on 'Clear Windows'"
        '    End If
        'End Sub
    
'3. Add the previous procedure prcShowFrmFormNameHere to the frmMain.tbToolBar_ButtonClick procedure or
    'add it to a procedure that was created from the menu
    
    
    
'Adding Permissions to a certain part of Collection Application

'1. Create a variable in the "global_vars" module like so
    'iEnableNewAttribute
        'Comment:enable the New Attribute form/feature to ... describe what it is for
        'Comment:sProfileAttrNamesAry(¥) = "Enable 'Collections' Window"....
        'Comment:where ¥ = the name of the column created in the next step (step 2)
        'Public iEnableNewAttribute As Integer
        
'2. Add a Column (with the next number available) to table qb_features
    'this will indicate a new feature has been added.
    'Please mark every row appropriately, this means that the column "features_index" corrisponds to every
    'collections user and then indicate whether they have access to it or not, or in some cases which level of
    'that feature they have.
    
'3. Add a record to qb_attr
    'set column "attr_index" to the corrisponding column name that was previously created in qb_features
    'set column "attr_name" to iEnableNewAttribute from step 1
    'set column "attr_desc" Describe what this feature is for
    'set column "attr_enabled" to 1 or 0
            '1 means: the ability to have this permission active
            '1 means: the ability to have this permission inactive

'4. Make sure to update the variable "iProfileCount" by adding one to it in "global_vars" module

'5. add to frmMain.prcPresetProfileAttributes and
    'set iEnableNewAttribute = 0

'6. add to frmMain.funGrabProfile and
    'sProfileAttrDtlsAry(1, ¥) = Trim(rsProfile![¥])
    
'7. Now we are set to use sProfileAttrDtlsAry(1, ¥) anywhere any the application to restrict or allow access
    ' to this resource
