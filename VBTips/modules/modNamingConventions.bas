Attribute VB_Name = "modNamingConventions"
'**********************************************************
'Variable Scope Prefixes
'Scope              Prefix              Example
'Global                 g           gstrUserName
'Module-level           m           mblnCalcInProgress
'Local to procedure     None        dblVelocity

'Variable Type Prefixes
'
'Convention for noting the data type of a variable.
'Suggested prefixes for Type are as follows.
'
'   Type   Description      Example
'   a       Array           aMonths
'   c       Character       cLastName
'   y       Currency        yCurrentValue
'   d       Date            dBirthDay
'   t       Datetime        tLastModified
'   b       Double          bValue
'   f       Float           fInterest
'   l       Logical         lFlag
'   n       Numeric         nCounter
'   o       Object          oEmployee
'   u       Unknown         uReturnValue

'*********************************************************
'Constants
'The body of constant names should be mixed case with capitals initiating each word. Although standard Visual Basic constants do not include data type and scope information, prefixes like i, s, g, and m can be very useful in understanding the value and scope of a constant. For constant names, follow the same rules as variables. For example:
'mintUserListMax      'Max entry limit for User list
                  '(integer value,local to module)
'gstrNewLine            'New Line character
                  '(string, global to application)
'**********************************************************
' Naming Standards
' Make all subs & Functions begin with a noun
' Omit an underscore, thus helping to distinguish between your subs
' and VB's event proceedures
' Use standard suffixes
'
'   Suffix  Usage                                       Example
'   Init    Initialize something only called once       FormInit
'   Save    save data from a form to a table            FormSave
'   Get     Retrieve a value from another form          StatusGet
'   Set     Set a value of something from another form  StatusSet
'   Check   Check the data the user typed in            FormCheck
'   Clear   Clear the controls in a from to default     FormClear
'   Load    Load a combo or list box                    mboStateLoad
'   Kill    Kill Module Level Variables                 FormKill

'**********************************************************
'Variable Data Types
'Use the following prefixes to indicate a variable's data type.

'Data type          Prefix  Example
'Boolean            bln     blnFound
'Byte               byt     bytRasterData
'Collection object  col     colWidgets
'Currency           cur     curRevenue
'Date (Time)        dtm     dtmStart
'Double             dbl     dblTolerance
'Error              err     errOrderNum
'Integer            int     intQuantity
'Long               lng     lngDistance
'Object             obj     objCurrent
'Single             sng     sngAverage
'String             str     strFName
'User-defined type  udt     udtEmployee
'Variant            vnt     vntCheckSum

'***********************************************************
'Object Naming Conventions

'Objects should be named with a consistent prefix that makes it easy to identify the type of object. Recommended conventions for some of the objects supported by Visual Basic are listed below.
'Suggested Prefixes for Controls

'Control type           prefix      Example
'3D Panel               pnl         pnlGroup
'ADO Data               ado         adoBiblio
'Animated button        ani         aniMailBox
'Check box              chk         chkReadOnly
'Combo box, list box    cbo         cboEnglish
'Command button         cmd         cmdExit
'Common dialog          dlg         dlgFileOpen
'Communications         com         comFax
'Control                ctr         ctrCurrent (used within procedures when the specific type is unknown)
'Data                   dat         datBiblio
'Data-bound combo box   dbcbo       dbcboLanguage
'Data-bound grid        dbgrd       dbgrdQueryResult
'Data-bound list box    dblst       dblstJobType
'Data combo             dbc         dbcAuthor
'Data grid              dgd         dgdTitles
'Data list              dbl         dblPublisher
'Data repeater          drp         drpLocation
'Date picker            dtp         dtpPublished
'Directory list box     dir         dirSource
'Drive list box         drv         drvTarget
'File list box          fil         filSource
'Flat scroll bar        fsb         fsbMove
'Form                   frm         frmEntry
'Frame                  fra         fraLanguage
'Gauge                  gau         gauStatus
'Graph                  gra         graRevenue
'Grid                   grd         grdPrices
'Hierarchical flexgrid  flex        flexOrders
'Horizontal scroll bar  hsb         hsbVolume
'Image                  img         imgIcon
'Image combo            imgcbo      imgcboProduct
'ImageList              ils         ilsAllIcons
'Label                  lbl         lblHelpMessage
'Lightweight check box  lwchk       lwchkArchive
'Lightweight combo box  lwcbo       lwcboGerman
'Light command button   lwcmd       lwcmdRemove
'Lightweight frame      lwfra       lwfraSaveOptions
'Light horiz scroll bar lwhsb       lwhsbVolume
'Lightweight list box   lwlst       lwlstCostCenters
'Light option button    lwopt       lwoptIncomeLevel
'Lightweight text box   lwtxt       lwoptStreet
'Light vert scroll bar  lwvsb       lwvsbYear
'Line                   lin         linVertical
'List box               lst         lstPolicyCodes
'ListView               lvw         lvwHeadings
'MAPI message           mpm         mpmSentMessage
'MAPI session           mps         mpsSession
'MCI                    mci         mciVideo
'Menu                   mnu         mnuFileOpen
'Month view             mvw         mvwPeriod
'MS Chart               ch          chSalesbyRegion
'MS Flex grid           msg         msgClients
'MS Tab                 mst         mstFirst
'OLE container          ole         oleWorksheet
'Option button          opt         optGender
'Picture box            pic         picVGA
'Picture clip           clp         clpToolbar
'ProgressBar            prg         prgLoadFile
'Remote Data            rd          rdTitles
'RichTextBox            rtf         rtfReport
'Shape                  shp         shpCircle
'Slider                 sld         sldScale
'Spin                   spn         spnPages
'StatusBar              sta         staDateTime
'SysInfo                sys         sysMonitor
'TabStrip               tab         tabOptions
'Text box               txt         txtLastName
'Timer                  tmr         tmrAlarm
'Toolbar                tlb         tlbActions
'TreeView               tre         treOrganization
'UpDown                 upd         updDirection
'Vertical scroll bar    vsb         vsbRate

'*********************************************************
'Suggested Prefixes for Data Access Objects (DAO)

'Use the following prefixes to indicate Data Access Objects.

'Database object    Prefix  Example

'Container          con     conReports
'Database           db      dbAccounts
'DBEngine           dbe     dbeJet
'Document           doc     docSalesReport
'Field              fld     fldAddress
'Group              grp     grpFinance
'Index              ix      idxAge
'Parameter          prm     prmJobCode
'QueryDef           qry     qrySalesByRegion
'Recordset          rec     recForecast
'Relation           rel     relEmployeeDept
'TableDef           tbd     tbdCustomers
'User               usr     usrNew
'Workspace          wsp     wspMine



'*************************************************************
' Code Commenting Conventions
' All procedures and functions should begin with a brief comment
' describing the functional characteristics of the procedure (what it does).
' This description should not describe the implementation details (how it does it)
' because these often change over time, resulting in unnecessary comment maintenance work,
' or worse yet, erroneous comments.
'
' The code itself and any necessary inline comments will describe the implementation.
' Arguments passed to a procedure should be described when their functions are not
' obvious and when the procedure expects the arguments to be in a specific range. Function return values and global variables that are changed by the procedure, especially through reference arguments, must also be described at the beginning of each procedure.
'
'******************************************************************************


'Procedure header comment blocks should include the following section headings:

'*******************************************************
' Purpose:  (Req)
'
' Assumes:
' Effects:
' Inputs:   (Req)                Returns:(Req)
' Comments:
' Dependants:
' Author:   James R. Fleming    Date:
'*******************************************************


'For examples, see the next section, "Formatting Your Code."

' Section heading Comment description:
'
' Purpose: What the procedure does (not how).
'
' Assumptions: List of each external variable, control, open file, or other element that is not obvious.
'
' Dependancies: Anything that is required that is not obvious (such as included controls, libraries, etc only if this is not obvious.)
' Effects: List of each affected external variable, control, or file and the effect it has (only if this is not obvious).
' Inputs: Each argument that may not be obvious. Arguments are on a separate line with inline comments.
' Returns: Explanation of the values returned by functions.
' Comments: Any additional comments about the code.
' Author: give credit/responsibility where due
' Date: Date added. This is especially important when later changed.
'
'   Formatting Your Code

' Because many programmers still use VGA displays, screen space should be conserved
' as much as possible while still allowing code formatting to reflect logic structure and nesting.
' Here are a few pointers:
'
'Standard, tab-based, nested blocks should be indented four spaces (the default).
'
'
'The functional overview comment of a procedure should be indented one space.
'The highest level statements that follow the overview comment should be indented
'one tab, with each nested block indented an additional tab.

'For example:

'*****************************************************
' Purpose:   Locates the first occurrence of a
'            specified user in the UserList array.
' Inputs:
'   strUserList():   the list of users to be searched.
'   strTargetUser:   the name of the user to search for.
' Returns:   The index of the first occurrence of the
'            rsTargetUser in the rasUserList array.
'            If target user is not found, return -1.
' Comments: Any additional comments go here.
'*****************************************************
'//The function below is part of the commenting example
'Function intFindUser (strUserList() As String, strTargetUser As String)As Integer
'   Dim i As Integer                ' Loop counter.
'   Dim blnFound As Integer         ' Target found flag.
'   intFindUser = -1
'   i = 0
'   While i <= UBound(strUserList) And Not blnFound
'      If strUserList(i) = strTargetUser Then blnFound = True
'         intFindUser = i
'      End If
'      i = i + 1
'    Wend
'End Function

