Attribute VB_Name = "modStartUp"
Option Explicit
Sub Main()
'*****************************************************
' Purpose:  This is the function that launches the entire application
'           It establishes the connection to the database, opens an instance of the MDI parent
'           Displays the splash screen and opens tblTips as the default form.
' Assumes:  modConstants
' Inputs: None  ' Returns:  None
' Comments: Main must be set as the startup object
'           from the menu Project/Source_Code_Tool_Kit Properties...
'*****************************************************
On Error GoTo MainErr
    Dim iChk As Integer, iShowTips As Integer
    Dim blnDatabaseOK As Boolean                    ' to check for database opening OK
    gsDatabase = App.path & "\databases\VBTips.mdb" ' note the path is hard coded here and only here
    Call Instanciate                                ' instantiate the forms
    Call ScreenSplash(fSplash)                      ' fire off the splash screen.
    Set ftblTips = New frmtblTips                   ' create a new instance of frmtblTips
    blnDatabaseOK = DataInit                        ' open database
    If blnDatabaseOK Then
        blnDatabaseOK = DataOpen(gsDatabase)
    Else
        ErrMsgBox ("Database open didn't work")     ' oops
        Resume Next
    End If
    fMDI.Show                                       ' show the mdi form (It was loaded in ScreenSplash)
    Unload fSplash                                  ' unload the splash screen
       
    ' See what we should be shown at startup
    iChk = GetSetting(App.Title, "Settings", "chkFeatureShow", 1)
    If iChk = 0 Then frmFeatures.Show vbModal
    iShowTips = GetSetting(App.Title, "Settings", "Show Tips at Startup", 1)
    If iShowTips = 1 Then frmTipOfDay.Show vbModal  ' show the tips at start up
    ftblTips.Show                                   ' show my default startup form
    Exit Sub                                        ' exit the routine
MainErr:
    Select Case Err.Number
        Case 364                                    ' TipOfTheDay form was unloaded automatically
            ftblTips.Show                           ' show my default startup form
            Exit Sub
        Case Else
            ErrMsgBox (Err.Description & " # " & Err.Number & " Occured in subMain of modStartUp.")
            Resume Next
    End Select
End Sub

Private Sub ScreenSplash(Splash As Form)
'*****************************************************
' Purpose:  This is the function that handles the display of the splash screen
'
' Inputs: The name of the splash screen form  ' Returns:  None
' Comments: You may adjust the minimum amount of time that the splash screen
'   stays open for. As a rule 3 seconds is a good minimum. You may adjust this to
'   suit your project, but it should be open long enough so people can read your splash...
'*****************************************************
    Dim strTime As String                       ' dim a string
    strTime = Time                              ' set the string to the current time
    Splash.Show                                 ' Show the splash screen
    Do Until DateDiff("s", strTime, Time) > 4   ' display for 3 seconds
      DoEvents                                  ' releases the processor to continue application
    Loop                                        ' back to the top
    Load fMDI                                   ' load instance of MDIMain
    Call ImplodeForm(Splash, 2, 500, 1)         ' this closes the splash screen
End Sub
