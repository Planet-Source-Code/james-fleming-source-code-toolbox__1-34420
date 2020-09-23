Attribute VB_Name = "modRecords"
Option Explicit

Public Function LoadAllRecords(lstTitle As ListBox, cmbTipType As ComboBox, cmbTBTipType As ComboBox) As String
'*****************************************************
' Purpose:  This sub handles all of the loading and displaying of all records.
' Inputs:     None    ' Returns:  None
' Comment:  This is the same routine used by start-up and clearing of search box.
'*****************************************************

    LoadAllRecords = TipCount(ListPopulate(qryList, lstTitle)) ' loads the list box
    ' loads the tbltips source type combo box ' loads the mdi source type combo box
    Call ComboboxDualLoad(cmbTipType, cmbTBTipType, qryCombo)
End Function
