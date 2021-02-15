sub openScan

	oActiveCell = ThisComponent.getCurrentSelection()
	oConv = ThisComponent.createInstance("com.sun.star.table.CellRangeAddressConversion")
	oConv = ThisComponent.createInstance("com.sun.star.table.CellAddressConversion")
	
	oDoc = ThisComponent
	oSel = oDoc.GetCurrentSelection()
	CurrentRow = oSel.CellAddress.Row()
	oSheets = oDoc.getSheets()
	oSheet = oSheets.getByIndex(1)
	oCell = oSheet.getCellByPosition(3, CurrentRow)
	StudyInstanceUID = oCell.getString()
	oCell = oSheet.getCellByPosition(4, CurrentRow)
	SeriesInstanceUID = oCell.getString()

    msgbox StudyInstanceUID + "   " + SeriesInstanceUID
	
	shellResponse = Shell ("python",1, "/media/SSD2/cloud_data/Projects/CACSFilter/src/openScan.py " + StudyInstanceUID + " " + SeriesInstanceUID)

end sub

Sub sendEmail()
	
	filepath = "/media/SSD2/cloud_data/Projects/CACSFilter/src/text_tmp.txt"
	num = FreeFile()
	open filepath for output as #num 
	
	oDoc = ThisComponent
	oSel = oDoc.GetCurrentSelection()
	
	oSheets = oDoc.getSheets()
	oSheet = oSheets.getByIndex(2)
	oSelections = oDoc.GetCurrentSelection()
	
	If oSelections.supportsService("com.sun.star.sheet.SheetCell") Then
		oCell = oSelections
	ElseIf oSelections.supportsService("com.sun.star.sheet.SheetCellRange") Then
    	oCell = oSelections
    ElseIf oSelections.supportsService("com.sun.star.sheet.SheetCellRanges") Then
		oRanges = oSelections
		For i = 0 To oRanges.getCount() - 1
			r = oRanges.getByIndex(i)
			CurrentRow = r.CellAddress.Row()
			write #num, "---"
			write #num, CStr(CurrentRow)
			write #num, oSheet.getCellByPosition(1, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(2, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(3, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(4, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(5, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(6, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(7, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(8, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(9, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(10, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(11, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(12, CurrentRow).getString()
			write #num, oSheet.getCellByPosition(13, CurrentRow).getString()
		Next
	Else
    	Print "Something else selected = " & oSelections.getImplementationName()
 	End If
 	
	close #num
	
	shellResponse = Shell ("python",1, "/media/SSD2/cloud_data/Projects/CACSFilter/src/sendEmail.py " + filepath)
	
	'msgbox "Done"
End Sub

