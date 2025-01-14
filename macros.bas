' Copyright (C) 2025  Gary Griffin
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You can receive a copy of the GNU General Public License
' by writing to 
'  Free Software Foundation, Inc.,
' 51 Franklin Street, Fifth Floor, 
' Boston, MA 02110-1301 USA.
'
'
'
' Kit count for both Sibling and Cousin are arbitrarily set to 125. 
'
Sub onSingleCellChanged(pCell, pEvent)
' ######################
' ######################
'
' pCell is the cell that was changed
'
' If cell is within the painting area, check to see if it needs changed
'
oSheet = ActiveSheet()
If oSheet.Name <> "Main" Then
	if pCell.CellAddress.Row >= paintAreaStartRow() And pCell.CellAddress.Column >= paintAreaStartCol() And pCell.CellAddress.Row <= paintAreaEndRow() And pCell.CellAddress.Column <= paintAreaEndCol(oSheet) Then
		gpLookup = buildGrandParentLookup()
		checkLegendMatch(pCell, gpLookup)
		LoadHalfFull(pCell)
	End If
'
' If cell is a recombination point, copy to rows 21 and 34
'
	if pCell.CellAddress.Row = 1 Then
		pCell.HoriJustify = 3
		oSheet = ActiveSheet()
		oCell = oSheet.getCellByPosition(pCell.CellAddress.Column,paintAreaStartRow() - 1)
		oCell.setString(pCell.getString())
		oCell.HoriJustify = 3
		oCell.VertJustify = 2
		oCell = oSheet.getCellByPosition(pCell.CellAddress.Column,paintAreaEndRow() + 2)
		oCell.SetString(pCell.getString())
		oCell.HoriJustify = 3
		oCell.VertJustify = 2
	End If
End If

End Sub

Sub LegendMapChanged
' ######################
' ######################
Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object

' 
' Progress Meter
' 
DialogLibraries.loadLibrary("Standard")
oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
' set minimum and maximum progress value
oProgressLabel = oDialog.getModel().getByName("Label1")
oProgressLabel.Label = "Painted Segment Color Update ---"
oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
oProgressBarModel.setPropertyValue( "ProgressValueMin", 1)
oProgressBarModel.setPropertyValue( "ProgressValueMax", 23 )
' show progress bar
oDialog.setVisible( True )
'
' Refress the painted area on all chromosomes in case the Grandparent Table Format changed
'
	For Chromosome = 1 to 23
		oProgressBarModel.setPropertyValue( "ProgressValue", Chromosome )
		oSheet = getChromosomeSheet(Chromosome)
		refreshLegendMatch(oSheet)
	Next
 End Sub

Sub refreshLegendMatch(oSheet)
' ######################
' ######################
'
' refresh the complete painting area with the conditional format
'
gpLookup = buildGrandParentLookup()
i1 = paintAreaStartRow()
i2 = paintAreaEndRow()
j1 = paintAreaStartCol()
j2 = paintAreaEndCol(oSheet)

For icol = j1 to j2
	For irow = i1 to i2
		oCell = oSheet.getCellByPosition(icol,irow)
		checkLegendMatch(oCell, gpLookup)
	Next
	LoadHalfFull(oCell)
Next
End Sub

Sub resetPaintValue()
' ######################
' ######################
	oSheet = ActiveSheet()
	gpLookup = buildGrandParentLookup()
	i1 = paintAreaStartRow()
	i2 = paintAreaEndRow()
	j1 = paintAreaStartCol()
	j2 = paintAreaEndCol(oSheet)
	For icol = j1 to j2
		For irow = i1 to i2
			oCell = oSheet.getCellByPosition(icol,irow)
			oCell.setString("")
			checkLegendMatch(oCell,gpLookup)
		Next
		LoadHalfFull(oCell)
	Next
End Sub

Sub replacePaintValue()
' ######################
' ######################
'
' Take value from cell A1 and replace with value of A2 in the complete Paint Area
'
	oSheet = ActiveSheet()
	oldVal = oSheet.getCellByPosition(0,0).getString()
	newVal = oSheet.getCellByPosition(1,0).getString()
' Iterate on paint area if oldVal is non-blank
	If oldVal <> "" Then
		gpLookup = buildGrandParentLookup()
		i1 = paintAreaStartRow()
		i2 = paintAreaEndRow()
		j1 = paintAreaStartCol()
		j2 = paintAreaEndCol(oSheet)
		For icol = j1 to j2
			For irow = i1 to i2
				oCell = oSheet.getCellByPosition(icol,irow)
				if oCell.getString() = oldVal Then 
					oCell.setString(newVal)
					checkLegendMatch(oCell, gpLookup)
				End If
			Next irow
			LoadHalfFull(oCell)
		Next icol
	End If
End Sub

Sub checkLegendMatch(pCell, gpLookup)
' ######################
' ######################
'
' Check if the string in the cell matches one of the legend strings in the Main sheet. If so, set the color. If not, reset color to default
'
	resetBack = True
'	For grandparent_row = 7 To 18 ' This is the size of the GrandParent Table
'		oLegendCell = ThisComponent.Sheets(0).getCellByPosition(2,grandparent_row)
'		if oLegendCell.getString() = pCell.getString() Then
'	   		pCell.CellBackColor = oLegendCell.CellBackColor
'	   		pCell.CharColor = oLegendCell.CharColor
'	   		ResetBack = False
'		End If
'	Next
	For grandparent_row = 0 to 11
		if pCell.getString() = gpLookup(grandparent_row,0) Then
			pCell.CellBackColor = gpLookup(grandparent_row,1)
			pCell.CharColor = gpLookup(grandparent_row,2)
			resetBack = False
		End If
	Next

	If resetBack Then
		pCell.CellBackColor = RGB(255,255,255)
		pCell.CharColor = RGB(0,0,0)
	End If
'
' Update the Half Full section for this column
'
'	LoadHalfFull(pCell)
End Sub

Sub LoadUnloadCousin(pEvent)
' ######################
' ######################
'
' Toggle between load and unload of Cousin
'
oModel = pEvent.Source.Model()
If oModel.Label = "Show Cousins" Then 
  	oModel.Label = "Clear Cousins"
  	LoadCousin
Else
  	oModel.Label = "Show Cousins"
  	UnloadCousin
EndIf 

End Sub

Sub LoadCousin()
' ######################
' ######################
'
'	Load the Cousin graph(s)
'
Dim kit(125)
Dim shortName(125)
Dim kitCousin(125)
Dim kitCousinShortName(125)
Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
'
' Get the names from the Sibling and Cousin Tables
'
oDoc = ThisComponent
oSheets = oDoc.getSheets()
oSheet = oDoc.Sheets(0)
For i = 0 To siblingKits() - 1
    oCellKit = oSheet.getCellByPosition(3,23+i)
    kit(i) = oCellKit.getString()
    oCellKit = oSheet.getCellByPosition(1,23+i)
    shortName(i) = oCellKit.getString()
Next
nonSkipKit = 0
For i = 0 To 125
	if oSheet.getCellByPosition(10,i+7).getString() <> ""  and oSheet.getCellByPosition(5,i+7).getString() = "" Then
		oCellCousinKit = oSheet.getCellByPosition(10,7+i)
		kitCousin(nonSkipKit) = oCellCousinKit.getString()
		oCellCousinKit = oSheet.getCellByPosition(8,7+i)
		kitCousinShortName(nonSkipKit) = oCellCousinKit.getString()
		nonSkipKit = nonSkipKit + 1
	End If
Next
'
' draw graphs
'
irow = sumFactor(siblingKits()) + siblingKits() +  sumFactor(siblingKits()) + 1
kitCount = siblingKits() - 1
kitCousinCount = cousinKits(false) - 1
' 
' Progress Meter
' 
DialogLibraries.loadLibrary("Standard")
oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
oProgressLabel = oDialog.getModel().getByName("Label1")
oProgressLabel.Label = "Load Cousin Progress ---"
' set minimum and maximum progress value
oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
oProgressBarModel.setPropertyValue( "ProgressValueMin", 0)
oProgressBarModel.setPropertyValue( "ProgressValueMax", (kitCousinCount+1) * (kitCount + 1) )
' show progress bar
oDialog.setVisible( True )
'
' loop on Cousin kits
'
For kit2 = 0 to kitCousinCount
	For kit1 = 0 to kitCount
		oProgressBarModel.setPropertyValue( "ProgressValue", kit2 * (kitCount+1) + kit1 )
		If (kit(kit1) <> "Unknown") and (kitCousin(kit2) <> "Unknown" and kit(kit1) <> "" and kitCousin(kit2) <> "" ) Then ' Skip any kit that has a number of Unknown or blank
      		For Chromosome = 1 to 23
				oSheet = getChromosomeSheet(Chromosome)
				For i = 0 To 1
      				oRange = oSheet.getCellRangeByPosition(i,3 * irow - 1,i,3 * irow)
      				oRange.merge(True)
      			Next
				oCell = oSheet.getCellByPosition(1,3*irow-1)
				oCell.setString(shortName(kit1))
				oCell.HoriJustify = 2
				oCell.VertJustify = 2
				oCell = oSheet.getCellByPosition(0,3*irow-1)
				oCell.setString(kitCousinShortName(kit2))
				oCell.HoriJustify = 2
				oCell.VertJustify = 2
				LoadImage(kit(kit1),kitCousin(kit2),3 * irow,1)
				LoadImage(kit(kit1),kitCousin(kit2),3 * irow+1,2)
			Next
        irow = irow + 1
      Endif
   Next
Next
End Sub

Sub UnloadCousin
' ######################
' ######################
'
' Clear the Cousin Text area
'
Dim kit(125)
Dim kitCousin(125)
Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object


irow = sumFactor(siblingKits()) + siblingKits() +  sumFactor(siblingKits()) + 1
kitCount = siblingKits() - 1
kitCousinCount = cousinKits(true) - 1
' 
' Progress Meter
' 
DialogLibraries.loadLibrary("Standard")
oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
' set minimum and maximum progress value
oProgressLabel = oDialog.getModel().getByName("Label1")
oProgressLabel.Label = "UnLoad Cousin Progress ---"
oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
oProgressBarModel.setPropertyValue( "ProgressValueMin", 1)
oProgressBarModel.setPropertyValue( "ProgressValueMax", 23 )
' show progress bar
oDialog.setVisible( True )
oProgressBarModel.setPropertyValue( "ProgressValue",1 )
'
' Loop on kits
'

For kit2 = 0 to kitCousinCount
	For kit1 = 0 to kitCount
		If (kit(kit1) <> "Unknown") and (kitCousin(kit2) <> "Unknown" and kit(kit1) <> "" and kitCousin(kit2) <> "" ) Then ' Skip any kit that has a number of Unknown or blank
      		For Chromosome = 1 to 23
				oSheet = getChromosomeSheet(Chromosome)
				For i = 0 To 1
      				oRange = oSheet.getCellRangeByPosition(i,3 * irow - 1,i,3 * irow)
      				oRange.merge(False)
      			Next
      			oCell = oSheet.getCellByPosition(1,3*irow-1)
      			oCell.setString("")
      			oCell = oSheet.getCellByPosition(0,3*irow-1)
				oCell.setString("")
        	Next
        irow = irow + 1
      Endif
   Next
Next
'
' Clear the Cousin graph area
'
'
' Remove any existing Cousin Images. Leave the match images. 
'
For Chromosome = 1 to 23
	oProgressBarModel.setPropertyValue( "ProgressValue",Chromosome )
	oSheet = getChromosomeSheet(Chromosome)
    oDP = oSheet.DrawPage
'
' Determine how many images are in the match area, then skip that many when removing images
'
'	kitImageCount = 2 * sumFactor(siblingKits())
    kitImageCount = 0
    For kit1 = 0 to kitCount
		For kit2 = kit1+1 to kitCount
			If kit(kit1) <> "Unknown" and kit(kit2) <> "Unknown" Then
				kitImageCount = kitImageCount + 2
			End If
		Next
	Next
    
    For i = kitImageCount to oDP.getCount()-1
      oImage = oDP.getByIndex(kitImageCount)
      oDP.remove(oImage)
    Next
Next
End Sub

Sub LoadHalfFull(oCellChanged)
' ######################
' ######################
Dim oCell(125,1)
'
' oCellChanged is the cell that was changed
' oCell is the set of painted cells in the column that was changed
' oCellHalfFull is the set of cells that are being updated
'
'
' Get the col that was changed
'
		oSheet = ActiveSheet()
		oCellAddress = oCellChanged.getCellAddress()
		oCol = oCellAddress.Column
'
' Get the painted cells in the column that was changed
'
		for irow = 0 to siblingKits()-1
			oCell(irow,0) = oSheet.getCellByPosition(oCol,paintAreaStartRow()+irow*3)
			oCell(irow,1) = oSheet.getCellByPosition(oCol,paintAreaStartRow()+irow*3+1)
		Next
'
' Update all of the Half/Full cells in the column
'
		startRow = paintAreaEndRow() + 3
		for irow = 0 to siblingKits() - 1
			for jrow = irow+1 to siblingKits() - 1
				oCellHalfFull = oSheet.getCellByPosition(oCol,startRow)
				oCellHalfFull.CellBackColor = findHalfFull(oCell(irow,0),oCell(irow,1),oCell(jrow,0),oCell(jrow,1))
				startRow = startRow + 2
			Next
		Next
				
End Sub

Function findHalfFull(oCell1, oCell2, oCell3, oCell4)
' ######################
' ######################
'
'	Determine the match cell color (Red, Green, Yellow) based on the individual painting
'
	findHalfFull = RGB(255,255,255)
	If oCell1.getString() <> "" and oCell2.getString() <> "" and oCell3.getString() <> "" and oCell4.GetString() <> ""  Then
		if (oCell1.getString() = oCell3.getString()) and (oCell2.getString() = oCell4.getString()) Then
			findHalfFull = RGB(0,255,0)
		ElseIf oCell1.getString() = oCell3.getString() or oCell2.getString() = oCell4.getString() Then
			findHalfFull = RGB(255,255,0)
		Else
			findHalfFull = RGB(255,0,0)
		End If
	End If
End Function

Sub GenerateSegmentList
' ######################
' ######################
'
'  Create Segment sheet and populate based on chromosome matches
'
Dim shortName(125)
Dim segRow(125)

oSheets = ThisComponent.Sheets
oSheet = ThisComponent.Sheets(0)
oSheetName = "Segments"
If  oSheets.hasByName(oSheetName) Then
	oSheets.removeByName(oSheetName, oSheets.getCount())
End If
oSheets.insertNewByName(oSheetName, oSheets.getCount())
segSheet = oSheets.getbyName("Segments")
'
'  Build header
'
kitCount = siblingKits()
For i = 0 To kitCount - 1
    shortName(i) = oSheet.getCellByPosition(1,23+i).getString()
    oCell = segSheet.getCellByPosition(i * 5, 2)
    oCell.setString(shortName(i))
    oCell = segSheet.getCellByPosition(i * 5, 3)
    oCell.setString("Grandparent")
    oCell = segSheet.getCellByPosition(i * 5 + 1, 3)
    oCell.setString("Segment Info")
    segRow(i) = 4
Next i
MbpRow = paintAreaStartRow() - 2 ' Row where Mbp values stored in Chromosome Sheet
minCol = paintAreaStartCol()
' maxCol = paintAreaEndCol()
For Chromosome = 1 To 23
	cSheet = getChromosomeSheet(Chromosome)
	maxCol = paintAreaEndCol(cSheet)
	ChrText = Chromosome
	If Chromosome = 23 Then
		ChrText = "X"
	End If
	For curCol = minCol  to maxCol
		MbpVal = cSheet.getCellByPosition(curCol,MbpRow).Value 
		If MbpVal < 300 Then 
			MbpVal = MbpVal * 1000000
		End If
		If MbpVal > 0 Then
			MbpValStart = cSheet.getCellByPosition(curCol - 1,MbpRow).Value
			If MbpValStart < 300 Then
				MbpValStart = MbpValStart * 1000000
			End If
			For i = 0 to kitCount - 1
				For j = 0 to 1
					MbpName = cSheet.getCellByPosition(curCol, paintAreaStartRow() + i * 3+j).getString()
					If MbpName <> "" Then
						sCell = segSheet.getCellByPosition(i * 5, segRow(i))
						sCell.setString(MbpName)
						sCell = segSheet.getCellByPosition(I * 5 + 1, segRow(i))
						sCell.setString(ChrText & "," & MbpValStart & "," & MbpVal & ",0,0")
						segRow(i) = segRow(i) + 1
					End If
				Next j
			Next i
		End If
	Next
Next
'
' Define the Auto Filter ranges
'
For i = 0 to kitCount - 1
	if NOT ThisComponent.DatabaseRanges.hasByName(shortName(i)) Then
		oRange = segSheet.getCellRangeByPosition(i * 5, 3, i * 5 + 1, segRow(i)-1)
		oAddr = oRange.getRangeAddress()
		ThisComponent.DatabaseRanges.addNewByName(shortName(i),oAddr)
	End If
	oRange = ThisComponent.DatabaseRanges.getByName(shortName(i))
	oRange.AutoFilter = True
Next i

End Sub

Sub LoadAllImages
' ######################
' ######################
'
'  extract kit numbers from Main sheet
'
Dim kit(125)

oDoc = ThisComponent
oSheet = oDoc.Sheets(0)
For i = 0 To 125
    oCellKit = oSheet.getCellByPosition(3,23+i)
    kit(i) = oCellKit.getString()
Next
'
' Remove any existing images so that this loads new ones
'
For Chromosome = 1 to 23
	oSheet = getChromosomeSheet(Chromosome)
   oDP = oSheet.DrawPage
   For i = 0 to oDP.getCount()-1
      oImage = oDP.getByIndex(0)
      oDP.remove(oImage)
   Next
Next
'
' draw graphs
'
irow = 0
kitCount = siblingKits()
For kit1 = 0 to kitCount - 1
   For kit2 = kit1+1 to kitCount - 1
      irow = irow + 1
      If (kit(kit1) <> "Unknown") and (kit(kit2) <> "Unknown" and kit(kit1) <> "" and kit(kit2) <> "" ) Then ' Skip any kit that has a number of Unknown or blank
      	if kit(kit1) = kit(kit2) Then
      		sam = 0
      	End If
        LoadImage(kit(kit1),kit(kit2),3 * irow,1)
        LoadImage(kit(kit1),kit(kit2),3 * irow+1,2)
      Endif
   Next
Next
End Sub

Sub RebuildChromosomeSheets
' ######################
' ######################
'
' Rebuild all of the Chromosome sheets
'
For Chromosome = 1 To 23
	RemoveChromosomeSheet(Chromosome)
	CreateChromosomeSheet(Chromosome)
Next
'
' Remove Segments Sheet since it is not valid any more
'
oSheets = ThisComponent.Sheets
oSheetName = "Segments"
If  oSheets.hasByName(oSheetName) Then
	oSheets.removeByName(oSheetName, oSheets.getCount())
End If
'
' Set active to Main
'
ThisComponent.CurrentController.select(ThisComponent.Sheets(0))
'
' Create the painting rules
'
LoadLegendPainter
End Sub

Sub RemoveChromosomeSheets()
' ######################
' ######################
'
' Remove all of the Chromosome sheets
'
For Chromosome = 1 To 23
	RemoveChromosomeSheet(Chromosome)
Next
End Sub

Sub CreateChromosomeSheet(Chromosome)
' ######################
' ######################
'
' Create the Chromosome sheet if it does not exist. Format based on the number of kits in the Sibling Table
'
Dim shortName(125)
Dim initialName(125)
Dim lineFormat as New com.sun.star.table.BorderLine2

oSheets = ThisComponent.Sheets
oSheetName = "Chr" & cstr(Chromosome)
'
' Get the Names from the Sibling Table
'
oSheet = ThisComponent.Sheets(0)
For i = 0 To siblingKits() - 1
    initialName(i) = oSheet.getCellByPosition(0,23+i).getString() 
    shortName(i) = oSheet.getCellByPosition(1,23+i).getString() 
Next
'
' Create the sheet if needed and format
'
If not oSheets.hasByName(oSheetName) Then
	oSheets.insertNewByName(oSheetName, oSheets.getCount())
	oSheet = oSheets.getbyName("Chr" & cstr(Chromosome))
	kitCount = siblingKits()
	lineFormat.LineStyle = 0
	lineFormat.LineWidth =  50
	defaultCell = oSheet.getCellByPosition(0,0)
	defaultCellWidth = defaultCell.Size.Width
	defaultCellHeight = defaultCell.Size.Height
'
' Make the labels for the images
'
	irow = 0
	For kit1 = 0 To kitCount -1
		For kit2 = kit1 + 1 To kitCount -1
			oRange = oSheet.getCellRangeByPosition(0,3 * irow +2,0,3 * irow+3)
      		oRange.merge(True)
      		oCell = oSheet.getCellByPosition(0,3 * irow + 2)
      		oCell.setString(shortName(kit1))
      		oCell.CharWeight = 150
			oCell.HoriJustify = 2
			oCell.VertJustify = 2
      		oRange = oSheet.getCellRangeByPosition(1,3 * irow +2,1,3 * irow+3)
      		oRange.merge(True)
      		oCell = oSheet.getCellByPosition(1,3 * irow + 2)
      		oCell.setString(shortName(kit2))
      		oCell.CharWeight = 150
			oCell.HoriJustify = 2
			oCell.VertJustify = 2
      		irow = irow + 1
      	Next
	Next
'
' Resize Columns
'
	oRange = oSheet.getCellRangeByPosition(7,1,26,3*iRow)
	oRange.RightBorder = lineFormat
	oRange.LeftBorder = lineFormat
	oSheet.getColumns().getByName("C").Width = defaultCellWidth * 0.10
	oSheet.getColumns().getByName("F").Width = defaultCellWidth * 0.10
	oSheet.getColumns().getByName("G").Width = defaultCellWidth * 0.10
	For i = 7 To 26
		oSheet.getColumns().getByIndex(i).Width = defaultCellWidth * 0.70
	Next
'
' Resize rows - include Cousins
'
'	For jrow = 1 To 300
'		oSheet.getRows().getByIndex(jrow).Height = defaultCellHeight * 1.5
'	Next
'
' Make the labels for the painting area
'
	for kit1 = 0 To kitCount - 1
		oRange = oSheet.getCellRangeByPosition(2,3 * irow + 3 + 3 * kit1,5,3 * irow + 3 + 3 * kit1 + 1)
		oRange.merge(True)
		oCell = oSheet.getCellByPosition(2,3 * irow + 3 + 3 * kit1)
		oCell.setString(shortName(kit1))
		oCell.CharWeight = 150
		oCell.HoriJustify = 2
		oCell.VertJustify = 2
	Next

	oRange = oSheet.getCellRangeByPosition(7,3 * irow + 3,26,3*iRow + 3 * kitCount+1)
	oRange.RightBorder = lineFormat
	oRange.LeftBorder = lineFormat
	oRange.TopBorder = lineFormat
	oRange.BottomBorder = lineFormat
'
' Create Named Range in sheet
'
	oRanges = ThisComponent.NamedRanges
	Dim oCellAddress As new com.sun.star.table.CellAddress
	oCellAddress.Sheet = oSheets.getCount() - 1
	oCellAddress.Column = 7
	oCellAddress.Row = 3 * irow + 3
	If oRanges.hasByName(oSheetName & "paint") Then
		oRanges.removeByName(oSheetName & "paint")
	End If
	oRanges.addNewByName(oSheetName & "paint",oRange.AbsoluteName,oCellAddress,0)
'
' Make the labels for the Segment Match area
'
	startRow = (sumFactor(kitCount)+kitCount) * 3 + 4
	For kit1 = 0 To kitCount -1
		For kit2 = kit1 + 1 To kitCount -1
      		oCell = oSheet.getCellByPosition(3, startRow)
      		oCell.setString(shortName(kit1))
      		oCell.CharWeight = 150
			oCell.HoriJustify = 2
			oCell.VertJustify = 2
      		oCell = oSheet.getCellByPosition(4, startRow)
      		oCell.setString(shortName(kit2))
      		startRow = startRow + 2
      		oCell.CharWeight = 150
			oCell.HoriJustify = 2
			oCell.VertJustify = 2
      	Next
	Next
'
' Make cell borders 
'
	startRange = (sumFactor(kitCount)+kitCount) * 3 + 4
	oRange = oSheet.getCellRangeByPosition(7,startRange,26, startRange+ 2 * sumFactor(kitCount)-2 )
	oRange.RightBorder = lineFormat
	oRange.LeftBorder = lineFormat
	oRange.TopBorder = lineFormat
	oRange.BottomBorder = lineFormat
'
' Freeze cells
'
	ThisComponent.CurrentController.select(oSheet)
	ThisComponent.CurrentController.FreezeAtPosition(0,startRange)
	
End If
End Sub

Sub RemoveChromosomeSheet(Chromosome)
' ######################
' ######################
oSheets = ThisComponent.Sheets
oSheetName = "Chr" & cstr(Chromosome)
If  oSheets.hasByName(oSheetName) Then
	oSheets.removeByName(oSheetName, oSheets.getCount())
End If
End Sub

Sub LoadImage(kit1,kit2,yPos,imgFlag)
' ######################
' ######################
'
' Find the folder that contains the images
'

On Error GoTo ErrorHandler
oURL = ThisComponent.getURL()
oTitle = ThisComponent.Title
Folder = Left(oURL,inStr(oURL,oTitle)-2) & "/GEDmatchImages/"
'
' Loop over each sheet
'
For Chromosome = 1 to 23
   oImagen_obj = ThisComponent.createInstance("com.sun.star.drawing.GraphicObjectShape")
   imagen = kit1 & "_" & kit2 & "_" & cstr(Chromosome) & "_"  & cstr(imgFlag) &  ".gif"
   ImagenURL = convertToURL(Folder & imagen)
   oImagen_obj.GraphicURL = ImagenURL
'
' Get the sheet and location for the graphic
'
	oSheet = getChromosomeSheet(Chromosome)
	oDP = oSheet.DrawPage
	oSize = oImagen_obj.Size
'
' If there is an existing graph on the sheet, we size to that. Otherwise we size to paintArea calc
'
	If oDP.getCount() > 0 Then
		oImage = oDP.getByIndex(0)
		oSize.Height = oImage.Size.Height
		oSize.Width = oImage.Size.Width
	Else
		oCell = oSheet.getCellByPosition(paintAreaStartCol(),yPos-1) ' Convert to cell address
		oCell2 =  oSheet.getCellByPosition(paintAreaEndCol(oSheet)+1,yPos) 'End of graphic cell
		oSize.Height = oCell2.Position.Y - oCell.Position.Y
		oSize.Width = oCell2.Position.X - oCell.Position.X
	End If


	oCell = oSheet.getCellByPosition(paintAreaStartCol(),yPos-1) ' Convert to cell address
'	oCell2 =  oSheet.getCellByPosition(paintAreaEndCol(oSheet),yPos) 'End of graphic cell
'
' Calculate size of graphic based on cell shape.
'
'	oSize = oImagen_obj.Size
'	oSize.Height = oCell2.Position.Y - oCell.Position.Y
'	oSize.Width = oCell2.Position.X - oCell.Position.X
	oImagen_obj.Size = oSize
	oPos = oImagen_obj.Position
'
' Update graphic
'
   oPos.X = oCell.Position.X
   oPos.Y = oCell.Position.Y
   oImagen_obj.Position = oPos
   oImagen_obj.Anchor = oCell
'
' Draw graphic
'
   oDP = oSheet.DrawPage
   oDP.add(oImagen_obj)

Next
Exit Sub
ErrorHandler:
	Resume Next
End Sub

Sub LoadLegendPainter
' ######################
' ######################
'
' Update the callback for Content Changed 
'
dim Prop(1) as new com.sun.star.beans.PropertyValue

	Prop(0).name = "EventType"
	Prop(0).value = "Script"
	Prop(1).name = "Script"
	Prop(1).value = "vnd.sun.star.script:Standard.Module1.onContentChanged?language=Basic&location=document" 
	For Chromosome = 1 to 23
		oSheet = getChromosomeSheet(Chromosome)
   		oSheet.Events.replaceByName("OnChange", Prop())
	Next
	oSheet = ThisComponent.Sheets(0)
	oSheet.Events.replaceByName("OnChange", Prop())
End Sub

Function paintAreaStartRow
' ######################
' ######################
'	paintAreaStartRow = 21 ' Row painting area starts
	paintAreaStartRow = 3  * sumFactor(siblingKits()) + 3
End Function

Function paintAreaEndRow
' ######################
' ######################
'	paintAreaEndRow = 31 ' Row painting area ends
	paintAreaEndRow = paintAreaStartRow() + 3 * siblingKits() -2
End Function

Function paintAreaStartCol
' ######################
' ######################
	paintAreaStartCol = 7 ' Column painting area starts at H
End Function

Function paintAreaEndCol(oSheet)
' ######################
' ######################
'	paintAreaEndCol = 27 ' Column painting area ends at AH
	oRange = ThisComponent.NamedRanges.getByName(oSheet.Name & "paint")
	paintAreaEndCol = 7 + oRange.getReferredCells().Columns.Count - 1
End Function

Function ActiveSheet
' ######################
' ######################
'
' Function to return the Active Sheet
'
	ActiveSheet=ThisComponent.CurrentController.getActiveSheet()
End Function

Function siblingKits
' ######################
' ######################
'
' Function to define the number of Siblng kits in the Sibling table
'
	siblingKits = 0
	For i = 0 To 125 ' Sibling Kit max number
		if ThisComponent.Sheets(0).getCellByPosition(3,i+23).getString() <> "" Then
			siblingKits = siblingKits + 1
		End If
	Next
End Function

Function sumFactor(num)
' ######################
' ######################
'
' Working function to determine the sum of the numbers. This is used to count the N X N kits
'
    sumFactor = 0
	For i = 0 to num-1
		sumFactor = sumFactor + i
	Next
End Function

Function cousinKits(includeSkipped)
' ######################
' ######################
'
' Function to define the number of Cousin Kits in the Cousin table
'
	cousinKits = 0
	If includeSkipped Then
		For i = 0 to 125
			if ThisComponent.Sheets(0).getCellByPosition(10,i+7).getString() <> "" Then
				cousinKits = cousinKits + 1
			End If
		Next
	Else
		For i = 0 To 125 ' Sibling Kit max number
			if ThisComponent.Sheets(0).getCellByPosition(10,i+7).getString() <> ""  and ThisComponent.Sheets(0).getCellByPosition(5,i+7).getString() = "" Then
				cousinKits = cousinKits + 1
			End If
		Next
	End If
End Function

Function getChromosomeSheet(Chromosome)
' ######################
' ######################
'
' Function to build the Chromosome Sheet name
'
	getChromosomeSheet = ThisComponent.getSheets().getbyName("Chr" & cstr(Chromosome))
End Function

Function buildGrandParentLookup()
' ######################
' ######################
	Dim localGrandParent(11,2)
'
' Function to build the Chromosome Sheet name
'
	gpSheet =  ThisComponent.Sheets(0)
	For grandparent_row = 0 To 11 ' This is the size of the GrandParent Table
		gpCell = gpSheet.getCellByPosition(2,grandparent_row+7)
		localGrandParent(grandparent_row,0) =gpCell.getString()
		localGrandParent(grandparent_row,1) = gpCell.CellBackColor
		localGrandParent(grandparent_row,2) = gpCell.CharColor
	Next grandparent_row
	buildGrandParentLookup = localGrandParent
End Function

Sub Main
'
' stub for debugging
'
'	oCell = ThisComponent.Sheets(0).getCellByPosition(2,17)
'
'	iCell = 0
'	kitMax = sumFactor(30)
'
'   odoc = ThisComponent
'   oSheet = oDoc.Sheets(0)
'   Prop(0).name = "EventType"
'   Prop(0).value = "Script"
'   Prop(1).name = "Script"
'   Prop(1).value = "vnd.sun.star.script:Standard.Module1.onMainContentChanged?language=Basic&location=document" 
'
'   ThisComponent.Sheets(0).Events.removeByName("OnChange")
'
' Build paint Named Ranges
'
	oSheets = ThisComponent.Sheets
	irow = 6
	kitCount = 4
	For Chromosome = 1 to 23
		oSheetName = "Chr" & cstr(Chromosome)
		oSheet = oSheets.getbyName(oSheetName)
		oRange = oSheet.getCellRangeByPosition(7,3 * irow + 3 ,26,3*iRow + 3 * kitCount+1 )
		oRanges = ThisComponent.NamedRanges
		Dim oCellAddress As new com.sun.star.table.CellAddress
		oCellAddress.Sheet = oSheets.getCount() - 1
		oCellAddress.Column = 7
		oCellAddress.Row = 3 * irow + 3 + 2
		oRanges.addNewByName(oSheetName & "paint",oRange.AbsoluteName,oCellAddress,0)
	Next Chromosome
'	removeChromosomeSheet(33)
'	createChromosomeSheet(33)
End Sub

'
' Routines borrowed
'

Sub onContentChanged(pEvent)
REM It may be necssary to split the handling in additional ways.
REM pEvent is passed (in addition) unchanged throug the chain of Sub
REM to allow for local tests.
REM The procedures here only give the structure.
If pEvent.SupportsService("com.sun.star.sheet.SheetCellRanges") Then 
    onRangesChanged(pEvent, pEvent)
End If
If pEvent.SupportsService("com.sun.star.sheet.SheetCellRange") Then
    onSingleRangeChanged(pEvent, pEvent)
End If
End Sub

Sub onRangesChanged(pRanges, pEvent)
REM It may be necssary to split the handling in additional ways.
For i = 0 To pRanges.Count - 1
    onSingleRangeChanged(pRanges(i), pEvent)
Next i
End Sub

Sub onSingleRangeChanged(pRange, pEvent)
REM It may be necssary to split the handling in additional ways.
u1 = pRange.Columns.Count - 1 : u2 = pRange.Rows.Count - 1
For j = 0 To u1
    For k = 0 To u2
        theCell = pRange.getCellByPosition(j, k)
        onSingleCellChanged(theCell, pEvent)
    Next k
Next j
End Sub 

