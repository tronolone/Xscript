'Add reference: Excel 2007 Interop DLL
'Add reference: Xscript Actions
'Add reference: Xscript Classes

Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions

Public Module ExcelModule
	
	'Persistent objects
	Dim dicExcelFileCache As New Dictionary(Of String, dictionary(Of String, Object(,)))
	
'==============================================================================================================
' FnSheetExistsInCache
'
' 		Returns a boolean to indicate if the workbook and internal worksheet have already been read in and
'		cached by Xscript.
'
' 		@strExcelFilePath:	Full file path to the Excel file
' 		@strSheetName:		Name of the sheet to look for in the cache
'==============================================================================================================

	Public Function FnSheetExistsInCache(ByVal strExcelFilePath As String, ByVal strSheetName As String) As Boolean
		
		'Convert to lower immediately
		strExcelFilePath = strExcelFilePath.ToLower
		strSheetName = strSheetName.ToLower
		
		If dicExcelFileCache.ContainsKey(strExcelFilePath) AndAlso dicExcelFileCache.Item(strExcelFilePath).ContainsKey(strSheetName) Then
			Return True
		Else
			Return False
		End If
		
	End Function


'==============================================================================================================
' FnLoadSimpleList
'
' 		Loads a single column of data from an Excel workbook into a List object with a single IO read.  Optional
'		parameters available to perform the operation across all sheets in the workbook.
'
'		If the user does not specify the header of the column (the value of the first cell in the column), the 
'		function will assume that the desired data is contained in the first column and has no header.
'
' 		@strPathToExcelFile:		The full file path to the Excel file	
'		@strHeaderLabel:			The value of the header cell over your column of data
'		@blnCeaseWhenBlankFound:	If true, data will cease aggregating once a blank cell is reached
'		@blnGetFirstSheetOnly:		If true, data will only be obtained from the first worksheet
'		@strGetTheseSheetsOnlyArr:	If supplied, data will only be retrieved from worksheets named in this array
'==============================================================================================================

	Public Function FnLoadSimpleList(ByVal strPathToExcelFile As String, _ 
									 Optional ByVal strHeaderLabel As String="", _ 
									 Optional ByVal blnCeaseWhenBlankFound As Boolean=True, _ 
									 Optional ByVal blnGetFirstSheetOnly As Boolean=True, _
									 Optional ByVal strGetTheseSheetsOnlyArr() As String=Nothing) As List(Of String)
	
		'Declare Objects
		Dim dicData As Dictionary(Of String, Object(,))
		
		'Declare variables		
		Dim intNumUsedRows As Integer
		Dim intNumUsedColumns As Integer
		Dim strCellValue As String
		Dim intOnlyColumnThatMatters As Integer
		Dim blnWeHitABlank As Boolean
		
		'Set default values
		intOnlyColumnThatMatters = 0
		blnWeHitABlank = False
		
		'Initialize list
		FnLoadSimpleList = New List(Of String)
		
		'Convert the user-entered array of sheet names to lowercase
		If Not strGetTheseSheetsOnlyArr Is Nothing Then
			For i As Integer = 0 To strGetTheseSheetsOnlyArr.Length
				strGetTheseSheetsOnlyArr(i) = strGetTheseSheetsOnlyArr(i).ToLower
			Next i
		End If
		
		'Form a dictionary of 2D obj arrays keyed by sheet name from the Excel sheet
		dicData = FnFetchExcelData(strPathToExcelFile)
		
		'Ensure there is some data available
	    If Not dicData Is Nothing Then
			
			'Loop through every sheet
			For Each kvp As KeyValuePair(Of String, Object(,)) In dicData
						
				'If the user didn't specify which sheets to get OR this sheet is included in the array of sheets to get, get the sheet's data
				If(strGetTheseSheetsOnlyArr Is Nothing OrElse Array.IndexOf(strGetTheseSheetsOnlyArr, kvp.Key.ToLower)<>-1) Then
				
					'Count the used rows
					intNumUsedRows = kvp.Value.GetUpperBound(0)
					
					'Count the used columns
					intNumUsedColumns = kvp.Value.GetUpperBound(1)
					
					'Loop over columns
				    For intCurrentColumn As Integer = 1 To intNumUsedColumns
						
						'Loop over rows
						For intCurrentRow As Integer = 1 To intNumUsedRows
											
							'Get the content of this cell
							strCellValue = kvp.Value(intCurrentRow, intCurrentColumn)
							
							'If this is the first row AND we haven't already amassed our data AND [the user didn't specify a header Or this Is the header he specified] Then Set a flag To retain all data In this column
							If(intCurrentRow = 1 And intOnlyColumnThatMatters = 0 And (String.IsNullOrEmpty(strHeaderLabel) OrElse String.Compare(strCellValue, strHeaderLabel, True)=0)) Then
								intOnlyColumnThatMatters = intCurrentColumn
							End If
							
							'If we're currently looping through the only column that matters AND [we're past row 1 OR the user didn't supply a header] THEN add to the list
							If intOnlyColumnThatMatters = intCurrentColumn And (intCurrentRow > 1 Or String.IsNullOrEmpty(strHeaderLabel) = True) Then
								
								'If the user wants to cease once we hit blanks, we need to check and maybe abort the loop here
								If(blnCeaseWhenBlankFound And String.IsNullOrEmpty(strCellValue)) Then Exit For
								
								'Add the value
								FnLoadSimpleList.Add(strCellValue)	
							
							End If
						
					    Next intCurrentRow
						
					Next intCurrentColumn
					
				End If
				
				If blnGetFirstSheetOnly Then Exit For

			Next kvp
			
	    End If

	End Function

'==============================================================================================================
' FnLoadSimpleDictionary
'
' 		Creates a dictionary object from two columns in an Excel spreadsheet using a single IO read.  Optional
'		parameters available to perform the operation across all sheets in the workbook.
'
'		If the user does not specify the name of the header containing the list of keys, the function assumes 
'		that it's a column aptly named "Name".  If the user does not specify the name of the header containing
'		the list of corresponding values, the function assumes that it's a column named "Value".
'
'		@strErrorMsg			A variable passed ByRef that will be filled with an error message if needed
' 		@strPathToExcelFile:	The full file path to the Excel file	
'		@strNameLabel:			The name of the column containing the keys (the names of your variables)
'		@strValueLabel:			The name of the corresponding column containing variable values
'		@blnIncludeAllSheets:	If true, the operation will gather data from across all available spreadsheets
'		@strSheetPriorityArr:	If supplied, data will only be retrieved from worksheets named in this array
'==============================================================================================================

	Public Function FnLoadSimpleDictionary(ByRef strErrorMsg As String, _
										   ByVal strPathToExcelFile As String, _
										   Optional ByVal strNameLabel As String="Name", _
										   Optional ByVal strValueLabel As String="Value", _
										   Optional ByVal blnIncludeAllSheets As Boolean=False, _
										   Optional ByVal strSheetPriorityArr() As String=Nothing) As Dictionary(Of String, String)
		
		'Declare Objects
		Dim dicData As Dictionary(Of String, Object(,))
		Dim dicDataKeys As List(Of String)
		
		'Declare variables
		Dim strCellValue As String = ""
		Dim strVarName As String = ""
		Dim intVarNameCol As Integer = 0
		Dim intVarValueCol As Integer = 0
		Dim intRound As Integer = 0	
		Dim j As Integer = 0
		Dim blnAddedDataSheetThisRound As Boolean = False
		Dim intNumUsedRows, intNumUsedColumns As Integer
		
		'Set Objects
		FnLoadSimpleDictionary = New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
		
		'Get the data
		dicData = FnFetchExcelData(strPathToExcelFile)
		
		'Convert the array of sheet priorities to lowercase
		If Not strSheetPriorityArr Is Nothing Then
			For i As Integer = 0 To strSheetPriorityArr.Length-1
				strSheetPriorityArr(i) = strSheetPriorityArr(i).ToLower
			Next i
		End If
		
	    'Assuming data was returned
	    If Not dicData Is Nothing Then
			
			'Obtain all the keys
			dicDataKeys = New List(Of String)(dicData.Keys)
			
			'Loop through each dicData key
			Do While j <= dicDataKeys.Count-1
				
				'[If this is round 0 AND we are to include all sheets AND [we were not given a priority array OR this sheet is not a priority]] OR
				'[If this is round 0 AND we are not meant to include all sheets AND we were not given any priority array] OR
				'[We are past the round 0 AND we were given a priority array AND this sheet is in the priority array AND the index of this priority equals the number of priorities minus current round]
				If (intRound=0 AndAlso blnIncludeAllSheets AndAlso (strSheetPriorityArr Is Nothing OrElse Array.IndexOf(strSheetPriorityArr, dicDataKeys(j).ToLower)=-1)) OrElse _
				   (intRound=0 AndAlso blnIncludeAllSheets=False AndAlso strSheetPriorityArr Is Nothing) OrElse _
				   (intRound>0 AndAlso Not strSheetPriorityArr Is Nothing AndAlso Array.IndexOf(strSheetPriorityArr, dicDataKeys(j).ToLower)>=0 AndAlso Array.IndexOf(strSheetPriorityArr, dicDataKeys(j).ToLower)=strSheetPriorityArr.length-intRound) Then
					
					'Set a boolean to indicate that we're successfully adding data during this whole round at least once
					blnAddedDataSheetThisRound = True
				
					'Count the used rows
					intNumUsedRows = dicData.Item(dicDataKeys(j)).GetUpperBound(0)
					
					'Count the used columns
					intNumUsedColumns = dicData.Item(dicDataKeys(j)).GetUpperBound(1)
						
					'Loop over rows
					For intCurrentRow As Integer = 1 To intNumUsedRows
							
						'Loop over columns
					    For intCurrentColumn As Integer = 1 To intNumUsedColumns
							
							'Get the content of this cell
							strCellValue = dicData.Item(dicDataKeys(j))(intCurrentRow, intCurrentColumn)
							
							'If this is the first row
							If(intCurrentRow = 1) Then
							
								'If the cell value equals the given strNameLabel, retain the column number
								If(String.Compare(strCellValue, strNameLabel, True) = 0) Then
									intVarNameCol = intCurrentColumn
								
								'If the cell value equals the given strValueLabel, retain the column number
								Else If(String.Compare(strCellValue, strValueLabel, True) = 0) Then
									intVarValueCol = intCurrentColumn
								End If
								
							'If we're on the var name column, retain it
							Else If(intCurrentColumn = intVarNameCol) Then
								strVarName = strCellValue
								
							'If this is the value column AND strName is not blank, add to dictionary
							Else If(intCurrentColumn = intVarValueCol And String.IsNullOrEmpty(strVarName) = False) Then
								If(FnLoadSimpleDictionary.ContainsKey(strVarName) = False) Then
									FnLoadSimpleDictionary.Add(strVarName, strCellValue)
								Else
									FnLoadSimpleDictionary.Item(strVarName) = strCellValue
								End If
							End If
							
					    Next intCurrentColumn
						
					Next intCurrentRow
			
				End If
				
				'If we're at the end of the dicDataKeys, figure out if we need another round in order to tack on the higher prioritized data sheets
				If j = dicDataKeys.Count-1
					
					'Ensure we actually added data during that round (unless it was the first round and we arent including all sheets)
					If blnAddedDataSheetThisRound OrElse (blnIncludeAllSheets=False AndAlso intRound=0)Then
						
						'Reset the boolean which indicates the success addition of a sheet
						blnAddedDataSheetThisRound = False
						
						'If there are more priorities to add, reset data to cycle through another round
						If Not strSheetPriorityArr Is Nothing AndAlso intRound < strSheetPriorityArr.Length
							j = -1
							intRound = intRound + 1
						End If
						
					'If we haven't added any data and this isn't round 0, we errored in some way
					Else If blnAddedDataSheetThisRound=False AndAlso intRound > 0 Then
						strErrorMsg = "There is no sheet named '" & strSheetPriorityArr(strSheetPriorityArr.Length-intRound) & "'."
						Exit Do
					End If
					
				End If
				
				j = j + 1
				
			Loop
					
		End If
				
	End Function

	
'==============================================================================================================
' FnCreateComplexTDAA
'
' 		Creates a 2D associative array from a spreadsheet with a single IO read.
'
'		User can choose to have the 2D array keyed by row and then column or vice versa.  User can only also
'		choose to have the rows keyed by the values found in the first column and/or have the columns keys by
'		the values found in the first row.  Alternatively, the user may also indicate which attribute should be
'		used as a record's idenitifying key.  Duplicate keys are handled by appending (#) onto the end Of the key
'		name.  If the user indicates that rows Or columns are unkeyed, integers will be used for keys instead.
'
' 		@strPathToExcelFile:			The full path to the Excel file to load	
'		@strSheetToUse					Name of worksheet to gather data from
'		@blnKeyByRowThenColumn:			Boolean indicating the data should be keyed by row then column
'		@blnFirstRowIsIdentifiers:		Boolean indicating the that first row in the sheet contains column headers
'		@blnFirstColumnIsIdentifiers:	Boolean indicating the that first column in the sheet contains row labels
'		@strAttributeToKeyUpon:			Optional parameter indicating which attribute to key records by
'
'		Example #1
'			Dim myTDAA As AssociativeArray2D
'			myTDAA = New AssociativeArray2D
'			myTDAA = FnCreateComplexTDAA("C:\EmployeeData.xlsx", True, True, True)
'			Console.Writeline("Charles Tronolone's hire date is: " & myTDAA.Item("Charles Tronolone", "Hire Date"))
'			For Each key1 As String In myTDAA.Keys 
'				For Each key2 As String In myTDAA.Keys(key1)
'					Console.Writeline("myTDAA('" & key1 & "', '" & key2 & "') = " & myTDAA.Item(key1, key2))
'				Next
'			Next
'==============================================================================================================

	Public Function FnCreateComplexTDAA(ByVal strPathToExcelFile As String, _
										ByVal strSheetToUse As String, _
										ByVal blnKeyByRowThenColumn As Boolean,	_
										ByVal blnFirstRowIsIdentifiers As Boolean, _
										ByVal blnFirstColumnIsIdentifiers As Boolean, _
										Optional ByVal strAttributeToKeyUpon As String="") As AssociativeArray2D
		
		'Declare Objects
		Dim dicData As Dictionary(Of String, Object(,))
		Dim listAttributeLabels = New List(Of String)
		
		'Declare variables
		Dim strCellValue As String	
		Dim strRecordLabel As String
		Dim intNumRecords As Integer
		Dim intNumAttributes As Integer
		Dim intTemp As Integer
		Dim strTemp As String
		Dim blnRecordLabelsExist As Boolean
		Dim blnAttributeLabelsExist As Boolean
		
		'Set things
		FnCreateComplexTDAA = New AssociativeArray2D
		strRecordLabel = ""
		
		'Lowercase some user input for comparision ease
		If(Not String.IsNullOrEmpty(strAttributeToKeyUpon)) Then strAttributeToKeyUpon = strAttributeToKeyUpon.ToLower
		If(Not String.IsNullOrEmpty(strSheetToUse)) Then strSheetToUse = strSheetToUse.ToLower
		
		'Create some easy-to-comprehend booleans to indicate if record and attribute labels exist.
		If(blnKeyByRowThenColumn) Then
			If(blnFirstColumnIsIdentifiers) Then blnRecordLabelsExist = True
			If(blnFirstRowIsIdentifiers) Then blnAttributeLabelsExist = True
		Else
			If(blnFirstColumnIsIdentifiers) Then blnAttributeLabelsExist = True
			If(blnFirstRowIsIdentifiers) Then blnRecordLabelsExist = True
		End If
		
		'Select the sheet to use
		dicData = FnFetchExcelData(strPathToExcelFile)
		
	    'Scan the cells
	    If dicData.Item(strSheetToUse) IsNot Nothing Then
			
			'Determine how many records and attributes are used on the spreadsheets
			If(blnKeyByRowThenColumn) Then
				intNumRecords = dicData.Item(strSheetToUse).GetUpperBound(0)
				intNumAttributes = dicData.Item(strSheetToUse).GetUpperBound(1)
			Else
				intNumRecords = dicData.Item(strSheetToUse).GetUpperBound(1)
				intNumAttributes = dicData.Item(strSheetToUse).GetUpperBound(0)
			End If
				
			'Loop over records
		    For intCurrentRecord As Integer = 1 To intNumRecords
				
				'Loop over each attribute for this record
				For intCurrentAttribute As Integer = 1 To intNumAttributes
					
					'Get the content of this cell
					If(blnKeyByRowThenColumn) Then
						strCellValue = dicData.Item(strSheetToUse)(intCurrentRecord, intCurrentAttribute)
					Else
						strCellValue = dicData.Item(strSheetToUse)(intCurrentAttribute, intCurrentRecord)
					End If
					
					'If this is the first cell of an entire record, determine a name for the record label
					If(intCurrentAttribute = 1) Then
												
						'Reset variables
						strRecordLabel= ""
													
						'Regardless of whether or not record labels exist, if the user specified a particular attribute to use
						'as the record label, then we are going to ignore the official record labels and set a new one now
						If(Not String.IsNullOrEmpty(strAttributeToKeyUpon)) Then
							
							'Get the index of the specified attribute
							intTemp = listAttributeLabels.IndexOf(strAttributeToKeyUpon)
							
							'Assuming the attribute exists, get the value for this record at the specified attribute
							If(intTemp >= 0) Then
								If(blnKeyByRowThenColumn) Then
									strRecordLabel = dicData.Item(strSheetToUse)(intCurrentRecord, intTemp + If(blnRecordLabelsExist, 2, 1))
								Else
									strRecordLabel = dicData.Item(strSheetToUse)(intTemp + If(blnRecordLabelsExist, 2, 1), intCurrentRecord)
								End If							
							End If
						
						'If the user didn't send a specific attribute to use as a record label but there are existing records labels, use the existing record label
						Else If(blnRecordLabelsExist) Then
							strRecordLabel = strCellValue
						End If
						
						'If the strRecordLabel is blank, we need to make our own label from scratch
						If(String.IsNullOrEmpty(strRecordLabel) Or String.IsNullOrWhiteSpace(strRecordLabel)) Then							
							If(blnAttributeLabelsExist=False) Then
								strRecordLabel = intCurrentRecord.ToString
							Else
								strRecordLabel = (intCurrentRecord-1).ToString
							End If
						End If
						
						'In case there are duplicate record labels in the spreadsheet, rename our record label here
						If(FnCreateComplexTDAA.ContainsKey(strRecordLabel)) Then
							intTemp = 2
							Do Until FnCreateComplexTDAA.ContainsKey(strRecordLabel & " (" & intTemp.ToString & ")") = False
								intTemp = intTemp + 1 
							Loop
							strRecordLabel = strRecordLabel & " (" & intTemp.ToString & ")"
						End If
																		
					End If
					
					'If this is the first record, and we're not on the first attribute OR we are but there are no record labels, get/create attribute label
					If(intCurrentRecord = 1 And (intCurrentAttribute > 1 Or blnRecordLabelsExist=False)) Then		
						
						'If the user specified attribute labels and this cell isn't blank, then make it a candidate label
						If(blnAttributeLabelsExist And Not String.IsNullOrEmpty(strCellValue) And Not String.IsNullOrWhiteSpace(strCellValue)) Then
							strTemp = strCellValue.ToLower
							
						'Otherwise, we have to invent our own attribute label
						Else
							If(blnRecordLabelsExist) Then
								strTemp = (intCurrentAttribute-1).ToString
							Else
								strTemp = intCurrentAttribute.ToString
							End If
						End If
						
						'In case there are duplicate attribute labels in the spreadsheet, rename our label here
						If(listAttributeLabels.Contains(strTemp)) Then
							intTemp = 2
							Do Until listAttributeLabels.Contains(strTemp & " (" & intTemp.ToString & ")") = False
								intTemp = intTemp + 1 
							Loop
							strTemp = strTemp & " (" & intTemp.ToString & ")"
						End If
						
						'Add the label to the list
						listAttributeLabels.Add(strTemp)
						
					End If
					
					'This condition ensures that only actual data is passed into the 2D array and we don't mistakenly add labels into the array
					If((intCurrentRecord > 1 Or blnAttributeLabelsExist=False) And (intCurrentAttribute > 1 Or blnRecordLabelsExist=False)) Then
						
						'Finally, add this piece of data to the 2D associative array
						FnCreateComplexTDAA.Add(strRecordLabel, listAttributeLabels.Item(intCurrentAttribute - If(blnRecordLabelsExist, 2, 1)), strCellValue)
						
					End If
					
			    Next
					
			Next
		
		End If		

	End Function
	
	
'==============================================================================================================
' FnCompareExcel
'
' 		Does a simple cell-by-cell value comparison for all cells within the first sheet of two Excel files.
'		An optional parameter is available to have the function output a "diff file" that contains yellow
'		highlighting wherever differences are found.
'
'		FnCompareExcel returns a string that explains where differences were found.  If more than 5 differences
'		were found, it will report the first 5 and then say that there are X more differences.
'
'		NOTE: The reason this function exists in conjunction with a corresponding _DoNotCallDirectly routine 
'		is because VB.NET is leaving the Excel.exe process lingering in the background after the routine is 
'		finished.  I have tried endlessly to overcome this by exiting gracefully and reseting variables to null, 
'		but nothing has worked consistently.  Hence, this function is making the call to its corresponding 
'		_DoNotCallDirectly routine, then immediately afterward it is calling SubDoGarbageCollect.  Because .NET
'		has gone through the corresponding _DoNotCallDirectly routine, the routine is officially out of scope 
'		and finished when SubDoGarbageCollect is invoked.  Hence, the garbage collection is free to clean up and
'		eradicate any variables which were created within the _DoNotCallDirectly routine.  This is what it takes
'		to finally free the file from memory and subsequently terminate the Excel.exe process.
'
' 		@strPathX:	The full path to the first of two files to compare
' 		@strPathY:	The full path to the second of two files to compare
' 		@strPathZ:	Optionally, the full path to the diff file
'==============================================================================================================

	Public Function FnCompareExcel(ByVal strPathX As String, ByVal strPathY As String, Optional ByVal strPathZ As String=Nothing) As String
		FnCompareExcel = FnCompareExcel_DoNotCallDirectly(strPathX, strPathY, strPathZ)
		SubDoGarbageCollect
	End Function
	
	Private Function FnCompareExcel_DoNotCallDirectly(ByVal strPathX As String, ByVal strPathY As String, ByVal strPathZ As String) As String

		'Default return
		Dim strReturn As String=Nothing
	
		'Objects
		Dim dicDataX, dicDataY As Dictionary (Of String, Object(,))
		Dim dicDiffs As New Dictionary(Of String, String())
		
		'Variables
		Dim intNumRowsX, intNumRowsY As Integer
		Dim intNumColsX, intNumColsY As Integer
		Dim strSheetX, strSheetY As String
		Dim strCellValueX, strCellValueY As String
		Dim intCounter As Integer
	
		'Read in both Excel files
		dicDataX = FnFetchExcelData(strPathX)
		dicDataY = FnFetchExcelData(strPathY)
		
		'Determine the names of the first sheets for each workbook
		strSheetX = (New List(Of String)(dicDataX.Keys))(0)
		strSheetY = (New List(Of String)(dicDataY.Keys))(0)
		
		'Determine number of rows and columns for each spreadsheet
		intNumRowsX = dicDataX.Item(strSheetX).GetUpperBound(0)
		intNumColsX = dicDataX.Item(strSheetX).GetUpperBound(1)
		intNumRowsY = dicDataY.Item(strSheetY).GetUpperBound(0)
		intNumColsY = dicDataY.Item(strSheetY).GetUpperBound(1)
		
		'While there are rows in X
 		For intCurrentRow As Integer = 1 To intNumRowsX
			
			'While there are columns in this X row
			For intCurrentCol As Integer = 1 To intNumColsX
				
				'Get the X cell value
				strCellValueX = dicDataX.Item(strSheetX)(intCurrentRow, intCurrentCol)	
				
				'Get the Y cell value
				If(intCurrentRow <= intNumRowsY AndAlso intCurrentCol <= intNumColsY) Then
					strCellValueY = dicDataY.Item(strSheetY)(intCurrentRow, intCurrentCol)
				Else
					strCellValueY = Nothing
				End If
				
				'If this cell is not available in Y OR the value at this cell differs between Excel files
				If Not String.Equals(strCellValueX, strCellValueY) Then 
											
					'Add this coordinate to a dictionary of string arrays
					dicDiffs.Add(FnNumToLetter(intCurrentCol) & intCurrentRow, {strCellValueX, strCellValueY})
					
				End If
					
			Next intCurrentCol
						
			'While there are more columns on this row for Y
			For intCurrentCol As Integer = intNumColsX+1 To intNumColsY
				
				'Get the Y cell value
				strCellValueY = dicDataY.Item(strSheetY)(intCurrentRow, intCurrentCol)
				
				'If there is value in this cell, add this coordinate to a dictionary of string arrays
				If Not String.IsNullOrEmpty(strCellValueY) Then				
					dicDiffs.Add(FnNumToLetter(intCurrentCol) & intCurrentRow, {Nothing, strCellValueY})
				End If
			
			Next intCurrentCol
				
		Next intCurrentRow
		
		'While there are more rows in Y than X
		For intCurrentRow As Integer = intNumRowsX+1 To intNumRowsY
			
			'While this Y row has columns 
			For intCurrentCol As Integer = 1 To intNumColsY
			
				'Get the Y cell value
				strCellValueY = dicDataY.Item(strSheetY)(intCurrentRow, intCurrentCol)
				
				'If there is value in this cell, add this coordinate to a dictionary of string arrays
				If Not String.IsNullOrEmpty(strCellValueY) Then
					dicDiffs.Add(FnNumToLetter(intCurrentCol) & intCurrentRow, {Nothing, strCellValueY})
				End If
				
			Next intCurrentCol
			
		Next intCurrentRow
		
		'If differences were found, form the return string which indicates such
		If(dicDiffs.Count > 0)
			
			'Loop through each difference
			For Each kvp As KeyValuePair(Of String, String()) In dicDiffs
			
				'Keep track of how many differences we're outputting
				intCounter = intCounter + 1
				
				'If this is the first difference, form the beginning of the string appropriately
				If intCounter = 1 Then
					strReturn = dicDiffs.Count & " difference" & If(dicDiffs.Count=1, "", "s") & " found between the spreadsheets."
				End If
				
				'Tack on a difference
				strReturn = strReturn & "  " & kvp.Key & ": [" & kvp.Value(0) & "] != [" & kvp.Value(1) & "]."
				
				'If we've already listed 5 differences and there are more to go, just say "# more differences not listed here."
				If intCounter >= 5 And dicDiffs.Count > 5 Then
					strReturn = strReturn & "  " & dicDiffs.Count-5 & " more not listed here."
					Exit For
				End If
				
			Next
			
			'If the user wants a diff file
			If Not strPathZ Is Nothing Then 
				
				'Concat instructions that he can see the differences in the diff file
				strReturn = strReturn & "  See the highlighted difference" & If(dicDiffs.Count=1, "", "s") & " inside " & strPathZ & "."
				
				'Open the second workbook
				Dim xl As New Excel.Application
				Dim wb As Excel.Workbook = xl.Workbooks.Open(strPathY)
				
				'Highlight every diff cells
				For Each strRange As String In FnConvertCellListToDelimitedStringListUnder256Chars(New List(Of String)(dicDiffs.Keys))
					wb.Sheets(1).Range(strRange).Interior.Color = 65535
				Next strRange
				
				'Close/Save the workbook as the Z file
				xl.DisplayAlerts = False
				wb.SaveAs(strPathZ)
				wb.Close()
				xl.Quit
				xl = Nothing
				
			End If
			
		End If
		
		Return strReturn
	
	End Function
	
	
'==============================================================================================================
' FnConvertCellListToDelimitedStringListUnder256Chars
'
' 		Takes a list of strings and concats them all together with a comma delimiter between each value.  Then,
'		it chops the giant string up into pieces as close to 255 characters as possible.
'
'		This is a specialized helper function that has only one real use.  In Excel, you can define a range such
'		as "A1,B8,A45,H4" and so on.  By putting multiple cells in a range, you can perform an action on all
'		contained cells simultaneously.  This cuts down on IO from .NET to Excel.  However, the Range function in
'		Excel is limited to 256 characters.  So, to minimize IO calls, I need to create several of these comma-
'		delimited strings for Excel to cleanly interact with many hundreds of cells simultaneously.
'
'		This function is used when highlighting hundreds of "diff cells" via the FnCompareExcel function.
'
' 		@strCellList:	The list of cells to combine and divide into chunks for 255 characters or less.
'==============================================================================================================

	Public Function FnConvertCellListToDelimitedStringListUnder256Chars(ByVal strCellList As List(Of String)) As List(Of String)
	
		'Instantiate return
		FnConvertCellListToDelimitedStringListUnder256Chars = New List(Of String)
		
		'Variables
		Dim intCellCounter As Integer=0
		Dim strBuilder As String=""
		
		'Do while we have more cells to add
		Do While intCellCounter < strCellList.Count
			
			'If the addition of a comma and the next cell doesn't put us at or over 256 characters, tack it onto the string
			If Len(strBuilder & strCellList.Item(intCellCounter))+1 < 255 Then		
				strBuilder = strBuilder & strCellList.Item(intCellCounter) & ","
				intCellCounter = intCellCounter + 1
				
			'But if the addition of a comma and the next cell is too much, then tack the builder onto the list
			Else 
				FnConvertCellListToDelimitedStringListUnder256Chars.Add(Left(strBuilder, Len(strBuilder)-1))
				strBuilder = Nothing
			End If
			
			'If there are no more cells, then tack on the final builder and be done
			If intCellCounter = strCellList.Count Then
				FnConvertCellListToDelimitedStringListUnder256Chars.Add(Left(strBuilder, Len(strBuilder)-1))
			End If
			
		Loop
		
	End Function
	

'==============================================================================================================
' FnFetchExcelData
'
' 		Retrieves 2D object arrays from the Excel cache.  If the cache doesn't contain the specified file, then
'		this function makes a call to load it into the cache.
'
'		Returns a dictionary keyed by worksheet names which map to 2D object arrays of corresponding cell values
'
'		@strPathToExcelFile		The full file path to an Excel file
'==============================================================================================================
	
	Public Function FnFetchExcelData(ByVal strPathToExcelFile As String) As Dictionary(Of String, Object(,))
	
		'If the file doesn't already exist in cache, go put it in the cache
		If Not dicExcelFileCache.ContainsKey(strPathToExcelFile) Then
			SubCacheExcelFile(strPathToExcelFile)
		End If
		
		'Read the data from the cache
		FnFetchExcelData = dicExcelFileCache.Item(strPathToExcelFile.ToLower)
		
	End Function
	

'==============================================================================================================
' SubCacheExcelFile
'
' 		Reads in entire Excel workbook in a single IO and puts it in this module's cache.  
'
'		The cache, dicExcelFileCache, is a persistent dictionary of Excel file paths pointing to a nested dictionary.
'		The nested dictionary is keyed by worksheet name and points to a 2D object array of cell values.
'
'		NOTE: The reason this function exists in conjunction with a corresponding _DoNotCallDirectly routine 
'		is because VB.NET is leaving the Excel.exe process lingering in the background after the routine is 
'		finished.  I have tried endlessly to overcome this by exiting gracefully and reseting variables to null, 
'		but nothing has worked consistently.  Hence, this function is making the call to its corresponding 
'		_DoNotCallDirectly routine, then immediately afterward it is calling SubDoGarbageCollect.  Because .NET
'		has gone through the corresponding _DoNotCallDirectly routine, the routine is officially out of scope 
'		and finished when SubDoGarbageCollect is invoked.  Hence, the garbage collection is free to clean up and
'		eradicate any variables which were created within the _DoNotCallDirectly routine.  This is what it takes
'		to finally free the file from memory and subsequently terminate the Excel.exe process.
'
' 		@strPathToExcelFile:	The full file path to the Excel file
' 		@blnOverwriteCache:		Overwrite the cached copy of the file if it already exists in the cache
'==============================================================================================================

	Public Sub SubCacheExcelFile(ByVal strPathToExcelFile As String, Optional ByVal blnOverwriteCache As Boolean=False)
		SubCacheExcelFile_DoNotCallDirectly(strPathToExcelFile, blnOverwriteCache)
		SubDoGarbageCollect
	End Sub

	Private Sub SubCacheExcelFile_DoNotCallDirectly(ByVal strPathToExcelFile As String, ByVal blnOverwriteCache As Boolean)
	
		'Immediately lowercase the path string
		strPathToExcelFile = strPathToExcelFile.ToLower
		
		'If we would like to overwrite existing cache or we don't have a copy of this file in the cache
		If blnOverwriteCache OrElse Not dicExcelFileCache.ContainsKey(strPathToExcelFile) Then
	
			'Objects
			Dim xl As New Excel.Application
			Dim wb As Excel.Workbook = xl.Workbooks.Open(strPathToExcelFile,, True) 'Read only
			
		    'Variables		
		    Dim ws As Worksheet
			Dim r As Range
			Dim obj2DArr(,) As Object
			
			'Create a clean slate in which to store the forthcoming 2D obj array
			If dicExcelFileCache.ContainsKey(strPathToExcelFile) Then
				dicExcelFileCache.Item(strPathToExcelFile).Clear
			Else
				dicExcelFileCache.Add(strPathToExcelFile, New Dictionary(Of String, Object(,)))
			End If
			
			'Loop once for every worksheet in the workbook
			For i As Integer = 1 To wb.Sheets.Count
				
			    'Get this sheet
			    ws = wb.Sheets(i)
				
				'Assuming the sheet isn't empty, add the entire used range to the dictionary of 2D object arrays
				If Not (ws.UsedRange.Rows.Count=1 And ws.UsedRange.Columns.Count=1 And ws.Cells(1,1).Value="") Then 
				
					'Get the used range
					r = ws.Range(ws.Cells(1,1), FnGetLastUsedCell(ws))
				
					'Create a 2D object array from the worksheet
			    	obj2DArr = r.Value(XlRangeValueDataType.xlRangeValueDefault)
				
					'Add this worksheet to the cache
					dicExcelFileCache.Item(strPathToExcelFile).Add(ws.Name.ToLower, obj2DArr)
					
				End If
				
			Next i
			
			'Close components
		    wb.Close()
		    xl.Quit()
			
		End If
	
	End Sub

	
'==============================================================================================================
' FnUpdateObjectReferences
'
'		This function is called every time the user calls the LoadObjects command in Xscript.  It is responsible
'		for updating object names in all active, inactive, and library test files.
'
'		Conceptually, this function is simple.  If any object within the referenced object workbook has a value
'		in its "New coded name" column, it means the user wishes to do a batch update across all test files.  So,
'		this function scans each file and does the update where necessary.  After combing through each file, it
'		moves the new name from the "New coded name" column into the "Coded name" column in the objects workbook.
'
'		However, there are a few other things to consider:
'
'		[1] Updates are not allowed to happen if any of the object or test files are write-locked.  To meet this
'			requirement, each directory is scanned for hidden xlsx files that start with the ~ character.  If it's
'			present, then the function tries to delete it.  If it can't delete it, then the file is legitimately
'			write-locked and the function aborts.
'
'		[2] The Default.xlsm workbook needs to be scanned for updates, too.  To do this, I have fudge the steps
'			in the test file so that it thinks there is a LoadObjects(Default.xlsm) instruction present.  That way,
'			this routine is called for Default.xlsm by the Interpreter.
'
'		[3] Because duplicates are allowed to exist in object worksheets, the LoadObjects command will end up
'			overwriting "older" objects with the xpath of "newer" objects.  This is not strictly a rule however as
'			the user may specify "trumping worksheets" when they invoke the LoadObjects command.  Hence, when we
'			find the LoadObjects command in a test, we have to execute it ourselves to build the object map just
'			as it would be during runtime.  Only then can we know what the user's object map looks like so that we
'			may properly run Excel's Replace function.
'
'		NOTE: The reason this function exists in conjunction with a corresponding _DoNotCallDirectly routine 
'		is because VB.NET is leaving the Excel.exe process lingering in the background after the routine is 
'		finished.  I have tried endlessly to overcome this by exiting gracefully and reseting variables to null, 
'		but nothing has worked consistently.  Hence, this function is making the call to its corresponding 
'		_DoNotCallDirectly routine, then immediately afterward it is calling SubDoGarbageCollect.  Because .NET
'		has gone through the corresponding _DoNotCallDirectly routine, the routine is officially out of scope 
'		and finished when SubDoGarbageCollect is invoked.  Hence, the garbage collection is free to clean up and
'		eradicate any variables which were created within the _DoNotCallDirectly routine.  This is what it takes
'		to finally free the file from memory and subsequently terminate the Excel.exe process.
'
' 		@strObjectFileName:	The filename of the objects file the user is trying to load
' 		@test:				The currently executing Test object
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================
	
	Public Function FnUpdateObjectReferences(ByVal strObjectFileName As String, ByRef test As Test, ByRef dicConstants As Dictionary(Of String, String)) As String
		FnUpdateObjectReferences = FnUpdateObjectReferences_DoNotCallDirectly(strObjectFileName, test, dicConstants)
		SubDoGarbageCollect
	End Function

	Private Function FnUpdateObjectReferences_DoNotCallDirectly(ByVal strObjectFileName As String, ByRef test As Test, ByRef dicConstants As Dictionary(Of String, String)) As String
												
		'Declare variables
		Dim intNumRows, intNumColumns, intTemp, intNewNameColNum, intCurrentNameColNum As Integer
		Dim blnOverhaulRequested, blnAllFilesAvailableForWriting, blnCopyErrorOccurred, blnFailTestIfLocked As Boolean
		Dim strTestFilesList, strStepList, strTempList As New List(Of String)
		Dim objParamList As New List(Of Object)		
		Dim strTemp, strDateTime, strCmd, strReturn, strObjectSheetFilePath As String
		Dim strTestFileDirs As String()
		Dim strMapDic, strTempMapDic, strUpdatesDic, strObjectFileRowNumUpdatesDic As New Dictionary(Of String, String)
		Dim xl As Excel.Application
		Dim wb As Excel.Workbook
		Dim ws As Excel.Worksheet
											
		'Set variables
		strReturn = Nothing
		blnAllFilesAvailableForWriting = True
		strTemp = Nothing
		strCmd = Nothing
		strObjectSheetFilePath = dicConstants.Item("ObjectsFolder") & "\" & strObjectFileName
		blnFailTestIfLocked = If(String.Compare(dicConstants.Item("FailTestObjectOverhaulOnFileWriteLocked"), "True", True)=0, True, False)
		strTestFileDirs = {dicConstants.Item("TestsLocation"), dicConstants.Item("InactiveTestsPath"), dicConstants.Item("LibraryLocation")}
		
		'Put the file in the cache if it isn't already there
		SubCacheExcelFile(strObjectSheetFilePath)
					
		'Loop through every sheet in this cached objects file
		For Each sheetToObj2DArrDic As KeyValuePair(Of String, Object(,)) In dicExcelFileCache.Item(strObjectSheetFilePath.ToLower)
		
			'print("Reading objects file sheet: " & sheetToObj2DArrDic.Key)
			
			intNumRows = sheetToObj2DArrDic.Value.GetUpperBound(0)
			intNumColumns = sheetToObj2DArrDic.Value.GetUpperBound(1)
			
			'Loop once for every number of columns present
			For intCurrentColumnIndex As Integer = 1 To intNumColumns
				
				'If we have found the new coded name column
				If String.Compare(sheetToObj2DArrDic.Value(1, intCurrentColumnIndex), dicConstants.Item("ObjectNewCodedNameColumnHeader"), True)=0 Then
				
					'Loop through the new coded name column from top to bottom
					For intCurrentRowIndex As Integer = 2 To intNumRows
						
						'If there is a value in this row at any point
						If Not String.IsNullOrEmpty(sheetToObj2DArrDic.Value(intCurrentRowIndex, intCurrentColumnIndex)) AndAlso Len(sheetToObj2DArrDic.Value(intCurrentRowIndex, intCurrentColumnIndex).Trim) > 0 Then
						
							'Set a boolean to indicate that a reference overhaul is requested
							blnOverhaulRequested = True
							
							'Abort loop
							If blnOverhaulRequested Then Exit For
								
						End If
						
					Next intCurrentRowIndex
					
				End If
				
				'Abort loop
				If blnOverhaulRequested Then Exit For
							
			Next intCurrentColumnIndex

			'Abort loop
			If blnOverhaulRequested Then Exit For
							
		Next sheetToObj2DArrDic
					
		'If a reference overhaul is requested
		If blnOverhaulRequested Then
			
			'If the object worksheet that is referenced in this LoadObjects command is write-locked
			If File.Exists(Path.GetDirectoryName(strObjectSheetFilePath) & "\~$" & Path.GetFileName(strObjectSheetFilePath)) Then
				
				Try
					
					'Try to delete the temp file - maybe it's just been lingering for no reason
					File.Delete(Path.GetDirectoryName(strObjectSheetFilePath) & "\~$" & Path.GetFileName(strObjectSheetFilePath))
					
				Catch eIO As IOException
					
					blnAllFilesAvailableForWriting = False
					
					'If deletion failed, this will tell the user that there's a write-lock
					strReturn = "Cannot perform object reference updates because '" & Path.GetFileName(strObjectSheetFilePath) & "' is currently write-locked." & _
								"Please close the file or ensure nobody else has it opened."
				End Try
			
			End If

			'If there is no write-lock on the objects file
			If blnAllFilesAvailableForWriting Then
				
				'Loop through each strTestFileDirs
				For Each strDir As String In strTestFileDirs
					
					'Get a list of visible test files in this directory
					strTestFilesList = FnGetFilePaths(strDir, {"xls", "xlsx", "xlsm"})
					
					'Loop through every visible test file in this directory
					For Each strTestFile As String In strTestFilesList
					
						'Create a string to represent a potentially hidden ~$ and corresponding file
						strTemp = Path.GetDirectoryName(strTestFile) & "\~$" & Path.GetFileName(strTestFile)
					
						'If this test file has a corresponding file prefixed with ~$ that is not in our list, it's likely a hidden write-lock file
						If File.Exists(strTemp) AndAlso strTestFilesList.Contains(strTemp)=False Then
							
							'Try to delete it
							Try
								
								File.Delete(strTemp)
							
							'Catch the failure to delete it
							Catch eIO As IOException
								
								blnAllFilesAvailableForWriting = False
							
								'If the user wants to fail this test when any one test file is write-locked, update the error msg
								If blnFailTestIfLocked Then
									strReturn = If(Len(strReturn)>0, "  ", "") & "Cannot update object names across all test files because '" & _
															   strTestFile & "' is currently write-locked by another user.  Please close the file and restart Xscript." 
								
								'If the user wants to proceed even when a write-lock is encountered, set a warning msg
								Else
									test.strWarningList.Add("Object reference update did not occur for any test file because there is a write-lock on '" & strTestFile & "'.  " & _
															"Please close the file and the object reference update will occur on the next Xsript run.")
								End If
								
							End Try
							
						End If
						
					Next strTestFile
								
				Next strDir
			
			End If
			
			'If all files are available for writing
			If blnAllFilesAvailableForWriting Then 
				
				'Get the datetime in the form YYYYMMDD_HHMMSS
				strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmmss")
				
				'Create the timestamped temp directory and timestamped backup directory
				Directory.CreateDirectory(dicConstants.Item("BackupLocation") & "\" & strDateTime)
				Directory.CreateDirectory(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime)
				
				'Open the xl application here so we only have to do it once
				xl = New Excel.Application
				xl.DisplayAlerts = False
			
				'Loop through each strTestFileDirs
				For Each strDir As String In strTestFileDirs
					
					'Loop for each unhidden test file
					For Each strTestFile As String In FnGetFilePaths(strDir, {"xls", "xlsx", "xlsm"})
					
						'Clear test-specific objects
						strMapDic.Clear
						strUpdatesDic.Clear	
						strStepList.Clear
						
						'Create a strStepList but insert "LoadObjects('Default.xlsm')" as the first step
						strStepList.Add("LoadObjects(""" & dicConstants.Item("DefaultObjectsWorkbook") & """)")
								
						'Add the steps in this test file to the strStepList
						strStepList.AddRange(FnLoadSimpleList(strTestFile, dicConstants.Item("InstructionsColumnHeader")))
						
						'Loop through each step of the test
						For intStepIndex As Integer = 0 To strStepList.Count-1
							
							'Parse the instruction into strCmd and objParamList
							strTemp = FnParseInstruction(strStepList.Item(intStepIndex), strCmd, objParamList)
							
							'Drop all the quotation marks
							For j As Integer = 0 To objParamList.Count-1
								objParamList.Item(j) = objParamList.Item(j).Replace("""", "")
							Next j
							
							'If this is the LoadObjects command (no matter the param) AND [we already loaded the relevant objects file OR we're doing that now]
							If String.Compare(strCmd, "LoadObjects", True)=0 AndAlso (strMapDic.Count > 0 OrElse String.Compare(Path.GetFileName(strObjectSheetFilePath), objParamList.Item(0), True)=0) Then
							
								'Build a temporary map of old names pointing to new names.  Use FnLoadSimpleDictionary command as if there were a real test run. 
								strTempMapDic = FnLoadSimpleDictionary(strReturn, dicConstants("ObjectsFolder") & "\" & objParamList.Item(0), dicConstants.Item("ObjectNameColumnHeader"), _
																	   dicConstants.Item("ObjectNewCodedNameColumnHeader"), True, If(objParamList.Count>1, objParamList.GetRange(1, objParamList.Count-1).ToArray, Nothing))
								
								'If this LoadObjects command represents the Object sheet that Xscript is trying to truly update references on at this moment, strip out all pairs that don't have a 'new coded name' value
								If String.Compare(Path.GetFileName(strObjectSheetFilePath), objParamList.Item(0), True)=0 Then 
								
									'Merge the temp map into the strMapDic in case the user has now or previously called LoadObjects with multiple parameters
									For Each strTempMapPair As KeyValuePair(Of String, String) In strTempMapDic
										If strMapDic.ContainsKey(strTempMapPair.Key) Then
											strMapDic.Item(strTempMapPair.Key) = strTempMapPair.Value
										Else
											strMapDic.Add(strTempMapPair.Key, strTempMapPair.Value)
										End If
									Next strTempMapPair	
									
									'Get a list of the keys for strMapDic							
									strTempList = New List(Of String)(strMapDic.Keys)
								
									'Remove any pairs from strMapDic that do have actually value in their 'new coded name' column
									For k As Integer=0 To strTempList.Count-1
										If String.IsNullOrEmpty(strMapDic.Item(strTempList.Item(k))) Then
											strMapDic.Remove(strTempList.Item(k))
										End If
									Next k
									
									'Order the map so it is keyed by longest values first
									strMapDic = FnReverseSortDictionaryKeys(strMapDic)
									
								'If this is the LoadObjects command for an unrelated Objects sheet, just use it to drop items from the strMapDic
								Else
									
									For Each kvpStrTempMapDic As KeyValuePair(Of String, String) In strTempMapDic
										If strMapDic.ContainsKey(kvpStrTempMapDic.Key) Then strMapDic.Remove(kvpStrTempMapDic.Key)
									Next kvpStrTempMapDic
									
								End If
									
							'Else, this isn't the LoadObjects command, but if we already have a map, then proceed
							Else If strMapDic.Count > 0
								
								'Loop through the map
								For Each kvpMapping As KeyValuePair(Of String, String) In strMapDic
									
									'Search for the old name within this step
									intTemp = Instr(strStepList.Item(intStepIndex).ToLower, kvpMapping.Key.ToLower)-1
									
									'If the old name is present within this instruction somewhere
									If intTemp > 0 Then
										
										'Run a case insensitive replacement on instruction to swap out the old new with the new name, then add it to the updates dictionary keyed the intStepIndex+1 (the one is to offset the fraudulent LoadObjects("Default.xlsm") instruction
										strUpdatesDic.Add(intStepIndex+1, Regex.Replace(strStepList.Item(intStepIndex), kvpMapping.Key, kvpMapping.Value, RegexOptions.IgnoreCase))
										
									End If
									
								Next kvpMapping
							
							End If
										
						Next intStepIndex
										
						'If strUpdatesDic has updates to make on this test file
						If strUpdatesDic.Count > 0 Then
						
							'Figure out which column represents the 'Steps' column (it's likely column 'A')
							For Each kvpSheetToData As KeyValuePair(Of String, Object(,)) In dicExcelFileCache.Item(strTestFile.ToLower)
								intNumColumns = kvpSheetToData.Value.GetUpperBound(1)
								For intCurrentColumn=1 To intNumColumns
									If String.Compare(dicConstants.Item("InstructionsColumnHeader"), kvpSheetToData.Value(1, intCurrentColumn), True)=0 Then
										intTemp=intCurrentColumn
										Exit For
									End If
								Next intCurrentColumn
								Exit For
							Next kvpSheetToData	
							
							'Create the corresponding temp directory and backup directory if they don't already exist
							If Not Directory.Exists(dicConstants.Item("BackupLocation") & "\" & strDateTime & "\" & Path.GetFileName(strDir)) Then
								Directory.CreateDirectory(dicConstants.Item("BackupLocation") & "\" & strDateTime & "\" & Path.GetFileName(strDir))
							End If
							If Not Directory.Exists(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir)) Then
								Directory.CreateDirectory(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir))
							End If
							
							'Copy this test to the temp and backup directories
							File.Copy(strTestFile, dicConstants.Item("BackupLocation") & "\" & strDateTime & "\" & Path.GetFileName(strDir) & "\" & Path.GetFileName(strTestFile), True)
							File.Copy(strTestFile, dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir) & "\" & Path.GetFileName(strTestFile), True)
							
							'Open the test file and first sheet from the temp folder
							wb = xl.Workbooks.Open(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir) & "\" & Path.GetFileName(strTestFile))
							ws = wb.Sheets(1)
														
							'Loop through strUpdatesDic and make updates to the proper cell
							For Each kvpRowNumToValueUpdates As KeyValuePair(Of String, String) In strUpdatesDic							
								ws.Cells(kvpRowNumToValueUpdates.Key, intTemp).Value = kvpRowNumToValueUpdates.Value
							Next kvpRowNumToValueUpdates
							
							'Save the file and close the workbook
							wb.Save()
							wb.Close()
							
							'Delete the cache reference from dicExcelFileCache
							dicExcelFileCache.Remove(strTestFile.ToLower)
							
						End If
							
					Next strTestFile
				
				Next strDir
						
				'Loop through each strTestFileDirs
				For Each strDir As String In strTestFileDirs
					
					'If there is a folder where I'd expect to find the updated version of files within this dir
					If Directory.Exists(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir)) Then
						
						'For each file in this specific Temp\strDateTime subfolder
						For Each strFileNameFullPath As String In FnGetFilePaths(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strDir), {"xls", "xlsx", "xlsm"})
						
							Try
						
								'Figure out the updated file's original counterpart filepath
								If String.Compare(Path.GetFileName(dicConstants.Item("TestsLocation")), Path.GetFileName(Path.GetDirectoryName(strFileNameFullPath)), True)=0 Then
									strTemp = dicConstants.Item("TestsLocation") & "\" & Path.GetFileName(strFileNameFullPath)
								Else If String.Compare(Path.GetFileName(dicConstants.Item("InactiveTestsPath")), Path.GetFileName(Path.GetDirectoryName(strFileNameFullPath)), True)=0 Then
									strTemp = dicConstants.Item("InactiveTestsPath") & "\" & Path.GetFileName(strFileNameFullPath)
								Else
									strTemp = dicConstants.Item("LibraryLocation") & "\" & Path.GetFileName(strFileNameFullPath)
								End If
								
								'Copy the folder onto its original
								File.Copy(strFileNameFullPath, strTemp, True)
								
								'Delete the update file
								File.Delete(strFileNameFullPath)
								
							Catch ex As Exception
								
								'Set a warning saying that someone likely opened the file since the LoadObjects command began executing.  You can find the updated file In insertUpdatedDirHere.
								test.strWarningList.Add("Object name updates did not occur for '" & strTemp & "' because the file spontaneously became write-locked.  It was not write-locked " & _
														"a moment ago, so someone likely just opened it.  Xscript has instead copied the file and performed the updates to it in another location. " & _
														"You can find the copied and updated file at '" & strFileNameFullPath & "'.")
								
								'Set boolean saying that error occurred
								blnCopyErrorOccurred = True
								
							End Try
								
						Next strFileNameFullPath
					
					End If
						
				Next strDir
				
				'Copy the objects worksheet to the temp location and backup
				File.Copy(strObjectSheetFilePath, dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strObjectSheetFilePath), True)
				File.Copy(strObjectSheetFilePath, dicConstants.Item("BackupLocation") & "\" & strDateTime & "\" & Path.GetFileName(strObjectSheetFilePath), True)
				
				'Open the copied objects worksheet
				wb = xl.Workbooks.Open(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strObjectSheetFilePath))
								
				'Loop through each sheet of the cached object workbook
				For Each sheetToObj2DArrDic As KeyValuePair(Of String, Object(,)) In dicExcelFileCache.Item(strObjectSheetFilePath.ToLower)
				
					intNumRows = sheetToObj2DArrDic.Value.GetUpperBound(0)
					intNumColumns = sheetToObj2DArrDic.Value.GetUpperBound(1)
					intCurrentNameColNum = 0
					intNewNameColNum = 0
					
					'Loop across the top row of this sheet to find the column numbers for the current and new coded names
					For intCurrentColumnIndex As Integer = 1 To intNumColumns
						If String.Compare(sheetToObj2DArrDic.Value(1, intCurrentColumnIndex), dicConstants.Item("ObjectNameColumnHeader"), True)=0 Then
							intCurrentNameColNum = intCurrentColumnIndex
						Else If String.Compare(sheetToObj2DArrDic.Value(1, intCurrentColumnIndex), dicConstants.Item("ObjectNewCodedNameColumnHeader"), True)=0 Then
							intNewNameColNum = intCurrentColumnIndex
						End If						
					Next intCurrentColumnIndex
					
					'If we found the old and new name columns in this sheet
					If intCurrentNameColNum>0 AndAlso intNewNameColNum>0 Then
					
						'Set the worksheet
						ws = wb.Sheets(sheetToObj2DArrDic.Key)
												
						'Loop through the 'new name' column of the cache.  If there is value in it, copy it over to the current col name, then delete it.
						For intCurrentRowNum As Integer=2 To intNumRows
							If Not String.IsNullOrEmpty(ws.Cells(intCurrentRowNum, intNewNameColNum).Value) Then
								ws.Cells(intCurrentRowNum, intCurrentNameColNum).Value = ws.Cells(intCurrentRowNum, intNewNameColNum).Value
								ws.Cells(intCurrentRowNum, intNewNameColNum).Value = Nothing
							End If
						Next intCurrentRowNum
						
					End If
					
				Next sheetToObj2DArrDic
				
				'Delete cached copy of this file
				dicExcelFileCache.Remove(strObjectSheetFilePath.ToLower)
	
				'Close the xl application
				wb.Save
				wb.Close
				xl.Quit
				
				'Overwrite file now
				Try					
					
					File.Copy(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strObjectSheetFilePath), strObjectSheetFilePath, True)
					File.Delete(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime & "\" & Path.GetFileName(strObjectSheetFilePath))
					
				Catch ex As Exception
					blnCopyErrorOccurred=True
				End Try
						
				'If no error occurred, delete the Updated folder
				If blnCopyErrorOccurred = False Then Directory.Delete(dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\" & strDateTime, True)
				
				'Update the steps list in the test object
				test.strStepsList = FnLoadSimpleList(test.strExcelFilePath, dicConstants.Item("InstructionsColumnHeader"), True, True)
			
			'Ends the block that ensures all files are available for writing before proceeding with overhaul
			End If 
		
		'Ends the block that indicates the user wants to perform a reference overhaul
		End If
		
		Return strReturn
							
	End Function


'==============================================================================================================
' FnGetLastUsedCell
'
'		Given an Excel worksheet, this function will return the bottom-right most used cell.  This is very
'		different from Excel's "UsedRange" constant because UsedRange takes into account cells which have no
'		data but do actually have formatting applied to them.  I don't care about empty cells with formatting.
'		This function only returns the bottom-right most cell with actual content.
'
' 		@ws:	The Excel Worksheet object to examine
'==============================================================================================================
	
	Private Function FnGetLastUsedCell(ByRef ws As Worksheet) As Range
		
		Dim intLastUsedRow, intLastUsedCol As Integer
		
		intLastUsedRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Row
		intLastUsedCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Column
		
		FnGetLastUsedCell = ws.Cells(intLastUsedRow, intLastUsedCol)
		
	End Function
	
End Module
