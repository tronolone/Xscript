'Add reference: Excel 2007 Interop DLL
'Add reference: Xscript Actions
'Add reference: Xscript Classes

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Graphics
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Module ResultsModule

	Dim _desktop As Desktop = Agent.Desktop
	Dim blnAlreadyInvokedOnce As Boolean = False
	
'==============================================================================================================
' SubReportResultToExcel
'
' 		Outputs a single Test's result to a highly-formatted, easy-to-absorb Excel spreadsheet.  This routine
'		is called once for every single Test object.  Previously, we were executing all tests, then sending a
'		collection of test results to this routine.  However, if Silk died/hung at 95% progress through Xscript, 
'		then e would not know the fate of any of the tests, so that was modified.
'
'		A test result will always be added to the 'B' column on the Results.xlsm spreadsheet.  If the cell at 
'		'B1' does not share the same start date as the current test, a new 'B' column will be inserted and used.
'
'		Cells B2 through B# will contain the literal result of the test.  The color of the cell will be green,
'		orange, or red depending if it passed, bitmap failed, or normal failed.  The colors are generated indirectly
'		via conditional formatting.  I chose to do it this way so that sorting the entire spreadsheet would not be
'		a giant hassle full of color and border resassignment for every single cell.  Also, for bitmap and normal
'		failures, right-clicking on the cell will open a screenshot or a directory of screenshots for the user to examine.
'
' 		@strFilePath:				The path to the Results spreadsheet to populate
'		@test:						The test in which we want to report results
'		@dtCurrentDateTime:			Date/time that the suite of tests were started
'		@blnIncludeStackTrace:		If TRUE, any available stacktrace will be included in the results output
'		@strScreenshotFolderPath:	Path to the screenshots folder
'		@strImageFileExtension:		File extensions utilized by the screenshot function earlier
'
'		NOTE: You must import the Microsoft.Office.Interop.Excel.DLL file through the properties tab for this 
'		to work.  You cannot Do this directly.  You must go To START > RUN And enter the path "C:\windows\assembly\gac_msil".
'		From there, browse To 'Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c'.  Copy 
'		'Microsoft.Office.Interop.Excel.dll' to a safe place that everyone can access and won't be moved.  If 
'		you cannot find it there, it's probably because you don't have Microsoft Office installed.  If you still
'		can't find it, install the "Office Primary Interop Assemblies" package from Microsoft's website.  Now, 
'		import the assembly reference within Silk (via Properties tab) and use the 'Browse' feature to find the file.
'		In your code, you must have 'Imports Microsoft.Office.Interop' and 'Imports Microsoft.Office.Interop.Excel' 
'		at the top.
'
'		NOTE: There are a TON Of constants you'll need that look like 'xlValues' or 'xlCenter'.  You need To find
'		what class they're a member of and write them as 'XlFindLookIn.xlValues' or similar.  Use the list Of 
'		Enumerations found here: http://www.datapigtechnologies.com/downloads/Excel_Enumerations.txt.  If the 
'		above page is gone, contact Chuck (tronolone@gmail.com) as he retained a copy in his gmail account as 
'		'enumerations'.
'==============================================================================================================
	
	Public Sub SubReportResultToExcel(ByVal strFilePath As String, _
									  ByRef test As Test, _
									  ByVal dtCurrentDateTime As DateTime, _
									  ByVal blnIncludeStackTrace As Boolean, _
									  Optional ByVal strScreenshotFolderPath As String = "", _
									  Optional ByVal strImageFileExtension As String = "png")
		
		'Ensure the file is available for writing
		Dim fInfo As New FileInfo(strFilePath)	
		fInfo.IsReadOnly = False			
		
		'Declare objects
		Dim app As New Excel.Application
		app.DisplayAlerts = False
		Dim wb As Excel.Workbook = app.Workbooks.Open(strFilePath,,,,,,,,,,False)		
		Dim ws As Excel.Worksheet = wb.Worksheets(1)
		
		'Declare variables
		Dim r As range = Nothing
		Dim intUseThisRow As Integer = 2
		Dim intLastUsedRow As Integer = 1
		Dim intLastUsedCol As Integer = 1
				
		'Set final values
		test.SubBuildResult(blnIncludeStackTrace)
		
		'Figure out the last used row and column integers.  These are useful for everything going forward.
		r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False)
		If Not r Is Nothing Then
			intLastUsedRow = r.Row
			intLastUsedCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByColumns, SearchDirection:=XlSearchDirection.xlPrevious, MatchCase:=False).Column
		End If
		
		'If we haven't already written to the results file during this Xscript run
		If blnAlreadyInvokedOnce=False Then
			
			'Set a boolean so subsequent calls to this method know that a new column for this run has already been added
			blnAlreadyInvokedOnce = True
		
			'Add a new column to the right of the first column
			ws.Columns(2).Insert(Shift:=XlDirection.xlToRight, CopyOrigin:=XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
			intLastUsedCol = intLastUsedCol + 1
			
			'Add a space to every new cell that was just created to prevent text bleeding
			ws.Range("B2").Value = " "
			If intLastUsedRow > 2 Then ws.Range("B2").AutoFill(ws.Range("B2:B" & intLastUsedRow), XlAutoFillType.xlFillDefault)
		
			'If the named range doesn't already exist, add it 
			If(ws.Names.Count = 0) Then
				ws.Names.Add(Name:="ImageFolderPath", RefersToR1C1:="=R1C1")
			End If
		
			'Set the comment of the named range
			ws.Names.Item("ImageFolderPath").Comment = strScreenshotFolderPath
					
			'Insert date/time into the top row of the first unused column and style appropriately
			With ws.Cells(1, 2)			
				.Value = dtCurrentDateTime.ToString()
				.NumberFormat = "mm/dd/yyyy hh:mm:ss"
				.HorizontalAlignment = Constants.xlCenter
			End With
						
			'Size the column appropriately
			ws.Columns(2).ColumnWidth = 18
						
		   'Wipe out all conditional formatting
		    ws.Cells.FormatConditions.Delete
		    			
		    'Put the mouse in the first cell or else your formulas below get massively warped
		    ws.Range("A1").Select
		    			
		    'Header timestamps
		    r = ws.Range("$B$1:$XFD$1")
			r.FormatConditions.Add(Type:=XlFormatConditionType.xlCellValue, Operator:=XlFormatConditionOperator.xlGreater, Formula1:="=36526")
		    With r.FormatConditions(1)
		        .Font.Bold = True
		        .Font.ThemeColor = XlThemeColor.xlThemeColorDark1
		        .Font.TintAndShade = 0
		        .Interior.ThemeColor = XlThemeColor.xlThemeColorLight1
		    End With
		    			
		    'Fails
		    r = ws.Range("$B$2:$XFD$1048576")
			r.FormatConditions.Add(Type:=XlFormatConditionType.xlTextString, String:="Failure at step #", TextOperator:=XlContainsOperator.xlContains)
		    With r.FormatConditions(1)
		        .Borders(Constants.xlLeft).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 13311
		        .Borders(Constants.xlRight).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 13311
		        .Borders(Constants.xlTop).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 13311
		        .Borders(Constants.xlBottom).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 13311
		    End With
			
		    'Bitmaps fails
			r.FormatConditions.Add(Type:=XlFormatConditionType.xlTextString, String:="Bitmap check failed", TextOperator:=XlContainsOperator.xlContains)
		    With r.FormatConditions(2)
		        .Borders(Constants.xlLeft).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 2992895
		        .Borders(Constants.xlRight).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 2992895
		        .Borders(Constants.xlTop).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 2992895
		        .Borders(Constants.xlBottom).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 2992895
		    End With
			
		    'Passes
			r.FormatConditions.Add(Type:=XlFormatConditionType.xlTextString, String:="PASS", TextOperator:=XlContainsOperator.xlContains)
		    With r.FormatConditions(3)
		        .Borders(Constants.xlLeft).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 52377
		        .Borders(Constants.xlRight).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 52377
		        .Borders(Constants.xlTop).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 52377
		        .Borders(Constants.xlBottom).LineStyle = XlLineStyle.xlContinuous
		        .Interior.Pattern = 1
		        .Interior.Color = 52377
		    End With
		
		End If

		'Figure out if our test name already appears in the first column somewhere
		r = ws.Range("A:A").Find(What:=test.strTestName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, SearchOrder:=XlSearchOrder.xlByRows)
		
		'If the test name doesn't already appear in the first column
		If r Is Nothing Then
						
			'Insert the new row
			ws.Rows(2).Insert(Shift:=XlDirection.xlDown, CopyOrigin:=XlInsertFormatOrigin.xlFormatFromRightOrBelow)
			intLastUsedRow = intLastUsedRow + 1
			
			'Write a space character to cell C2.  This prevents text bleeding from B2 into C2.
			ws.Cells(2,3).Value = " "
								
			'Write the name of this test into A2
			ws.Cells(2, 1).Value = test.strTestName
					
			'Expand the column to fit all text without shrinking
			ws.Range("A:A").EntireColumn.AutoFit
			
			'Set the range to to use for this output
			intUseThisRow = 2
		
		'Otherwise, we found the pre-existing test name, so just reference the row in which is already exists
		Else
			intUseThisRow = r.Row
		End If
		
		'Within the result cell, write the message, apply centering, and shrink text to fit
		With ws.Cells(intUseThisRow, 2)
			.Value =  test.strFinalMsg
			.HorizontalAlignment = Constants.xlCenter
			.ShrinkToFit = True
		End With
		
		'If there are at least 3 rows total in this worksheet, then we need to run a sort routine
		If intLastUsedRow >= 3 Then 
			
			'Add a new helper column at C
		    ws.Columns("C:C").Insert(Shift:=XlDirection.xlToRight, CopyOrigin:=XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
		    
		    'Put a forumula in C2
		    ws.Range("C2").FormulaR1C1 = "=IF(RC[-1]="" "", ""Z"", ""A"")"
		    
		    'Autofill every cell below it with the same formula
		    ws.Range("C2").AutoFill(Destination:=ws.Range("C2:C" & intLastUsedRow), Type:=XlAutoFillType.xlFillDefault)
		        
		    'Clear the sort options
		    ws.Sort.SortFields.Clear
		    
		    'Add a sort options
			ws.Sort.SortFields.Add(Key:=ws.Range("C2:C" & intLastUsedRow), SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal)
		    ws.Sort.SortFields.Add(Key:=ws.Range("A2:A" & intLastUsedRow), SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal)
		        
		    'Set the sort switches and run it
			With ws.Sort
		        .SetRange(ws.Range("A2:" & FnNumToLetter(intLastUsedCol+1) & intLastUsedRow))
		        .Header = XlYesNoGuess.xlNo
		        .MatchCase = False
		        .Orientation = Constants.xlTopToBottom
		        .SortMethod = XlSortMethod.xlPinYin
		        .Apply
		    End With
		    
		    'Remove the helper column
		    ws.Columns("C:C").Delete(Shift:=XlDirection.xlToLeft)
			
		End If
		
		'Make sure the row heights are all set propertly because the stacktrace makes things expand wildly
		ws.UsedRange.RowHeight = 15
		
		'And unselect everything please
		ws.Range("A1").Select
			
		'Exit gracefully
		wb.SaveAs(strFilePath)
		app.Quit
		app = Nothing
		fInfo.IsReadOnly = True
		
	End Sub
	
		
'==============================================================================================================
' SubTrimBackScreenshots
'
' 		Recursively deletes files and empty folders from the specified directory and its subdirectories until 
'		they number less than the specified limit.  Deletion proceeds in alphabetical order from A to Z.
'
' 		@strDirectory:			The directory containing files/folders to delete
' 		@intNumFilesAllowed:	The maximum number of files allowed in the specified directory
'==============================================================================================================

	Public Sub SubTrimBackScreenshots(ByVal strDirectory As String, ByVal intNumFilesAllowed As Integer)
		
		'Objects
		Dim listOfFiles As New List(Of String)
		
		'Set objects
		listOfFiles = FnGetAllFilePathsRecursively(strDirectory)
		
		'If we have more than the allowed number of files	
		If(listOfFiles.Count > intNumFilesAllowed) Then
			
			'Sort the list in alphabetical order.  This puts the oldest files at the top.
			listOfFiles.Sort()
			
			Try
			
				'Delete files
				For i = 0 To listOfFiles.Count - intNumFilesAllowed - 1
					File.Delete(listOfFiles.Item(i))
				Next
				
				'Delete empty folders
				SubDeleteEmptyFolders(strDirectory, False)
			
			'In case there's a write-lock on one of these files, we don't want Xscript to die.  Just skip it.
			Catch e As Exception
			End Try
			
		End If
	
	End Sub
	
	
'==============================================================================================================
' SubScreenCap
'
' 		Takes a screenshot of the entire primary monitor and saves the file where specified.
'
'		I created this function instead of calling Agent.Desktop.CaptureBitmap because that routine doesn't
'		include the cursor in the screenshot, it also does not allow me use PNG format, and .CaptureBitmap
'		actually retains a lingering hook on the file it outputs which can lead to access violations down the line.
'
' 		@strFilePath:		Full file path to a screenshot file
' 		@blnIncludeMouse:	Whether to include the cursor
'==============================================================================================================

    Public Sub SubScreenCap(ByVal strFilePath As String, Optional blnIncludeMouse As Boolean=True)
	
		Dim bmp = New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, Imaging.PixelFormat.Format32bppRgb)
        Dim gdest As Graphics = Graphics.FromImage(bmp)
        Dim hDestDC As IntPtr = gdest.GetHdc()
        Dim gsrc As Graphics = Graphics.FromHwnd(GetDesktopWindow)
        Dim hSrcDC As IntPtr = gsrc.GetHdc()
        BitBlt(hDestDC, 0, 0, bmp.Width, bmp.Height, hSrcDC, 0, 0, CInt(CopyPixelOperation.SourceCopy))
		
        If blnIncludeMouse Then
            Dim pcin As New CURSORINFO()
            pcin.cbSize = Marshal.SizeOf(pcin)
            If GetCursorInfo(pcin) Then
                Dim piinfo As ICONINFO
                If GetIconInfo(pcin.hCursor, piinfo) Then
                    DrawIcon(hDestDC, pcin.ptScreenPos.x - piinfo.xHotspot, pcin.ptScreenPos.y - piinfo.yHotspot, pcin.hCursor)
                    If Not piinfo.hbmMask.Equals(IntPtr.Zero) Then DeleteObject(piinfo.hbmMask)
                    If Not piinfo.hbmColor.Equals(IntPtr.Zero) Then DeleteObject(piinfo.hbmColor)
                End If
            End If
        End If
        gdest.ReleaseHdc()
        gdest.Dispose()
        gsrc.ReleaseHdc()
        gsrc.Dispose()
				
		If Directory.Exists(Path.GetDirectoryName(strFilePath)) = False Then Directory.CreateDirectory(Path.GetDirectoryName(strFilePath))
		
        bmp.Save(strFilePath, System.Drawing.Imaging.ImageFormat.Png)
		
    End Sub

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure POINTAPI
        Public x As Int32
        Public y As Int32
    End Structure
	
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure CURSORINFO
        Public cbSize As Int32
        Public flags As Int32
        Public hCursor As IntPtr
        Public ptScreenPos As POINTAPI
    End Structure
	
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure ICONINFO
        Public fIcon As Boolean
        Public xHotspot As Int32
        Public yHotspot As Int32
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure
	
    <DllImport("user32.dll", EntryPoint:="GetCursorInfo")> _
    Private Function GetCursorInfo(ByRef pci As CURSORINFO) As Boolean
    End Function
	
    <DllImport("user32.dll")> _
    Private Function DrawIcon(hDC As IntPtr, X As Int32, Y As Int32, hIcon As IntPtr) As Boolean
    End Function
	
    <DllImport("user32.dll", EntryPoint:="GetIconInfo")> _
    Private Function GetIconInfo(hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean
    End Function
	
    <DllImport("user32.dll", SetLastError:=False)> _
    Private Function GetDesktopWindow() As IntPtr
    End Function
	
    <DllImport("gdi32.dll")> _
    Private Function BitBlt(ByVal hdc As IntPtr, ByVal nXDest As Int32, ByVal nYDest As Int32, ByVal nWidth As Int32, ByVal nHeight As Int32, ByVal hdcSrc As IntPtr, ByVal nXSrc As Int32, ByVal nYSrc As Int32, ByVal dwRop As Int32) As Boolean
    End Function
	
    <DllImport("gdi32.dll")> _
    Private Function DeleteObject(hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
		
End Module


