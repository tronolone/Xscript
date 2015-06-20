'Add Application Configuration for Internet Explorer.  UNCHECK BASESTATE OPTION.  URL = "about:blank".  Locator = "Internet Explorer"
'Add Application Configuration for MedCheck.  UNCHECK BASESTATE OPTION.  Process Name = "E:\Program Files\CorVel\MedCheck Server\Testing\MedCheck.exe".  Working Dir = "%TMP%".  Locator = "MedCheck".  Command Line Arguments = "citrix Testing".
'Add reference: Xscript Actions
'Add reference: Xscript Classes
'Add reference: Xscript Excel
'Add reference: Xscript Interpreter
'Add reference: Xscript Results

Imports SilkTest.Ntf.XBrowser
Imports System.Threading
Imports System.IO
Imports System.Text.RegularExpressions

'==============================================================================================================
' Xscript 
'
'		This is a data-driven script that runs a dynamic series of tests based off instructions from Excel files.
'
'		For every desired test, there is a corresponding Excel spreadsheet that contains a series of encoded
'		instructions.  These instructions define each test from start to finish.  The test files are all located
'		in a central respository and are automatically picked up for processing by the script one at a time.  
'		After interpreting and executing each step of every test, the results are written to a highly-formatted 
'		results spreadsheet.
'
'		Hard-coded items:
'			-Path to Constants.xlsx		- This can be found near the top of the Xscript script
'			-Path to Excel Interop DLL  - Referenced in the properties of the Excel Data Loader script as well
'										  as the properties of Results Output script
'
'		Complete documentation can be found in the Xscript.html file included in the original zip file.
'
'		[Created by Chuck Tronolone, tronolone@gmail.com, 31-JUL-2014]
'==============================================================================================================

Public Module Main
	
	'These items need to be declared at the Module level because 'SubClickOnStick' needs to access them
	Public dicConstants As Dictionary(Of String, String)
	Public test As Test
	
	Public Sub Main()

		'Declare objects		
		Dim altThread As New Threading.Thread(AddressOf SubClickOnStick)
		
		'Declare local variables
		Dim dtCurrentDateTime As DateTime
		Dim strXscriptFailMsg As String=Nothing
		Dim blnChildTestJustAdded As Boolean
		Dim strTempErrorMsg As String
		Dim blnMostRecentStepFailed As Boolean
		Dim blnTrailingHyphen As Boolean
				
		Try 
			
			'We need to kill all instances of EXCEL
			SubKillProcess({"EXCEL.EXE"})
			
			'Turn off Accessibility in case it was accidentally left on.  It slows down playback; only toggle it on when needed.
			Agent.SetOption(Options.EnableAccessibility, False)
			
			'Load the constants
			dicConstants = FnLoadSimpleDictionary(strXscriptFailMsg, "C:\System\Xscript\Constants.xlsx", "Name", "Value")
			
			'Give the temp folder a clean slate by deleting it if it exists, then creating from scratch
			If Directory.Exists(dicConstants.Item("DirectoryToStoreTemporaryFiles")) Then Directory.Delete(dicConstants.Item("DirectoryToStoreTemporaryFiles"), True)
			Directory.CreateDirectory(dicConstants.Item("DirectoryToStoreTemporaryFiles"))
			
			'Get the current date and time at this single point
			dtCurrentDateTime = DateTime.Now
																		
			'Loop once for each unhidden xls/xlsx/xlsm file in the 'active' folder
			For Each strTestFilePath As String In FnGetFilePaths(dicConstants.Item("TestsLocation"), {"xls", "xlsx", "xlxm"}, False)
			
				'Create the root test object
				test = New Test(Path.GetFileNameWithoutExtension(strTestFilePath), strTestFilePath)
					
				'Load the list of steps in this file
				test.strStepsList = FnLoadSimpleList(strTestFilePath, dicConstants.Item("InstructionsColumnHeader"), True, True)
				
				'Start clicker thread here if it's not already started
				If dicConstants.Item("SecondsToWaitOnStuck") >= 0 And altThread.IsAlive = False Then
					altThread.Start()
				End If
				
				'Cycle through each step in this test until an error is hit all the steps are executed
				Do While Not test.blnTestDone 
					
					Try
						
						'Troubleshooting
						'Console.WriteLine("Running test [" & test.strTestName & "] @ [" & test.strInstruction & "]")
						
						'Set a boolean to indicate that if a failure occurs at this step, it is to be disregarded
						If String.Equals(Right(test.strInstruction, 1), "-") Then blnTrailingHyphen=True Else blnTrailingHyphen=False
						
						'Execute command
						strTempErrorMsg = FnInterpret(test, dtCurrentDateTime, blnTrailingHyphen, dicConstants)
						
						'If there was a failure, update the definitive boolean and add the failure message to the list of failures
						If Not String.IsNullOrEmpty(strTempErrorMsg) Then
							blnMostRecentStepFailed = True
							test.objFailureDic.Add(test.intCurrentStepIndex, strTempErrorMsg)
						Else
							blnMostRecentStepFailed = False
						End If
							
						'Execute a standard wait between all steps to increase reliability
						SubSleep(dicConstants.Item("DelayBetweenAllSteps"))					
					
					'If an exeception was thrown, add it to the list of failures and set definitive boolean
					Catch ex As Exception
						test.objFailureDic.Add(test.intCurrentStepIndex, ex)
						blnMostRecentStepFailed = True
					End Try
					
					'If this most recent step failed
					If blnMostRecentStepFailed Then
						
						'If the user legitimately cares about errors on this step, follow error procedures
						If Not blnTrailingHyphen Then
							SubScreenCap(dicConstants.Item("ResultsLocation") & "\" & dicConstants.Item("ScreenshotsFolderName") & "\" & dtCurrentDateTime.ToString("yyyyMMdd_HHmmss") & "\" & _
										 test.strRelativeScreenShotPath & "\Step " & FnAddLeadingZeros(test.FnGetExcelPosition(test.intCurrentStepIndex), 4) & ".png")
							
							'If this is a normal circumstance where encountered errors end the test, kill the application
							If test.blnAbortOnFail Then
								
								'If there are processes which must always be killed regardless of Xscript directly kicking them off
								If Len(dicConstants.Item("AlwaysKillOnFailure").Trim) > 0 Then
									For Each strOneProcess In Split(dicConstants.Item("AlwaysKillOnFailure"), ",")
										strOneProcess = strOneProcess.ToLower.Trim
										If Not FnLowercaseList(FnGetActiveExeFileNameList).Contains(strOneProcess)
											FnGetActiveExeFileNameList.Add(strOneProcess)
										End If
									Next strOneProcess
								End If
								
								'Kill the processes kicked off by Xscript as well as any listed in the AlwaysKillOnFailure constant
								SubKillProcess(FnGetActiveExeFileNameList.ToArray)
								SubClearActiveExeFileNameList
								
							End If
						
						'If the user has specified that failures on this step should not hault execution, nix the evidence
						Else
							test.objFailureDic.Remove(test.intCurrentStepIndex)
							blnMostRecentStepFailed = False
						End If
					
					End If
					
					'If the test has more steps AND [we didn't hit an error on this most recent step -OR- this test is meant to continue despite errors occurring]
					If test.intCurrentStepIndex+1 < test.strStepsList.Count AndAlso (blnMostRecentStepFailed=False OrElse test.blnAbortOnFail=False) Then
						
						'If we have a child test that is not concluded, switch test references so it can start executing
						blnChildTestJustAdded = False
						For Each kvp As KeyValuePair(Of String, Test) In test.testChildDic
							If Not kvp.Value.blnTestDone Then
								test = kvp.Value
								blnChildTestJustAdded = True
								Exit For
							End If
						Next kvp
						
						'If we didn't just add a child test, all we need to do is increment the step counter on the current test
						If Not blnChildTestJustAdded Then
							test.intCurrentStepIndex = test.intCurrentStepIndex + 1	
						End If
						
					'For whatever reason, the test is now concluded
					Else
						
						test.blnTestDone = True
						
						'Report the results
						If test.blnReportResults Then 
							SubReportResultToExcel(dicConstants.Item("ResultsLocation") & "\" & dicConstants.Item("ResultsFileName"), test, dtCurrentDateTime, If(String.Compare(dicConstants.Item("IncludeStackTrace"), "true", True)=0, True, False), _
												   dicConstants.Item("ResultsLocation") & "\" & dicConstants.Item("ScreenshotsFolderName"), "png")
						End If
												
						'If this test has a parent test
						If Not test.testParent Is Nothing Then
						
							'If this test is concluded because it simply ran out of steps, set the reference back to the parent test
							If test.intCurrentStepIndex+1 = test.strStepsList.Count AndAlso (blnMostRecentStepFailed=False OrElse test.blnAbortOnFail=False) Then
								test = test.testParent
								
							'Else, it ended beause it errored.  Tell each ancestor to start at their root (step #0).  Then, update the reference to the root.
							Else
								Do
									test = test.testParent
									test.intCurrentStepIndex = 0									
								Loop While Not test.testParent Is Nothing
							End If
							
						End If
						
						'Do manual garbage collection.  This overcomes the random "Playback stopped by user" message that would sometimes hault the script.
						SubDoGarbageCollect
						
					End If
				
				Loop
				
			Next
			
			'Now that all testing is done, we can kill the altThread
			If dicConstants.Item("SecondsToWaitOnStuck") >= 0 Then altThread.Abort()
				
			'Trim back an excess of screenshots
			SubTrimBackScreenshots(dicConstants.Item("ResultsLocation") & "\" & dicConstants.Item("ScreenshotsFolderName"), dicConstants.Item("ScreenshotsRetentionLimit"))
		
		'If something went terribly wrong with the Xscript architecture itself, exit gracefully.
		Catch exAny As Exception
			Workbench.Verify(False, "Unhandled exception thrown that was not specific to any test: " & exAny.GetType.ToString & ".  " & If(String.IsNullOrEmpty(strXscriptFailMsg), "", strXscriptFailMsg & "  ") & _
									"Message: " & exAny.Message & " StackTrace: " & exAny.StackTrace)
			SubKillProcess({"EXCEL.EXE"})
			If dicConstants.Item("SecondsToWaitOnStuck") >= 0 Then altThread.Abort()
			SubDoGarbageCollect
		End Try
		
	End Sub
	
'==============================================================================================================
' SubClickOnStick
'
' 		This sub is a never-ending thread that runs simultaneously with the Main routine above.  It is a dirty
'		hack to overcome an annoying IE8 bug.  In IE8, there exists an issue where that status bar will always
'		display "1 item downloading..." as it tries to grab the last file to load the page in full.  Despite the
'		file having already downloaded, this message remains.  As long as this message remains, Silk will wait
'		for it to finish before continuing script execution.  This means that the IE8 bug will forever hang the
'		script.  This happens very frequently for the file named "Plus.gif" in the ScanOne ODC website.
'
'		Ideally, we could utlize Silk's "Synchronization exclude list" and mark the problematic file for exclusion.
'		However, this feature of Silk appears to be broken as I have repeatedly failed to have it forgo Plus.gif.
'		Alternatively, one would think that we could simply update a counter in the Timing section of the Options
'		menu, but these timers seem not to apply to this particular type of synchronization.  The only way I've 
'		found To clear the "1 item downloading..." message Is To click on any object within	the website itself.
'		Strangely, this resolves the bug.  
'
'		This routine is simulating a mouse click at the cursor's current location if there have been no new
'		instructions in the last X seconds.  X is defined by the Constant named 'SecondsToWaitOnStuck'.  To avoid
'		having the mouse accidentally click on a hyperlink or similar, efforts are made to keep the mouse on top
'		of the last benign element used such as a text field or similar.
'
'		UPDATE 7/29/2014: It appears that this routine is the cause of a very inconsistent bug in Xscript.
'		On rare occassions, Silk will hault execution of Xscript and puts up a message that reads "Playback stopped
'		by user".  If you have Silk's "Output" window open when this happens, you will see reference to an
'		exception that occurred somewhere in the SubClickOnStick routine.  I do not know what is going wrong in
'		this section, but if you simply disable SubClickOnStick by setting the 'SecondsToWaitOnStuck' constant
'		to -1, the problem never crops up.  Of course, this may mean that using IE8 in Silk/Xscript is no longer
'		a viable option for websites with buggy javascript (ScanOne's ODC).
'==============================================================================================================

	Public Sub SubClickOnStick
		
		'Declare variables
		Dim strLastKnownInstruction As String
		
		'Loop forever.  When the primary thread dies or finishes, this thread will end.
		Do While(True)
								
			'Figure out what the current instruction value is on the initial thread
			strLastKnownInstruction = test.strInstruction
			
			'Sleep for a bit
			SubSleep(dicConstants.Item("SecondsToWaitOnStuck"))
		
			'If the instruction isn't empty and it was the same the last time we checked and the objectToClick exists
			If(String.IsNullOrEmpty(strLastKnownInstruction) = False And String.Equals(test.strInstruction, strLastKnownInstruction)) Then
			
				'Mouse down, wait, mouse up
				mouse_event(2, 200, 200, 0, 0)
				SubSleep(0.25)
				mouse_event(4, 200, 200, 0, 0)
				
			End If
			
		Loop
		
	End Sub

'==============================================================================================================
' mouse_event
'
' 		This routine allows us to simulate mouse clicks without having to interact with the browser.
'		We need to avoid using the browser because the primary Main sub needs uninterupted access to it.
'==============================================================================================================

	Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
	
End Module













