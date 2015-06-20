'Add reference: Xscript Classes
'Add reference: System.Data

Imports SilkTest.Ntf.XBrowser
Imports SilkTest.Ntf.WindowsForms
Imports SilkTest.Ntf.Win32
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Globalization

Public Module ActionsModule
		
	'These are declared at module level because different functions need access to them without having them passed from outside
	Private _desktop As Desktop = Agent.Desktop
	Private ci As New CultureInfo("en-US")
	Private strPrintList As New List(Of String)

'==============================================================================================================
' FnTestObjectInvoke
'
' 		Uses Reflection to invoke the specified method of a test object.  If an error occurs, a non-empty string
'		with an error message is returned.
'
'		Though this function appears to be calling methods of the TestObject class, it is not.  The TestObject
'		object (known below as testObj) is immediately reassigned to its specific object type upon executing
'		the line "testObj = _desktop.TestObject(strCtrlName)".  This causes testObj to actually become a DomLink,
'		DomTextField, Control, Window, DomButton, or whatever the strCtrlName represents in the application.
'		Once the cast is complete, the function below has access to every member of that derrived class.
'
'		Please note, Xscript allows the user to invoke methods such as SetText and Select against 'Control' objects.
'		'Control' objects don't actually have these methods, so instead of trying to invoke them, the function 
'		below will handle calls to these non-existant methods using TypeKeys, image recognition tricks, clicks and
'		so on.  To the user, it will appear as they successfully invoked SetText or Select and they will be unaware
'		of the trickery going on behind the scenes.
'
'		@strCmd:			The name of the method to invoke
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'		@objUserArgsList:	The arguments to pass to the method.  Note: the first arg is always the control name.
'							All subsequent args are meant to be passed directly to the specified method.
'==============================================================================================================

	Public Function FnTestObjectInvoke(ByVal strCmd As String, ByRef dicConstants As Dictionary(Of String, String), Optional ByVal objUserArgsList As List(Of Object)=Nothing) As String	
	
		'Declare objects
		Dim testObj As TestObject
		Dim testObjType As Type
		Dim testObjMethod As MethodInfo
		
		'Declare variables
		Dim typeArgsArr(0) As Type 
		Dim objArgsArr() As Object
		Dim strCtrlName As String
		
		'Set values
		FnTestObjectInvoke = ""
		strCtrlName = If(objUserArgsList Is Nothing, Nothing, objUserArgsList(0))
		
		'Populate the objArgsArr array with all user-submitted parameters minus the first one that is assumed to be the ctrl name
		objArgsArr = If(objUserArgsList Is Nothing, Nothing, objUserArgsList.GetRange(1, objUserArgsList.Count-1).ToArray)
		
		Try
			
			'Create an instance of the test object
			testObj = _desktop.TestObject(strCtrlName)
			
			'Create an instance of the test object's type
			testObjType = testObj.GetType
			
			'Proceed as long as the user is not trying to 'SetText' or 'Select' on a 'Control' object
			If String.Equals(testObjType.ToString, "SilkTest.Ntf.Control")=False OrElse (String.Compare(strCmd, "SetText", True)<>0 AndAlso String.Compare(strCmd, "Select", True)<>0) Then
						
				'If the user wants to click a DomLink object, we likely have to give mouse coordinates. This is because .Click clicks in the center Of the object's defining rectangle.
				'Text does Not always exist In this spot.  Alternatively, the user could call '.Select', however .Select does not physically move the mouse and click to follow the hyperlink.
				'The code below will click on the center of the first character in a DomLink, assuming it contains text, and assuming the user isn't already giving us a specific coordinate.
				If String.Equals(testObjType.ToString, "SilkTest.Ntf.XBrowser.DomLink") AndAlso String.Compare(strCmd, "Click", True)=0 AndAlso _
								 objUserArgsList.Count<3 AndAlso Len(_desktop.DomLink(strCtrlName).Text.Trim)>0 Then
					Dim r As SilkTest.Ntf.Rectangle = _desktop.DomLink(strCtrlName).TextRectangle(Left(_desktop.DomLink(strCtrlName).Text, 1))
					If objUserArgsList.Count=1 Then objUserArgsList.Add(MouseButton.Left)
					objUserArgsList.Add(New SilkTest.Ntf.Point(r.X + r.Width/2, r.Y + r.Height/2))
					objArgsArr = objUserArgsList.GetRange(1, objUserArgsList.Count-1).ToArray
				End If
			
				'Create an array of Type objects that reflects every type in the received parameter array
				If Not objArgsArr Is Nothing Then
					ReDim typeArgsArr(objArgsArr.Length-1)
					For i As Integer = 0 To objArgsArr.Length-1
						typeArgsArr(i) = objArgsArr(i).GetType
					Next i
				End If
								
				'Retrieve the requested method having the proper argument types
				testObjMethod = testObjType.GetMethod(strCmd, If(objArgsArr Is Nothing, Type.EmptyTypes, typeArgsArr))
				
				'If no method was found, try to correct the capitalization and try again
				If testObjMethod Is Nothing Then
					For Each oneTestObjMethod As MethodInfo In testObjType.GetMethods
						If String.Compare(strCmd, oneTestObjMethod.Name, True)=0 Then
							strCmd = oneTestObjMethod.Name
							testObjMethod = testObjType.GetMethod(strCmd, If(objArgsArr Is Nothing, Type.EmptyTypes, typeArgsArr))
							Exit For
						End If
					Next oneTestObjMethod
				End If
				
				'Ensure the method defined by both the command and argument types finds a match, then invoke the method
				If Not testObjMethod Is Nothing Then
					testObjMethod.Invoke(testObj, objArgsArr)	
				Else
					FnTestObjectInvoke = "Function lookup failed.  Either there is no .NET function named '" & strCmd & "' for the given object type or you have supplied an invalid sequence of arguments.  " & _
									     "Please ensure you have the right number of arguments and you haven't put double-quotes around non-string values such as booleans, " & _
									     "numbers, and enumerations (MouseButton.Right, ModifierKeys.Control, ClickType.Middle, etc)."
				End If
				
			'Else, the user is trying to call 'Select' or 'SetText' on a 'Control' object.  Control doesn't have these methods - we must emulate them instead.
			Else If String.Compare(strCmd, "Select", True)=0 Then
				FnPerformSelectOnControlObject(testObj, dicConstants, objArgsArr)
			Else 
				FnPerformSetTextOnControlObject(testObj, dicConstants, objArgsArr)
			End If
		
		'We are catching the Reflection exception and throwing the root cause exception so that the message that reaches the user applies to the command he tried to invoke.
		Catch eTI As TargetInvocationException
			Throw eTI.GetBaseException
		End Try
			
	End Function

	
'==============================================================================================================
' FnPerformSelectOnControlObject
'
' 		When the user tries to call "Select" on a "Control" object, there is no actual "Select" method.  This 
'		function is called by FnTestObjectInvoke to emulate the Select method via some clever tricks instead.
'
'		@ctrlObj:		The actual 'Control' test object
'		@dicConstants:	A dictionary of variables that help tweak behavior in the Xscript script
'		@objArgsArr:	The user's array of arguments to pass into the Select method
'==============================================================================================================

	Public Function FnPerformSelectOnControlObject(ByRef ctrlObj As Control, ByRef dicConstants As Dictionary(Of String, String), Optional ByRef objArgsArr() As Object=Nothing) As String
	
		'Declarations
		Dim intCurrentState As Integer
	
		'Set default return
		FnPerformSelectOnControlObject = ""
	
		With ctrlObj
			
			'Proceed only if the object is enabled
			If .Enabled Then
				
				'If we were given a parameter
				If Not objArgsArr Is Nothing AndAlso objArgsArr.Length > 0 Then
				
					'If the parameter is a string, then this is a dropdown menu
					If String.Equals(objArgsArr(0).GetType.ToString, "System.String") Then
						FnPerformSelectOnControlObject = FnSelectDropdownValueForControlObject(ctrlObj, objArgsArr(0), dicConstants)
					
					'If the parameter is an integer between 1 and 2 inclusively, then this is a checkbox
					Else If String.Equals(objArgsArr(0).GetType.ToString, "System.Int32") AndAlso objArgsArr(0)>=1 AndAlso objArgsArr(0)<=2 Then 
						
						'Get the state of the checkbox
						intCurrentState = FnGetObjectProperty(ctrlObj, "State", dicConstants, FnPerformSelectOnControlObject)
						
						'If the state is not want the user wants, toggle it
						If(intCurrentState <> objArgsArr(0))
							.SetFocus
							.TypeKeys("<Space>")
						End If
						
					'Otherwise, we are unequipped to handle the expression
					Else
						FnPerformSelectOnControlObject = "When using the 'Select' command on 'Control' objects, the only acceptable arguments are a string, an " & _
														 "integer of value 1 or 2, or nothing.  You have supplied an invalid parameter."
					End If
					
				'Else if we were not given a parameter, then this is a radio button.  Select it.
				Else
					.SetFocus
					.TypeKeys("<Space>")
				End If
				
			Else
				FnPerformSelectOnControlObject = "The object is disabled."
			End If
			
		End With
	
	End Function
	
	
'==============================================================================================================
' FnPerformSetTextOnControlObject
'
' 		This function is called by FnTestObjectInvoke whenever the user is trying to invoke the SetText method
'		on a Control object.  Control objects do not have SetText as a method, so this function instead emulates
'		what SetText would do by using a series of .TypeKeys and .PressKeys commands.  In short, it highlights
'		all text in the text field, deletes it, then types in every specified letter.
'
'		@ctrlObj:		The actual 'Control' test object
'		@dicConstants:	A dictionary of variables that help tweak behavior in the Xscript script
'		@objArgsArr:	The user's array of arguments to pass into the SetText method
'==============================================================================================================

	Public Function FnPerformSetTextOnControlObject(ByRef ctrlObj As Control, ByRef dicConstants As Dictionary(Of String, String), Optional ByRef objArgsArr() As Object=Nothing) As String
		
		'Default return
		FnPerformSetTextOnControlObject = ""
		
		With ctrlObj
			
			'Highlight all existing text and delete it
			.SetFocus
			.PressKeys("<Right Ctrl>")
			.TypeKeys("<End>")
			.PressKeys("<Left Shift>")
			.TypeKeys("<Home>")
			.ReleaseKeys("<Right Ctrl><Left Shift>")
			.TypeKeys("<Delete>")
			
			'If the user wants to put text into the text field/area, type it into there now
			If Not objArgsArr Is Nothing Then
				.TypeKeys(objArgsArr(0))
			End If
			
		End With
		
	End Function
	
	
'==============================================================================================================
' FnSelectDropdownValueForControlObject
'
' 		This function is called by FnPerformSelectOnControlObject when it has been determined that the user is
'		trying to invoke Select on a Control object representing a dropdown menu.  Since Control objects do not
'		have Select as a method, this function instead emulates what Select would do if it were applied to a 
'		normal dropdown menu.  It does this by calling .TypeKeys and/or clicking on the dropdown menu to expose
'		the possible values themselves.  Some dropdowns will allow you to type out your selection in full, some
'		will only jump to the first letter typed, and others will not respond to typing letters.  Frequently, the
'		up and down arrows on the keyboard are utilized to reach the desired value.
'
'		@ctrlObj:		The actual 'Control' test object
'		@strValue:		The value the user wants selected
'		@dicConstants:	A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================

	Private Function FnSelectDropdownValueForControlObject(ByRef ctrlObj As Control, ByVal strValue As String, ByRef dicConstants As Dictionary(Of String, String)) As String
		
		'Default return
		FnSelectDropdownValueForControlObject = ""
		
		'Variables
		Dim intRepeatCounter As Integer
		Dim strPreviousValue As String
		Dim strCurrentValue As String
		Dim blnTypingAffectsSelectedValue, blnAccidentallySelectedCorrectValue As Boolean
		Dim r As SilkTest.Ntf.Rectangle
		
		With ctrlObj
			
			'Obtain the starting value
			strPreviousValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)
						
			'If the text box isn't already set to the desired value
			If String.Compare(strPreviousValue, strValue, True)<>0 Then
			
				'Click the right side of the dropdown menu so the contents display
				'The wait is necessary because the previous command was a double-click which exposed the ctrl object.  This happened within the interpreter.
				'If you move from double-clicking to trying to click the down arrow to expose the dropdown elements too quickly, the click doesn't actually work.  This is purely a MC issue.
				r = .GetRect
				SubSleep(1.5)
				.Click(MouseButton.Left, New SilkTest.Ntf.Point(r.Width-5, r.Height/2))
				
				'Press the first letter to advance the selection
				.TypeKeys(Left(strValue, 1))
				
				'Get the current value after typing the first letter
				strCurrentValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)
				
				'If typing a letter modified the value, set a boolean to that effect.  Otherwise, try pressing <down> then <up> to see if the value changes
				If Not String.Equals(strCurrentValue, strPreviousValue) Then
					blnTypingAffectsSelectedValue=True
					If String.Compare(strCurrentValue, strValue, True)=0 Then blnAccidentallySelectedCorrectValue = True
				
				'If typing the first key didn't change the dropdown's value
				Else
					
					'Press the <Down> button and see if the value changes
					.TypeKeys("<Down>")
					strCurrentValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)
					
					'If typing the first key didn't change anything, but pressing <Down> did
					If Not String.Equals(strCurrentValue, strPreviousValue) Then
						blnTypingAffectsSelectedValue=True
						If String.Compare(strCurrentValue, strValue, True)=0 Then blnAccidentallySelectedCorrectValue = True
						
					'If typing the first key didn't change anything and pressing <Down> didn't help either, try <Up> instead
					Else
						
						strPreviousValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)
						.TypeKeys("<Up>")
						strCurrentValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)
						
						'If pressing the <Up> button actually changed its value
						If Not String.Equals(strCurrentValue, strPreviousValue) Then
							blnTypingAffectsSelectedValue = True
							If String.Compare(strCurrentValue, strValue, True)=0 Then blnAccidentallySelectedCorrectValue = True
						End If
						
					End If
				End If
				
				'If typing into the field immediately changes its value and we aren't already at the desired value
				If blnTypingAffectsSelectedValue And Not blnAccidentallySelectedCorrectValue Then
					
					'Loop until the value is what we want or our repeatCounter is excessively high
					Do While String.Compare(strValue, FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject), True) <> 0 And intRepeatCounter<32
					
						'Press <down> to the next value.  If we end up at the wrong firstLetter, then we're done going <down>.
						If intRepeatCounter <=8 And String.Equals(Left(strPreviousValue, 1), Left(strValue, 1)) Then
							.TypeKeys("<Down>")
							If String.Compare(Left(FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject), 1), Left(strValue, 1), True)<>0 Then
								intRepeatCounter=8
								.TypeKeys(Left(strValue, 1))
							End If
						
						'If we already tried up the <down> approach, press the firstLetter and then go <up>, the we're done going <up>.
						Else If intRepeatCounter <=16
							.TypeKeys("<Up>")
							If String.Compare(Left(FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject), 1), Left(strValue, 1), True)<>0 Then
								intRepeatCounter=16
								.TypeKeys(Left(strValue, 1))
							End If
							
						'If we already tried going <up> and <down> from the firstLetter position, then go all the way <down> to the bottom of the list
						Else If intRepeatCounter <=24
							.TypeKeys("<Down>")
							
						'If we already tried going <up> and all the way <down> from the firstLetter position, then go all the way <up> to the top of the list
						Else If intRepeatCounter <=32
							.TypeKeys("<Up>")
						
						'Else we can't find the value
						Else
							FnSelectDropdownValueForControlObject = "After traversing every value in the dropdown menu, it appears there is no value reading '" & strValue & "'."
						End If
						
						'If the new value is the same as the old value, increment the repeat counter
						If String.Equals(strPreviousValue, FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)) Then
							intRepeatCounter = intRepeatCounter + 1
						End If
						
						'Store the previous value
						strPreviousValue = FnGetObjectProperty(ctrlObj, "Text", dicConstants, FnSelectDropdownValueForControlObject)							
						
					Loop
					
				'If typing into the field does not actually change its active value, just type the whole word out and press ENTER
				'I'm not actually sure how to verify that the value was actually selected becausing pressing ENTER on certain types of CTRL-based dropdowns causes them to become unattachable.
				'At the moment, the fields that do this can be reactivated by clicking in them, waiting half a second, then clicking again.  However, I don't know if this is the norm yet.
				Else
					'These backspaces get rid of the last 3 characters typed.  Up to this point, it's the first letter of strValue, <down>, and <up>.
					'It's necessary to remove these characters in order to type a new string without having to wait 2 seconds for the application dropdown search algorithm to reset.
					.TypeKeys("<Backspace><Backspace><Backspace>" & strValue)
				End If
				
				'Press Enter to apply the typed out keys and select the corresponding value
				.TypeKeys("<Enter>")
				
			End If
			
		End With
			
	End Function
	
'==============================================================================================================
' FnRun
'
' 		Executes the specified file via .NET's Process class.  However, if the path to a browser is passed, then
'		it instead calls FnLaunchBrowser which utilizes Silk's BrowserBaseState class.  It is believed that using
'		BrowserBaseState allows for the use of the AJAX/HTML Synchronization feature in Silk; though I've never
'		actually tested to see if synchronization works when you just open browser via the Process class.
'
'		@strPath:			Full path to the file you want to execute
'		@blnMayUseExisting:	If TRUE and the process is already running, it will not execute the file.  If FALSE,
'							then any existing process for the given file will first be killed then executed.
'		@strArgs:			Arguments to send to the file upon executing it
'==============================================================================================================

	Public Function FnRun(ByVal strPath As String, ByVal blnMayUseExisting As Boolean, Optional ByVal strArgs As String="") As String
	
		'Variables
		Dim strFileName As String
		
		'Set variables
		FnRun = ""
		strFileName = Path.GetFileName(strPath)
	
		'If the user says we're not allowed to use existing instances, kill all instances now
		If Not blnMayUseExisting Then
			SubKillProcess({strFileName})
		End If
		
		'If we've just killed our instances OR our instance is not currently running, then start it up
		If(Not blnMayUseExisting OrElse Not FnProcessExists(strFileName)) Then
		
			'If the user is loading a browser, start the browser base state.  Otherwise, just execute the specified file.
			'It's important that we use Silk's browser toolkit here because the Synchronization feature will be useless otherwise.
			Select Case strFileName.ToUpper
				Case "IEXPLORE.EXE", "FIREFOX.EXE", "CHROME.EXE"
					FnRun = FnLaunchBrowser(strFileName, strArgs)
				Case Else
					Dim myProcess As New Process()
					Dim psi As New ProcessStartInfo()
					psi.FileName = strPath
					psi.Arguments = strArgs
					psi.WorkingDirectory = Environment.GetEnvironmentVariable("TMP")
					myProcess.StartInfo = psi
					myProcess.Start()
			End Select
			
		'If our instance is running, the user passed args, and the executable is a browser, just do a normal navigate
		Else If(Len(strArgs)>0 AndAlso (String.Compare(strFileName, "IEXPLORE.EXE", True)=0 Or String.Compare(strFileName, "FIREFOX.EXE", True)=0 Or String.Compare(strFileName, "CHROME.EXE", True)=0)) Then
			FnNavigate(strArgs)			
		End If
		
	End Function
	
'==============================================================================================================
' FnLaunchBrowser
'
' 		This function is called by FnRun when the user wants to run a browser.  It makes use of Silk's native
'		BrowserBaseState class to launch either IE, FireFox, or Chrome.  We chose to launch browsers in this
'		fashion because it's likely that it gives us access to built-in Silk features such as AJAX/HTML
'		Synchronization.
'
'		This function handles instances where the crash recovery dialog appears in IE8/9.
'
' 		@strFileName:	The file name of the browser to launch
'		@strURL:		The address in which to navigate
'==============================================================================================================

	Public Function FnLaunchBrowser(ByVal strFileName As String, Optional ByVal strUrl As String="about:blank") As String
		
		'Declare objects
		Dim bbs As BrowserBaseState = Nothing
	
		'Declare variables
		Dim intAttempts As Integer
		
		'Set variables
		FnLaunchBrowser = ""
		intAttempts = 0
				
		'Create a new browser base state
		Select Case strFileName.ToUpper
			Case "IEXPLORE.EXE"
				bbs = New BrowserBaseState(BrowserType.InternetExplorer, strUrl)
			Case "FIREFOX.EXE"
				bbs = New BrowserBaseState(BrowserType.FireFox, strUrl)
			Case "CHROME.EXE"
				bbs = New BrowserBaseState(BrowserType.GoogleChrome, strUrl)
			Case Else
				FnLaunchBrowser = "FnLaunchBrowser does not know which browser is specified by '" & strFileName & "'."
		End Select
				
		'This loop allows us to attempt basestate execution twice in case the first attempt is met with the crash recovery dialog box
		Do Until intAttempts >= 2
		
			Try
				
				'Excute the base state.  This opens the browser and goes the specified URL.
				bbs.Execute()
				
				'If the IE9+ crash recovery dialog appears, close it now
				If(_desktop.Exists("/BrowserApplication//Control[@windowClassName='DirectUIHWND'][2]")) Then			
					
					'We have to toggle this option because it allows Silk to interact with IE9+ dialog box child elements
					Agent.SetOption(Options.EnableAccessibility, True)
				
					'Close the dialog box
					_desktop.BrowserApplication.PushButton("[@caption='Close']").Click
					
					'Toggle the setting off because according to Silk documentation, it slows down playback
					Agent.SetOption(Options.EnableAccessibility, False)
					
				End If
				
				'Set an indicator to tell the loop that we succeeded and can move on now
				intAttempts = 3
				
			'Should there be a problem going to the base state, we'll handle the exception here.
			Catch ex As Exception
				
				'Increment our attempts counter
				intAttempts = intAttempts + 1
				
				'If that IE8 crash recovery dialog box appears
				If(_desktop.exists("/BrowserApplication//Dialog//Control[@caption='Go to*home page']")) Then 				
					
					'First, bring the dialog to the foreground because it might be hiding behind another window
					_desktop.BrowserApplication.Dialog("@caption='*Internet Explorer'").SetActive
					
					'Now, click on the 'Go to your home page' button
					_desktop.BrowserApplication.Dialog.Control("@caption='Go to*home page'").Click
					
				'If the crash recovery dialog box isn't the cause of this exception, return an error message 
				Else
					FnLaunchBrowser = "Unknown exception while loading browser base state.  Message: " & ex.Message
					intAttempts = 3
				End If

			End Try

		Loop
			
		'Maximize the IE window
		_desktop.BrowserApplication.Maximize
		
		'Give Windows a second to actually complete the maximize command before we let Silk attach to anything
		SubSleep(2)
		
	End Function
	

'==============================================================================================================
' FnNavigate
'
' 		Redirects the browser to the specified URL.
'
'		@strURL:	The URL in which to navigate
'==============================================================================================================

	Public Function FnNavigate(ByVal strURL As String) As String
		
		FnNavigate = ""
		
		_desktop.BrowserWindow.Navigate(strURL)
		
	End Function
	
	
'==============================================================================================================
' FnGetCellXpath
'
' 		Returns either the xpath of a particular HTML table cell, or the xpath to a particular input field embedded
'		within a MedCheck grid.
'
'		For HTML tables, the row and column identifiers can be either strings or integers.  If an integer is used
'		such as "3", then this indicates the requested cell exists in the 3rd row/column.  The user may pass in
'		a row identifier of "3" and a column identifier of "First Name".  This retrieve the xpath of the cell
'		in the third row under the column entitled "First Name".  However, it will only work if the first row
'		of the table is being utilized to store these column headers.
'
'		One common "gotcha" for trying to reference particular table cells by integers is that often there are
'		hidden rows/columns inserted by dev for cross-browser compatibility issues.  This throws off the expected
'		sequence of rows and columns.
'
'		For MedCheck grids, you are only allowed to reference particular cells by a string column header and
'		an integer row number.  These grids are impossible to disect appropriately because Silk is not readily
'		compatible with Delphi forms.  To get the xpath for a cell in such a grid, there is a lot of trickery
'		going on in this function that includes keypresses, double-clicks, OCR and image recognition.  Please
'		follow the commenting within the function to understand the process.
'
'		@strTableGridCtrlName:	Xpath of the HTML table or grid
'		@strRowIdentifier:		Row label text or integer sequence
'		@strColumnIdentifier:	Column header text or integer sequence
'		@dicConstants:			A dictionary of variables that help tweak behavior in the Xscript script
'		@strErrorMsg:			Variable passed ByRef to populate with an error message should it be needed
'==============================================================================================================
	
	Public Function FnGetCellXpath(ByVal strTableGridCtrlName As String, ByVal strRowIdentifier As String, ByVal strColumnIdentifier As String, ByRef dicConstants As Dictionary(Of String, String), ByRef strErrorMsg As String) As String
		
		'Get the table obj
		Dim testObj As TestObject
		Dim ctrlObj As Control=Nothing
		Dim rColumnLabel As SilkTest.Ntf.Rectangle
		Dim pointAsterisk As Point
		Dim pointOutlining As Point
		Dim x, y As Integer
		Dim intAttemptCounter As Integer=1
		Dim bmp1 As Bitmap
		
		'Default return
		FnGetCellXpath = Nothing
		
		'Get the table object
		testObj = _desktop.TestObject(strTableGridCtrlName)
	
		'If this is an html table object, call FnGetHTMLCellXpath to figure out the cell xpath
		If String.Equals(testObj.GetType.ToString, "SilkTest.Ntf.XBrowser.DomTable") Then
			FnGetCellXpath = FnGetHTMLCellXpath(strTableGridCtrlName, strRowIdentifier, strColumnIdentifier, strErrorMsg)
			
		'If this is a ctrl object, we are going to assume it's the medcheck grid
		Else If String.Equals(testObj.GetType.ToString, "SilkTest.Ntf.Control") Then 
		
			'Set our control object
			ctrlObj = _desktop.Control(strTableGridCtrlName)
		
			'Use OCR to find the coordinates of the header
			rColumnLabel = ctrlObj.TextRectangle(strColumnIdentifier)
			
			'Take a picture of the grid
			bmp1 = FnCaptureBitmap(ctrlObj, dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\Temp.bmp", True)
			
			'Get a .NET Point object representative of where the asterisk may exist inside the captured bitmap above
			pointAsterisk = FnGetContainedImagePoint(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckNewBillLineAsteriskFilename")))
			
			'If there is no asterisk in the grid, it means that no new row has just been added.  So, we can click and type outside of the new row to activate a particular cell.
			'NOTE: You can't willy-nilly click and type within the grid immediately after adding a new row.  If you do so, the new row will simply disappear.
			If pointAsterisk = Nothing Then
				
				'Scroll to the top of the grid
				ctrlObj.VerticalScrollBar.ScrollToMin
				
				'Click in the cell just below the header to ensure it's activated
				ctrlObj.Click(MouseButton.Left, New SilkTest.Ntf.Point(Math.Round(rColumnLabel.X + rColumnLabel.Width/2, 0), Math.Round(rColumnLabel.Y+18+rColumnLabel.Height/2, 0)))
				
				'Press the down button as many times as necessary to get to the proper cell
				For i As Integer = 1 To strRowIdentifier-1
					ctrlObj.TypeKeys("<Down>")
				Next i
				
				'Now that we're in the proper cell, we need to take another picture so we can search for the cell outlining
				bmp1 = FnCaptureBitmap(ctrlObj, dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\Temp.bmp", True)
				
				'Find the position of the outlining
				pointOutlining = FnGetContainedImagePoint(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckGridCellActiviationOutlinePatternFileName")))
				
				'Set x and y
				x = pointOutlining.X
				y = pointOutlining.Y
				
			'The asterisk is on the screen.  Set x and y based off coordinates of the asterisk and coordinates of the header.
			Else
				X = rColumnLabel.X
				Y = pointAsterisk.Y
			End If

			'Click on the cell to either highlight or activate it
			ctrlObj.Click(MouseButton.Left, New SilkTest.Ntf.Point(x+1, y+1))
			
			'If this is a new row, then it is only highlighted at the moment, not activated.  We need to click it again in order to activate it.  However, we must make it a slow click.
			If Not pointAsterisk = Nothing Then
				SubSleep(0.75)
				ctrlObj.Click(MouseButton.Left, New SilkTest.Ntf.Point(x+1, y+1))
			End If
			
			'If the control doesn't appear in the grid, keep waiting and clicking up to 10 times until it does appear.  We do this because MedCheck may be busy processing something.
			Do While FnGetCellXpath Is Nothing
				Try
					FnGetCellXpath = _desktop.Control(strTableGridCtrlName).Find(Of Control)("//Control[@windowClassName='*']").GenerateLocator	
				Catch onfEx As ObjectNotFoundException
					intAttemptCounter = intAttemptCounter + 1
					If intAttemptCounter >= 10 Then
						Exit Do
					Else
						SubSleep(2)
						ctrlObj.Click(MouseButton.Left, New SilkTest.Ntf.Point(x+1, y+1))
					End If
				End Try
			Loop
			
		End If
		
	End Function

'==============================================================================================================
' FnGetHTMLCellXpath
'
' 		This function is called by FnGetCellXpath when it's realized that the user is trying to get the xpath
'		of an HTML table cell and not the xpath of a MedCheck grid cell.
'
'		The user either passes string or integer values for both the row/column identifiers.  The function will
'		figure out where the row and column exists, then find their intersecting cell and return the unique
'		xpath to that cell.
'
'		@strTableCtrlName:		Xpath to the HTML table
'		@strRowIdentifier:		String or integer identifier for the row's label
'		@strColumnIdentifier:	String or integer identifier for the cell's column
'		@strErrorMsg:			String variable passed ByRef to populate with an error message just in case
'==============================================================================================================
	
	Public Function FnGetHTMLCellXpath(ByVal strTableCtrlName As String, ByVal strRowIdentifier As String, ByVal strColumnIdentifier As String, ByRef strErrorMsg As String) As String
		
		'Declare variables
		Dim intCorrectRowIndex As Integer
		Dim intCorrectColumnIndex As Integer
		Dim intTemp As Integer
		Dim i, j As Integer
		Dim strPossibleTags() As String = {"TH", "TD"}
		Dim strPossibleProperties() As String = {"Text", "textContents"}
		
		'Set Defaults
		FnGetHTMLCellXpath = ""
		intCorrectRowIndex = -1
		intCorrectColumnIndex = -1
	
		With _desktop
	
			'If the row identifier is an integer, we don't need to do any searching
			If Integer.TryParse(strRowIdentifier, intTemp) Then
			
				'Ensure the row actually exists in the table
				If(.Exists(strTableCtrlName & "//TR[" & strRowIdentifier & "]"))
					intCorrectRowIndex = intTemp
				Else
					strErrorMsg = "FnGetHTMLCellXpath could not locate row #" & strRowIdentifier & "."
				End If
				
			'If the row identifier is a row label, we need to go through every row in the table to find it
			Else				
				
				'Loop through every cell in the first column.  We are looking for a cell that has either 'Text' or 'textContents' that matches the given strRowIdentifier value.
				'NOTE: This only loops through the first column.  It may be the case that the first column is hidden or utilized in some unknown way.  If so, then really we ought to be
				'looping through the second, third, maybe even fourth columns if not match is found on the first column.  Fix this.  It should be simple.
				i = 1
				Do While .Exists(strTableCtrlName & "//TR[" & i & "]") And intCorrectRowIndex = -1
					For Each strProperty As String In strPossibleProperties
						If(.Exists(strTableCtrlName & "//TR[" & i & "]//TD[@" & strProperty & "='" & strRowIdentifier & "']")) Then
							intCorrectRowIndex = i
							Exit For
						End If
					Next strProperty
					i = i + 1
				Loop
				If intCorrectRowIndex = -1 Then
					strErrorMsg = "FnGetHTMLCellXpath could not locate a cell in the first column containing the text '" & strRowIdentifier & "'."
				End If
			End If
			
			'If we successfully found the row index, let's move on to the column index
			If(String.IsNullOrEmpty(strErrorMsg)) Then
				
				'If the column identifier is an integer, we don't need to do any searching
				If Integer.TryParse(strColumnIdentifier, intTemp) Then
				
					'Ensure the column actually exists in the table
					If(.Exists(strTableCtrlName & "//TR//TH[" & strColumnIdentifier & "]") OrElse .Exists(strTableCtrlName & "//TR//TD[" & strColumnIdentifier & "]"))
						intCorrectColumnIndex = intTemp
					Else
						strErrorMsg = "FnGetHTMLCellXpath could not locate column #" & strColumnIdentifier & "."
					End If
					
				'If the column identifier is a string, search for the corresponding column index
				Else
					
					'Search through every tag and property combination for the needed value.
					'NOTE: The logic is not strong enough.  If the column headers are not actually in the first row, then this fails.  It may very well be the case that there's a "super header" that stretches the entire
					'length of the table just to name the table, then the subsequent row is used for headers.  To account for this, we really ought to be looping through the first 5 rows to search for the header.
					For Each strTag As String In strPossibleTags								
						For Each strProperty As String In strPossibleProperties									
							If(.Exists(strTableCtrlName & "//TR//" & strTag & "[@" & strProperty & "='" & strColumnIdentifier & "']")) Then
								j = 1
								Do While intCorrectColumnIndex = -1 AndAlso (.Exists(strTableCtrlName & "//TR//TH[" & j & "]") OrElse .Exists(strTableCtrlName & "//TR//TD[" & j & "]"))
									If(.Exists(strTableCtrlName & "//TR//" & strTag & "[" & j & "]") AndAlso String.Compare(strColumnIdentifier, .TestObject(strTableCtrlName & "//TR//" & strTag & "[" & j & "]").GetProperty(strProperty), True)=0) Then
										intCorrectColumnIndex = j
									End If
									j = j + 1
								Loop
								If intCorrectColumnIndex > -1 Then Exit For
							End If
						Next strProperty
						If intCorrectColumnIndex > -1 Then Exit For
					Next strTag
					If intCorrectColumnIndex = -1 Then
						strErrorMsg = "FnGetHTMLCellXpath could not locate the column labeled '" & strColumnIdentifier & "'."
					End If
					
				End If
				
			End If
			
			'Set the cell's xpath return value depending on whether it's a TH or TD object
			If .Exists(strTableCtrlName & "//TR[" & intCorrectRowIndex & "]//TD[" & intCorrectColumnIndex & "]") Then
				FnGetHTMLCellXpath = strTableCtrlName & "//TR[" & intCorrectRowIndex & "]//TD[" & intCorrectColumnIndex & "]"
			Else If .Exists(strTableCtrlName & "//TR[" & intCorrectRowIndex & "]//TH[" & intCorrectColumnIndex & "]") Then 
				FnGetHTMLCellXpath = strTableCtrlName & "//TR[" & intCorrectRowIndex & "]//TH[" & intCorrectColumnIndex & "]"
			Else
				FnGetHTMLCellXpath = "The table does not have a cell specified by the coordinates of [" & intCorrectRowIndex & ", " & intCorrectColumnIndex & "].  " & _
								 	 "It may be the case that some cells are effectively merged through 'rowspan' and 'colspan' attributes.  Please adjust accordingly."
			End If
			
		End With
		
	End Function
	
'==============================================================================================================
' FnGetXpathFromCoordinateWithinParent
'
' 		Sometimes, every property of a test object is ambiguous.  For these objects, we typically use an index
'		value to reference the occurrance of the object.  However, the index may not be reliable because hidden
'		fields come and go depending on circumstance and that alters all the indexes in the application.  So,
'		this function was invented to derive the xpath of objects on the fly during runtime.
'
'		Given a parent object and a coordinate, this function returns the xpath of the largest object found
'		at the coordinate relative to the upper-left corner of the parent.  The reason it chooses the largest
'		object and not the smallest is because MedCheck sometimes hides several small text boxes behind larger
'		ones.  I believe that this is done as an alternative to making the small text boxes legitimately 
'		invisible.  If they were made to be legitimately invisible instead, Silk wouldn't recognize them and
'		choosing between the small or large object at the given coordinate would be a moot point.
'
'		This function has a major shortcoming in that it requires the user to give us a parent object.  While
'		this has worked for our purposes thus far, it may one day be the case that even in its nearest parent,
'		the object changes location depending on other fields appearing or disappearing.  Ideally, what ought
'		to be happening here is that the user doesn't need to give us a parent; instead, he need only give us
'		ANY object to use a starting point in order to locate the desired field.  This would allow the user
'		to send the xpath of a simple label that always appears immediately adjacent to the field in question.
'
'		There is a good reason that I didn't code it that way.  When the user supplies a parent object, Silk
'		only has to scan all child objects to find the object occupying the given coordinate.  This scan takes
'		a measureable amount of time.  For ~20 objects, there is probably a 5-second delay.  If the user were
'		to supply ANY object as the starting location, we would have to scan EVERY object on the page to find
'		the one occupying the space in question.  For some forms, this could mean examining 1000+ objects - the 
'		delay would be ridiculous.
'
'		However, if that becomes the only way identify some objects, then the delay is necessary.  If the 
'		object reference happens inside a loop, that delay will repeat for every loop and that's bad.  So, we
'		could instead cache the xpath once it's determined, then on repeat calls to that object, instead of 
'		deriving it all over again, we can see if the cached xpath occupies the space in question first.  If
'		so, then no scan is necessary.
'
' 		@strParentXpath:	The xpath to the parent which contains the needed object
' 		@pnt:				A native SilkTest point that identifies the needed object as it relates to the 
'							upper-left of the parent
'==============================================================================================================
		
	Public Function FnGetXpathFromCoordinateWithinParent(ByVal strParentXpath As String, ByRef pnt As SilkTest.Ntf.Point) As String
	
		'Variables
		Dim intLargestArea, intIndex As Integer
		Dim rParent, rTestObj As SilkTest.Ntf.Rectangle
		Dim testObjParent As TestObject
		
		'Set values
		FnGetXpathFromCoordinateWithinParent = Nothing
		testObjParent = _desktop.TestObject(strParentXpath)
		intIndex = 1
		
		'Get the position of the parent relative to the containing window
		rParent = testObjParent.GetRect
		
		'Cycle through each available object within the parent container
		For Each testObj As TestObject In testObjParent.FindAll("//*")
			
			'Get a rectangle for this test object
			rTestObj = testObj.GetRect
			
			'If this object is in the space defined by the given coordinate
			If Not rTestObj Is Nothing AndAlso rParent.X + pnt.X >= rTestObj.X AndAlso rParent.X + pnt.X <= rTestObj.X + rTestObj.Width AndAlso rParent.Y + pnt.Y >= rTestObj.Y AndAlso rParent.Y + pnt.Y <= rTestObj.Y + rTestObj.Height 
			
				'Ensure this object is the largest available object at the given coordinate if more than one object occupies the same space
				'This seems counterintuitive.  However, there are large text fields in MC that have smaller hidden text fields behind them.
				If intLargestArea = 0 OrElse rTestObj.Width * rTestObj.Height > intLargestArea Then
					
					'Set a record for the smallest area
					intLargestArea = rTestObj.Width * rTestObj.Height
					
					'Get the locator for this object
					FnGetXpathFromCoordinateWithinParent = testObj.GenerateLocator
					
					'If this testObj can't derive itself from the use of its own locator, then the locator is not unique.  We need to cycle through each index until we find the right one.
					If Not testObj.Equals(_desktop.TestObject(FnGetXpathFromCoordinateWithinParent)) Then
						Do While Not testObj.Equals(_desktop.TestObject(FnGetXpathFromCoordinateWithinParent & "[" & intIndex & "]"))
							intIndex = intIndex + 1
						Loop
						FnGetXpathFromCoordinateWithinParent = FnGetXpathFromCoordinateWithinParent & "[" & intIndex & "]"
					End If
					
				End If
				
			End If
				
		Next testObj
		
		If FnGetXpathFromCoordinateWithinParent Is Nothing Then
			Throw New Exception("You are trying to interact with an object defined in the object map by a coordinate relative to a containing object.  However, there are 0 objects present at the given coordinate.")
		End If
		
	End Function


'==============================================================================================================
' SubKillProcess
'
' 		Kills any Windows process matching the supplied name for the logged in user.  
'
' 		@strProcessNames():	An array of the names of the proccess to kill. ex: {"IEXPLORE.EXE", "MEDCHECK.EXE"}
'==============================================================================================================
	
	Public Sub SubKillProcess(ByVal strProcessNameArr() As String)

	    'Variables
	    Dim objProcess As Object
	    Dim intReturn As Integer
		Dim strNameOfUser As String
		
		'Set variables
		strNameOfUser = "placeHolder"
		
		'Loop once for every process to kill.  NOTE: I'd update the SQL below so I don't have to loop, but the approached has failed.
		For Each strProcessName As String In strProcessNameArr
		
		    'Loop through every process matching the given name.  This returns processes for which the current user is and is not the owner.
		    For Each objProcess In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & strProcessName & "'")
						
				Try
					
			        'Populate a new strNameOfUser variable with the name of this particular process's owner.
					'This may throw an exception if the user does not have permission to view info on the process because he's not the owner.
			        intReturn = objProcess.GetOwner(strNameOfUser)
			        
			        'If the GetOwner() method was successful
			        If (intReturn = 0) Then
			        
			            'If the username logged into this Windows session is equal to the owner of the process
			            If (String.Equals(Environ$("Username"), strNameOfUser)) Then
			            
			                'Kill the process
			                objProcess.Terminate
			                
			            End If
			            
			        End If
					
				Catch exKill As Exception
					'We are catching an exception here because it's likely due to a permissions violation between users on the same machine.
					'After the catch, we move onto the next process.  Eventually, we will end up killing only OUR processes as intended.
				End Try

		    Next objProcess
			
		Next strProcessName
	    
	    'Sleep for a second so Windows can catch its breath
	    SubSleep (1)

	End Sub

'==============================================================================================================
' FnProcessExists
'
' 		Returns TRUE if the current user has that given process running 
'
' 		@strProcessName:	The name of the process to check for existance.  ex: "IEXPLORE.EXE" or "EXCEL.EXE"
'==============================================================================================================
	
	Public Function FnProcessExists(ByVal strProcessName As String) As Boolean

	    'Variables
	    Dim objProcess As Object
	    Dim intReturn As Integer
		Dim strNameOfUser As String
		
		'Set variables
		FnProcessExists = False
		strNameOfUser = "placeHolder"
	    
	    'Loop through every process matching the given name.  This returns processes for which the current user is and is not the owner.
	    For Each objProcess In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & strProcessName & "'")
					
			Try
				
		        'Populate a new strNameOfUser variable with the name of this particular process's owner.
				'This may throw an exception if the user does not have permission to view info on the process because he's not the owner.
		        intReturn = objProcess.GetOwner(strNameOfUser)
		        
		        'If the GetOwner() method was successful
		        If (intReturn = 0) Then
		        
		            'If the username logged into this Windows session is equal to the owner of the process
		            If (String.Equals(Environ$("Username"), strNameOfUser)) Then
		            
		                'Set a boolean and exit
		                FnProcessExists = True
						Exit For
		                
		            End If
		            
		        End If
				
			Catch exKill As Exception
				'We are catching an exception here because it's likely due to a permissions violation between users on the same machine.
				'After the catch, we move onto the next process.  Eventually, we will end up killing only OUR processes as intended.
			End Try

	    Next objProcess

	End Function
	
	
'==============================================================================================================
' SubDoGarbageCollect
'
' 		Executes a manual .NET garbage collection.  This is necessary sometimes in order to release hold on
'		an object that is tied to an Excel.exe process.  Once the hold is release, the process terminates itself.
'==============================================================================================================

	Public Sub SubDoGarbageCollect
		GC.Collect()
		GC.WaitForPendingFinalizers
	End Sub
	

'==============================================================================================================
' FnWait
'
'		Waits the specified number of seconds for an object or a particular object proper to appear or disappear.
'		Alternatively, if only one argument is supplied, the a simple sleep will occur for the specified number
'		of seconds.
'
'		@dblNumSecondsToWait:	The number of seconds to wait outright or for an object/property to appear/disappear
'		@strCtrlName:			Xpath to the object in question
'		@objWaitForAppear:		If TRUE, then it waits for the object/property to appear
'		@strPropertyName:		Name of the property whose value we are waiting to appear/disappear
'		@strPropertyValue:		Value of the property we want to see appear/disappear
'		@dicConstants:			A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================

	Public Function FnWait(ByVal dblNumSecondsToWait As Double, _
						   Optional ByVal strCtrlName As String=Nothing, _
						   Optional ByVal objWaitForAppear As Object=Nothing, _
						   Optional ByVal strPropertyName As String=Nothing, _
						   Optional ByVal strPropertyValue As String=Nothing, _
						   Optional ByRef dicConstants As Dictionary(Of String, String)=Nothing) As String
					
		'Declarations
		Dim strTemp As String=Nothing
		Dim dlbOriginalWaitTime As Double=dblNumSecondsToWait
		
		'Set values
		FnWait = ""
		
		With _desktop
			
			'If the ctrl name parameter is blank, then we're just doing a simple sleep
			If strCtrlName Is Nothing Then
				SubSleep(dblNumSecondsToWait)
			
			'Else, the user gave us a ctrl name.  If the strPropertyName and strPropertyValue strings are nothing, then we're waiting for object to appear/disappear
			Else If strPropertyName Is Nothing And strPropertyValue Is Nothing Then
				If objWaitForAppear Is Nothing OrElse (String.Equals(objWaitForAppear.GetType.ToString, "System.Boolean") AndAlso objWaitForAppear=True) Then
					.WaitForObject(strCtrlName, dblNumSecondsToWait*1000)
				Else If String.Equals(objWaitForAppear.GetType.ToString, "System.Boolean") AndAlso objWaitForAppear=False AndAlso .Exists(strCtrlName) Then
					.TestObject(strCtrlName).WaitForDisappearance(dblNumSecondsToWait*1000)
				End If
					
			'Else, the user gave us a time, ctrl name, and either strPropertyName or strPropertyValue.  If he gave us all the params, then we're waiting for an object's property to appear/disappear.
			Else If String.Equals(objWaitForAppear.GetType.ToString, "System.Boolean") AndAlso Not strPropertyName Is Nothing AndAlso Not strPropertyValue Is Nothing Then
				
				'If the user is waiting for the property to appear
				If objWaitForAppear Then
					
					'We could simply call TestObject.WaitForProperty, but this doesn't handle cases where the property name has invalid case (uppercase/lowercase)
					'It also does not account for instances where the return value is blank but the user probably wants a real value.  We switch between text and textContents sometimes.
					Do While String.Compare(strPropertyValue, FnGetObjectProperty(.TestObject(strCtrlName), strPropertyName, dicConstants, strTemp), True)<>0 And dblNumSecondsToWait>0
						SubSleep(1)
						dblNumSecondsToWait=dblNumSecondsToWait-1
					Loop
					
					'If the object never had that property, return a failure
					If dblNumSecondsToWait <= 0 AndAlso String.Compare(strPropertyValue, FnGetObjectProperty(.TestObject(strCtrlName), strPropertyName, dicConstants, strTemp), True)<>0 Then
						FnWait = "After waiting " & dlbOriginalWaitTime & " seconds, the '" & strPropertyName & "' property of the object defined by '" &  strCtrlName & "' never had a value of '" & _
								 strPropertyValue & "'.  Instead, the value remained '" & FnGetObjectProperty(.TestObject(strCtrlName), strPropertyName, dicConstants, strTemp) & "'."
					End If
					
				'There is no native "wait for property to disappear" method.  We are inventing it here.
				Else					
					Do While String.Compare(FnGetObjectProperty(.TestObject(strCtrlName), strPropertyName, dicConstants, FnWait), strPropertyValue, True)=0 And dblNumSecondsToWait>0 And String.IsNullOrEmpty(FnWait)
						SubSleep(2)
						dblNumSecondsToWait = dblNumSecondsToWait-2						
					Loop
					If String.IsNullOrEmpty(FnWait) And dblNumSecondsToWait <= 0 Then
						FnWait = "After waiting " & dlbOriginalWaitTime & " seconds, the '" & strPropertyName & "' property of the object defined by '" & _ 
								 strCtrlName & "' remained at a value of '" & FnGetObjectProperty(.TestObject(strCtrlName), strPropertyName, dicConstants, strTemp) & "'."
					End If
				End If
			
			Else
				FnWait = "You have supplied an invalid combination of parameters to FnWait."
			End If
			
		End With
	
	End Function	

'==============================================================================================================
' FnEquals
'
' 		Compares two values.  Returns string as error message.
'
'		@strFirstValue:		First value to compare
'		@strSecondValue:	Second value to compare
'		@blnCaseSensitive:	If TRUE, the comparison will be case-sensitive
'==============================================================================================================

	Public Function FnEquals(ByVal strFirstValue As String, ByVal strSecondValue As String, ByVal blnCaseSensitive As Boolean) As String
		
		'Set default return value
		FnEquals = ""
			
		'If the strings are not equal, fail the test
		If String.Compare(strFirstValue, strSecondValue, If(blnCaseSensitive, False, True)) <> 0 Then
			FnEquals = "FnEquals has found the values are not equal: """ & strFirstValue & """ != """ & strSecondValue & """."
		End If
		
	End Function
	
	
'==============================================================================================================
' FnNotEquals
'
' 		Compares two values for inequality.  Returns string as error message.
'
'		@strFirstValue:		First value to compare
'		@strSecondValue:	Second value to compare
'		@blnCaseSensitive:	If TRUE, the comparison will be case-sensitive
'==============================================================================================================

	Public Function FnNotEquals(ByVal strFirstValue As String, ByVal strSecondValue As String, ByVal blnCaseSensitive As Boolean) As String
		
		'Set default return value
		FnNotEquals = ""
		
		'If the strings are equal, fail the test
		If String.Compare(strFirstValue, strSecondValue, If(blnCaseSensitive, False, True)) = 0 Then
			FnNotEquals = "FnNotEquals has found the values are equal: """ & strFirstValue & """ = """ & strSecondValue & """."
		End If

	End Function
	
	
'==============================================================================================================
' FnVerify
'
' 		Verifies that a browser object either exists or does not exist.  Given more than the object's control 
'		name, the function verifies that the objects exists with the specified attribute value.  Returns string
'		containing error message if an error is encountered.
'
'		@strControlName:	Xpath to the object in question
'		@blnInvertedTest:	If TRUE, verifies object/property exists.  If FALSE, verifieds object/property does not exist.
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'		@strPropertyName:	Name of the property we want to verify either exists or does not exist
'		@strExpectedValue:	Value of the given property we want to verify either exists or does not exist
'		@blnMatchWholeWord:	If TRUE, the property value comparision will need to match every character and its case.
'							If FALSE, the property value need only contain the expected value regardless of case.
'==============================================================================================================
							 
	Public Function FnVerify(ByVal strControlName As String, _
					 		 ByVal blnInvertedTest As Boolean, _
							 ByRef dicConstants As Dictionary(Of String, String), _
					 		 Optional ByVal strPropertyName As String=Nothing, _
					 		 Optional ByVal strExpectedValue As String=Nothing, _
					 		 Optional ByVal blnMatchWholeWord As Boolean=False)
		
		'Declare objects
		Dim strPropertyTypesList As List(Of String)
		Dim testObj As TestObject
		
		'Declare variables
		Dim strActualValue As String
		Dim intTemp As Integer
		
		'Set objects
		strPropertyTypesList = New List(Of String)
		
		'Set variables
		FnVerify = ""
		strActualValue = ""
		intTemp = 0
		
		With _desktop
		
			'If the -user- did not supply a second argument, then it's a simple existance check
			If (String.IsNullOrEmpty(strPropertyName)) Then
				
				'If the test is NOT meant to be inverted and the object does not exist
				If(blnInvertedTest = False) Then
					If(.exists(strControlName) = False) Then
						FnVerify = "Object check failed.  Cannot find '" & strControlName & "'."
					End If
					
				'If the test IS meant to be inverted but the object actually exists
				Else
					If(.exists(strControlName) = True) Then
						FnVerify = "Object check failed.  Object named '" & strControlName & "' is present when it's not supposed to be."
					End If
				End If
				
			'The user gave us more than 1 parameter, so we are verifying a property of some sort
			Else
				
				'Retreive the object
				testObj = .TestObject(strControlName)
				
				'Get the actual property value
				strActualValue = FnGetObjectProperty(testObj, strPropertyName, dicConstants, FnVerify)
				
				'Assuming we haven't already failed due to the specified property not existing for this object type
				If(String.IsNullOrEmpty(FnVerify)) Then
					
					'If the user is trying to get a number of rows or columns value, set the tolerance to a strict check
					If(String.Compare(strPropertyName, "numRows", True)=0 OrElse String.Compare(strPropertyName, "numColumns", True)=0)
						blnMatchWholeWord = True
					End If
					
					'If the user wanted a strict check
					If(blnMatchWholeWord) Then
						If(String.Compare(strActualValue, strExpectedValue, True) <> 0) Then
							FnVerify = "Check failed.  '" & strPropertyName & "' property of '" & strControlName & "' is not strictly '" & strExpectedValue & "'.  Actual value: '" & strActualValue & "'"
						End If
					
					'If the user wants a loose check
					Else If(ci.CompareInfo.IndexOf(strActualValue, strExpectedValue, CompareOptions.IgnoreCase) = -1) Then
						FnVerify = "Check failed.  '" & strPropertyName & "' property of '" & strControlName & "' does not contain '" & strExpectedValue & "'.  Actual value: '" & strActualValue & "'."
					End If
					
				End If
				
			End If
						
		End With
		
	End Function
	
	
'==============================================================================================================
' FnRetain
'
' 		Obtains the value of a specified browser object's property and stores it in the given variable reference.
'		Returns string containing error message if necessary.
'
'		@strControlName:	Xpath to object in question
'		@strPropertyName:	Name of the property to retrieve value from
'		@strVarRef:			The referenced variable in which to store the derived property value
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'		@intWordNum:		The particular word number to retrieve.  Ex: "2" from "Hello World!" returns "World!"
'==============================================================================================================

	Public Function FnRetain(ByVal strControlName As String, _
							 ByVal strPropertyName As String, _
							 ByRef strVarRef As String, _
							 ByRef dicConstants As Dictionary(Of String, String), _
							 Optional ByVal intWordNum As String=Nothing) As String
	
		'Objects
		Dim testObj As TestObject
		
		'Variables
		Dim strArr() As String
				
		'Set things
		FnRetain = ""
		
		With _desktop
			
			'Get the object
			testObj = .TestObject(strControlName)
			
			'Get the property and set an error message if one comes back
			strVarRef = FnGetObjectProperty(testObj, strPropertyName, dicConstants, FnRetain)
				
			'If the user wants a particular word and there's no error thus far
			If(String.IsNullOrEmpty(FnRetain) AndAlso Not intWordNum Is Nothing) Then
				
				'Break up the property by the space character
				strArr = Split(strVarRef, " ")		
				
				'Get the requested word
				If(strArr.length >= intWordNum)
					strVarRef = strArr(intWordNum-1)
				Else
					FnRetain = "FnRetain cannot retrieve the string as position " & intWordNum & " because the entire property only contains " & strArr.length & " words."
				End If
				
			End If
			
		End With
		
	End Function
	

'==============================================================================================================
' FnGetObjectProperty
'
' 		Simply retrieves and returns a test object's property.  There is a several good reasons that we use this
'		function instead of calling the property directly on the given test object.
'
'		[1] If the user asks for the "text" property, Silk may throw an exception because often the "text" property
'			does not exist; it is case-sensitivate and should sometimes be "Text".  
'		[2] The user may be asking for the "Text" property, but the given object doesn't have that property; instead,
'			it  has "textContents".  It's easier to correct this error automatically than throw an exception to the user. 
'		[3] I have invented properties that I believe ought to exist for particular objects.  Namely, the "numrows"
'			and "numColumns" properties ought to exist for the HTML table object, but they do not.  This function
'			makes up for that inadequacy by seamlessly making it appear to be a property to the end-user.  
'		[4] Oftentimes the retrieved property value has non-standard space characters in it.  These characters 
'			look like spaces but are programmatically something else.  This can be a real issue when the user 
'			tries to verify a property and is comparing it against a string they typed in themselves.  This is 
'			also an issue for Xscript's Retain function because it sometimes needs to split a property on space 
'			characters.  
'		[5] We need to remove line breaks from retrieved properties or order to compare the strings elsewhere.  
'		[6] If the user is trying to  get the "Text" or "textContents" property of a 'Control' object in MedCheck, 
'			sometimes it will return blank regardless of its value.  This function runs OCR on the object to find
'			its value in such cases.  
'		[7] The 'State' and 'Selected' properties of 'Control' objects do not actually exist, but Xscript allows
'			you to retrieve them as if they do exist.  FnGetObjectProperty is doing some tricky key combinations
'			and image recognition to return the appropriate values here.
'
'		@testObj:			The actual test object from which to retrieve the property
'		@strPropertyName:	The name of the property the user wants to get the value of
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'		@strErrorMessage:	String variable passed ByRef that houses error messages should it been necessary
'==============================================================================================================

	Private Function FnGetObjectProperty(ByRef testObj As TestObject, ByVal strPropertyName As String, ByRef dicConstants As Dictionary(Of String, String), ByRef strErrorMessage As String) As String
	
		'Declare
		Dim strPropertyList As List(Of String)
		Dim strObjectType As String
		Dim strTextPropertyName As String = Nothing
		Dim strTextContentsPropertyName As String = Nothing
		Dim strTemp As String=Nothing
		Dim bmp1 As Bitmap
		
		'Set values
		FnGetObjectProperty = ""
		strObjectType = testObj.GetType.ToString
		strPropertyList = testObj.GetPropertyList
			
		With _desktop
			
			'Cycle through all the available properties for this object.  If a match is found, set it to the match because it will have the correct upper and lower casing.
			For Each strObjProperty As String In strPropertyList
				If String.Compare(strPropertyName, strObjProperty, True)=0 Then strPropertyName = strObjProperty
				If String.Compare("text", strObjProperty, True)=0 Then strTextPropertyName = strObjProperty
				If String.Compare("textcontents", strObjProperty, True)=0 Then strTextContentsPropertyName = strObjProperty
			Next strObjProperty	
									
			'Ensure the requested property type actually exists for this object
			If(strPropertyList.Contains(strPropertyName)) Then
				
				'Fetch the property value
				FnGetObjectProperty = testObj.GetProperty(strPropertyName)
				
				'Sometimes words that are separated by spaces are actually separated by "&nbsp;".  We're fixing that here.
				FnGetObjectProperty = FnReplaceImproperSpaces(FnGetObjectProperty)
				
				'We also need to swap out line breaks for spaces
				FnGetObjectProperty = FnGetObjectProperty.Replace(vbNewLine, " ")
				
				'If the value is blank and the user is asking for either the 'text' or 'textContents' property, reverse values to try again
				If String.IsNullOrEmpty(FnGetObjectProperty) Then
					If String.Compare(strPropertyName, "text", True)=0 AndAlso Not strTextContentsPropertyName Is Nothing Then FnGetObjectProperty = testObj.GetProperty(strTextContentsPropertyName)
					If String.Compare(strPropertyName, "textContents", True)=0 AndAlso Not strTextPropertyName Is Nothing Then FnGetObjectProperty = testObj.GetProperty(strTextPropertyName)
				End If
				
				'If the user is looking for 'text' or 'textContents, but so far the value is STILL blank, and this is a control object, run OCR on it.
				'If the result of the OCR ends in the '6' character, this a medcheck dropdown.  Use OCR with trimming to get selected value.
				If String.IsNullOrEmpty(FnGetObjectProperty) AndAlso String.Compare("SilkTest.Ntf.Control", strObjectType, True)=0 AndAlso _
				   (String.Compare(strPropertyName, "text", True)=0 OrElse String.Compare(strPropertyName, "textContents", True)=0) Then
				
					'Do an OCR on the control object
					strTemp = testObj.TextCapture
				
					'If the string ends in a '6', then this is very likely a dropdown menu and we need to trim the retrieved value appropriately
					If String.Equals(Right(strTemp, 1), "6") Then
						FnGetObjectProperty = Left(strTemp, strTemp.Length-1).Trim
					End If
					
				End If
				
			'If the user is trying to get 'numRows' or 'numColumns" property, we invent it here
			Else If(String.Compare("numRows", strPropertyName, True)=0 OrElse String.Compare("numColumns", strPropertyName, True)=0) Then
			
				'If the user is trying to get the numRows property
				If String.Compare("numRows", strPropertyName, True)=0
									
					'Get the number of rows depending on the object type
					If(String.Compare("SilkTest.Ntf.XBrowser.DomTable", strObjectType, True)=0)
						FnGetObjectProperty = CType(testObj, DomTable).GetRowCount
					Else If(String.Compare("SilkTest.Ntf.Table", strObjectType, True)=0)
						FnGetObjectProperty = CType(testObj, Table).RowCount
					End If
				
				'If the user is trying to get the numColumns property
				Else
					
					'Get the number of columns depending on the object type
					If(String.Compare("SilkTest.Ntf.XBrowser.DomTable", strObjectType, True)=0)
						FnGetObjectProperty = CType(testObj, DomTable).GetColumnCount
					Else If String.Compare("SilkTest.Ntf.Table", strObjectType, True)=0
						FnGetObjectProperty = CType(testObj, Table).ColumnCount
					End If
					
				End If
			
			'If the user is trying to get the 'state' property of a 'control' object, it's safe to assume that the object is a checkbox
			Else If(String.Compare("state", strPropertyName, True)=0 AndAlso String.Compare(strObjectType, "SilkTest.Ntf.Control", True)=0) Then
			
				'Take a screenshot of the object
				bmp1 = FnCaptureBitmap(CType(testObj, Control), dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\Temp.bmp", True)
				
				'If the enabled checked checkbox image is contained OR the disabled unchecked checkbox image is not contained, return "1", else return "2"
				'The reason we're not directly checking for the "disabled checked checkbox" is that I couldn't find an example of this scenario in medcheck.  This will do fine though.
				If(     (   CType(testObj, Control).Enabled AndAlso FnImageContains(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckCheckboxEnabledSelectedFileName")))   )    OrElse _						
						(   CType(testObj, Control).Enabled=False AndAlso Not FnImageContains(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckCheckboxDisabledUnselectedFileName")))   )    ) Then
				
					FnGetObjectProperty = "1"
				Else
					FnGetObjectProperty = "2"
				End If
			
			'Else if the user is trying to get the 'selected' property of a 'control' object, it's safe to assume that the object is a radio button
			Else If(String.Compare("selected", strPropertyName, True)=0 AndAlso String.Compare(strObjectType, "SilkTest.Ntf.Control", True)=0) Then
				
				'Take a screenshot of the object
				bmp1 = FnCaptureBitmap(CType(testObj, Control), dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\Temp.bmp", True)
				
				'If either the enabled or disabled radiobutton images are contained in the screenshot, return "true", else return "false"
				If( (String.Equals(CType(testObj, Control).Enabled, "true") And FnImageContains(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckRadioButtonEnabledSelectedFileName")))) OrElse _
					(String.Equals(CType(testObj, Control).Enabled, "false") And FnImageContains(bmp1, New Bitmap(dicConstants.Item("HelperImagesPath") & "\" & dicConstants.Item("MedCheckRadioButtonDisabledSelectedFileName")))) ) Then
					FnGetObjectProperty = "True"
				Else
					FnGetObjectProperty = "False"
				End If
				
			Else
				strErrorMessage = "FnGetObjectProperty has found that there is no '" & strPropertyName & "' property for the specified object type '" & testObj.GetType.ToString & "'."
			End If
			
		End With
		
	End Function
	
	
'==============================================================================================================
' FnDismissDownload
'
' 		This function simply dismisses the file download dialog/notification in a browser.  It is called in 
'		Xscript by the user sending the "DismissDownload" command.  The reason the user doesn't simply send a
'		command to click on the "Close" button or "X" is because file download dialogs/notifications are different
'		between browsers and browser versions.  In order to allow Xscript tests to be cross-browser compatible,
'		this function handles each browser type.
'
'		Note: As of 8/6/2014, this function has only been tested with IE8 and IE9.
'==============================================================================================================

	Public Function FnDismissDownload() As String
		
		'Set default return value
		FnDismissDownload = ""
		
		'If the 'File Download' prompt is on the screen
		If _desktop.Exists("//Dialog[@caption='File Download']//PushButton[@caption='Cancel']") Then
			_desktop.PushButton("//Dialog[@caption='File Download']//PushButton[@caption='Cancel']").Click
		End If
		
		'If the IE9+ 'notification bar' is on the screen
		If(_desktop.BrowserApplication.Exists("//Control[@windowClassName='DirectUIHWND'][2]")) Then			
			
			'We have to toggle this option because it allows Silk to interact with IE9+ dialog box child elements
			Agent.SetOption(Options.EnableAccessibility, True)
		
			'Close the dialog box
			_desktop.BrowserApplication.PushButton("[@caption='Close']").Click
			
			'Toggle the setting off because according to Silk documentation, it slows down playback
			Agent.SetOption(Options.EnableAccessibility, False)
			
		End If
		
	End Function
	
	
'==============================================================================================================
' FnSaveDownload
'
' 		This function simply handles the file download dialog/notification in a browser by saving off the file.
'		It is called in Xscript by the user sending the "SaveDownload" command.  The reason the user doesn't 
'		instead send a series of commands to save it off as he sees fit is because file download dialogs/notifications
'		are different between browsers and browser versions.  In order to allow Xscript tests to be cross-browser
'		compatible, this function handles the operation for each browser type automatically.
'
'		Note: The function will overwrite pre-existing files.  Additionally, as of 8/6/2014, this has only been
'		tested with IE8 and IE9.
'
'		@strFullFilePath:	The full file path in which to save the file
'==============================================================================================================

	Public Function FnSaveDownload(ByVal strFullFilePath As String) As String
			
		'Set default return value
		FnSaveDownload = ""
		
		With _desktop
		
			'If the 'File Download' prompt is on the screen
			If .Exists("//Dialog[@caption='File Download']//PushButton[@caption='Save']") Then
				
				'Click the Save button
				.PushButton("//Dialog[@caption='File Download']//PushButton[@caption='Save']").Click
				
				'Type in the file path
				.Dialog("[@caption='Save As']").TextField("[@caption='File name*']").SetText(strFullFilePath)
				
				'Click the Save button
				.Dialog("[@caption='Save As']").PushButton("[@caption='Save']").Click
				
				'If the overwrite prompt pops up, click Yes
				If .Exists("//Dialog[@caption='Save As']//Dialog[@caption='Save As']//PushButton[@caption='Yes']") Then
					.PushButton("//Dialog[@caption='Save As']//Dialog[@caption='Save As']//PushButton[@caption='Yes']").Click
				End If
				
				'Wait up to 2 minutes for the 'Download Complete' dialog box to appear
				.WaitForObject("//Window[@caption='Download complete']", 60 * 1000 * 2)
				
				'Click the Close button
				.Window("[@caption='Download complete']").PushButton("[@caption='Close']").Click
				
			End If
			
			'If the IE9+ 'notification bar' is on the screen
			If(.BrowserApplication.Exists("//Control[@windowClassName='DirectUIHWND'][2]")) Then
				
				'We have to toggle this option because it allows Silk to interact with IE9+ dialog box child elements
				Agent.SetOption(Options.EnableAccessibility, True)			
				
				'Click the 'Save dropdown arrow'
				.Control("[@caption='Notification bar']").AccessibleControl("[@role='drop down button']").Click	
				
				'Click the 'Save As' menu item
				.Control("[@caption='Notification bar']").MenuItem("[@caption='Save as']").Click
				
				'Type the name of the file path into the dialog box
				.BrowserApplication.Dialog.TextField.SetText(strFullFilePath)
				
				'Click the save button
				.BrowserApplication.Dialog.PushButton("[@caption='Save']").Click
				
				'If you're prompted to overwrite a file, just do it
				If .BrowserApplication.Exists("//Dialog[@caption='Confirm Save As']") Then
					 .BrowserApplication.Dialog("//Dialog[@caption='Confirm Save As']").PushButton("//PushButton[@caption='Yes']").Click
				End If
				
				'Close the lingering notification bar
				If(.Exists("//Control[@caption='Notification bar']")) Then
					.Control("[@caption='Notification bar']").PushButton("[@caption='Close']").Click
				End If
				
				'Toggle the setting off because according to Silk documentation, it slows down playback
				Agent.SetOption(Options.EnableAccessibility, False)
				
			End If
			
		End With
		
	End Function
	
	
'==============================================================================================================
' FnSQL
'
' 		Executes SQL query against given Server/DB.  If specified, query result is stored in the given reference
'		to a simple data type, dictionary of string, or 2D associative array (AssociativeArray2D).  
'		Returns string containing error message.
'
'		NOTE: When passing an data type to retain the result, be sure you give the variable a default value before
'		passing it.  When passing a dictionary or 2D associative array, be sure you instantiate it with the 'New'
'		keyword before passing it.  Failure to do so will result in an error because the function cannot determine
'		what type of variable/object is being passed to it.
'
'		NOTE: It is sometimes necessary to use double-quotes around column and table names that would otherwise
'		reference a keyword in SQL.  You will need to escape the double-quote character in your query by duplicating
'		the double-quote.  See examples below.
'
' 		@strQuery:				Query to execute
'		@strServer:				Server to use.  ex: ScanODCTestSQL1
'		@strDB:					Database to use.  ex: ODCManager
'		@objToPopulate:			Reference to variable or object in which to store result
'		@strKeyByThisColumn:	Column in which to key each row returned
'
'		'Example #1 - No retention
'			strQuery = "Update ""User"" Set ""FailedLoginAttempts"" = 5 WHERE ""Username"" = 'Corveltest'"
'			strErrorMsg = FnSQL(strQuery, "ScanODCTestSQL1", "ODCManager")
'		
'		'Example #2 - Simple data type retention
'			Dim strAnswer As String
'			strAnswer = ""
'			strQuery = "select ""FailedLoginAttempts"" From ""User"" WHERE ""Username"" = 'Corveltest'"
'			strErrorMsg = FnSQL(strQuery, "ScanODCTestSQL1", "ODCManager", strAnswer)
'			Console.Writeline("strAnswer = " & strAnswer)
'		
'		'Example #3 - Retain multiple columns from a single row
'			Dim myDic As Dictionary(Of String, String)
'			myDic = New Dictionary(Of String, String)
'			strQuery = "select * From ""User"" WHERE ""Username"" = 'Corveltest'"
'			strErrorMsg = FnSQL(strQuery, "ScanODCTestSQL1", "ODCManager", myDic)
'			Console.Writeline("The user has failed to login " & myDic.Item("FailedLoginAttempts") & " times."
'			For Each kvp As KeyValuePair(Of String, String) In myDic 
'				Console.WriteLine("myDic(" & kvp.Key & ") = " & kvp.Value)
'			Next
'		
'		'Example #4 - Retain multiple rows keyed by a particular column
'			Dim myTDAA As AssociativeArray2D
'			myTDAA = New AssociativeArray2D
'			strQuery = "select * From ""User"""
'			strErrorMsg = FnSQL(strQuery, "ScanODCTestSQL1", "ODCManager", myTDAA, "UserName")	
'			Console.Writeline("The 'Corveltest' user has failed to login " & myTDAA.Item("Corveltest", "FailedLoginAttempts") & " times."
'			For Each key1 As String In myTDAA.Keys 
'				For Each key2 As String In myTDAA.Keys(key1)
'					Console.Writeline("myTDAA('" & key1 & "', '" & key2 & "') = " & myTDAA.Item(key1, key2))
'				Next
'			Next
'==============================================================================================================

	Public Function FnSQL(ByVal strQuery As String, _
				  		  ByVal strServer As String, _
				  		  ByVal strDB As String, _
				  		  Optional ByRef objToPopulate As Object = "8675309-12031982", _
				  		  Optional ByVal strKeyByThisColumn As String="") As String
								
		'Objects
		Dim sqlConn As SqlConnection
		Dim sqlCmd As SqlCommand
		Dim sqlDr As SqlDataReader		
		
		'Variables
		Dim strConnectionString As String
		Dim strColumnValue As String
		Dim intRowCounter As Integer		
		Dim intColumnCounter As Integer
		Dim strSpecifiedKeyValue As String
		Dim strRowKeyToUse As String
		Dim blnUserWantsSingularValue As Boolean
		Dim blnUserWantsDictionary As Boolean
		Dim blnUserWantsTDAA As Boolean
		
		'Set variables
		FnSQL = ""
		strColumnValue = ""
		strSpecifiedKeyValue = ""
		strRowKeyToUse = ""
		intRowCounter = -1
		intColumnCounter = -1
		
		'Form the connection string
		strConnectionString = "server=" & strServer & ";database=" & strDB & ";Trusted_Connection=True;MultipleActiveResultSets=True"
		
		'Connect to and open the database
		sqlConn = New SqlConnection(strConnectionString) 
		sqlConn.Open()
								
		'Form command
		sqlCmd = New SqlCommand(strQuery, sqlConn)
				
		'Execute query
		sqlDr = sqlCmd.ExecuteReader
		
		'Determine the type of object that was sent to us for retention
		If 	TypeOf objToPopulate Is String Or TypeOf objToPopulate Is Boolean Or _
		   	TypeOf objToPopulate Is Double Or TypeOf objToPopulate Is Integer Or _
			TypeOf objToPopulate Is Long Or TypeOf objToPopulate Is Date Or _
			TypeOf objToPopulate Is Decimal Or TypeOf objToPopulate Is Char Then
			If(objToPopulate.ToString <> "8675309-12031982") Then
				blnUserWantsSingularValue = True
			End If
		Else If TypeOf objToPopulate Is Dictionary(Of String, String) Then
			blnUserWantsDictionary = True
		Else If TypeOf objToPopulate Is AssociativeArray2D Then
			blnUserWantsTDAA = True
		Else
			FnSQL = "FnSQL either cannot determine the type of object/variable you sent or is not configured to handle the the particular type you sent.  Ensure you have initialized the object or data by using the 'New' keyword or giving the data type a default value."
		End If
		
		'If the user wants a return value
		If(blnUserWantsSingularValue Or blnUserWantsDictionary Or blnUserWantsTDAA) Then
			
			'Loop over every row until we've hit and error -OR- we're out of rows -OR- the user only wanted a string and it's now populated -OR- the user only wanted a dictionary and it's now populated
			Do Until String.IsNullOrEmpty(FnSQL) = False Or sqlDr.Read() = False Or (blnUserWantsSingularValue = True And intRowCounter > -1) Or (blnUserWantsDictionary = True And intRowCounter > -1)
				
				'Increment the row counter
				intRowCounter = intRowCounter + 1
				
				'Reset variables
				strSpecifiedKeyValue = ""
				intColumnCounter = -1
				
				'Loop over every column until we're out of columns -OR- the user only wanted a string and now it's populated
				Do Until intColumnCounter = sqlDr.FieldCount-1 Or (blnUserWantsSingularValue = True And intColumnCounter > -1)
					
					'Increment the column counter
					intColumnCounter = intColumnCounter + 1
									
					'In case the value in this column is null, we are going to test for it
					If(sqlDr.GetValue(intColumnCounter) Is DbNull.Value)
						strColumnValue = ""
					Else
						strColumnValue = sqlDr.GetValue(intColumnCounter)
					End If
									
					'If the user is only asking for a string, populate the supplied variable with the first value from the first row
					If(blnUserWantsSingularValue) Then 
						objToPopulate = strColumnValue
					
					'If the user is asking for a dictionary, add the first row to the dictionary keyed by column name
					Else If(blnUserWantsDictionary) Then
						objToPopulate.Add(sqlDr.GetName(intColumnCounter), strColumnValue)
						
					'If the user wants a 2D associatiave array
					Else If(blnUserWantsTDAA) Then
						
						'If we don't care to key our 2D associative array by a particular column value OR we already obtained the key
						If(String.IsNullOrEmpty(strKeyByThisColumn) Or String.IsNullOrEmpty(strSpecifiedKeyValue) = False) Then
							
							'Determine the value for this row's key
							If(String.IsNullOrEmpty(strSpecifiedKeyValue)) Then					
								strRowKeyToUse = intRowCounter.ToString
							Else
								strRowKeyToUse = strSpecifiedKeyValue
							End If
							
							'Add the value to our persistent dictionary of 2D associative arrays
							objToPopulate.Add(strRowKeyToUse, sqlDr.GetName(intColumnCounter), strColumnValue)
						
						'But if we do want to key our 2D associative array by a particular column AND this is the column
						Else If(String.Compare(sqlDr.GetName(intColumnCounter), strKeyByThisColumn, True) = 0) Then
							
							'Save the value of this special column
							strSpecifiedKeyValue = strColumnValue
							
							'Reset the column counter so we loop through the whole row all over again
							intColumnCounter = -1
						
						End If
						
					End If
					
				Loop
				
			Loop
			
		End If
				
		'Close the database
		sqlConn.Close()
		
		'Close the connection
		sqlConn.Close
	
	End Function


'==============================================================================================================
' FnMath
'
' 		Performs the specified mathematical operation between two values and stores the answer in the referenced
'		variable.  Returns string containing error message. 
'
'		Note: It may be possible for .NET to do an Eval() on a mathemtically expression instead.  That would be
'		far more powerful than anything I'd ever write.  See the link below.
'		http://stackoverflow.com/questions/1452282/doing-math-in-vb-net-like-eval-in-javascript
'		Example: Dim result = New DataTable().Compute("3+(7/3.5)", Nothing)
'
' 		@dblFirstNum:		The number to the left of the operand
'		@dblSecondNum:		The number to the right of the operand
'		@strOperation:		The operation to perform. ex: add, subtract
'		@strAnswer:			The reference variable in which to store the answer
'==============================================================================================================

	Public Function FnMath(ByVal dblFirstNum As Double, ByVal dblSecondNum As Double, ByVal strOperation As String, ByRef strAnswer As String) As String

		'Set default return value
		FnMath = ""
		
		Select Case strOperation.ToLower
			Case "add", "plus"
				strAnswer = dblFirstNum + dblSecondNum
			Case "subtract", "minus"
				strAnswer = dblFirstNum - dblSecondNum
			Case "multiply", "multiplied", "times"
				strAnswer = dblFirstNum * dblSecondNum
			Case "divide", "divided", "over", "by"
				strAnswer = dblFirstNum / dblSecondNum
			Case Else
				FnMath = "The FnMath function is not yet configured to perform the '" & strOperation & "' operation."
		End Select
		
	End Function
	
	
'==============================================================================================================
' FnGetFilePaths
'
' 		Returns a string List of filenames in the given directory.  If specified, the function will return only
'		filenames with a particular extension.  Can also include or omit hidden files if specified.
'
' 		@strDir:					The directory in which to look for files
'		@strExtension:				The type of file extensions to look for.  ex: xlsx
'		@blnIncludeHiddenFiles:		Boolean to indicate whether to include hidden files
'==============================================================================================================

	Public Function FnGetFilePaths(ByVal strDir As String, Optional ByVal strExtensions As String()=Nothing, Optional ByVal blnIncludeHiddenFiles As Boolean=False,) As List(Of String)

		'Declare objects
		Dim myDir As DirectoryInfo
		Dim myFI As FileInfo
		
		'Set objects
		myDir = New DirectoryInfo(strDir)
		FnGetFilePaths = New List(Of String)
		
		'If the user gave us an array of extensions, ensure each value begins with '.'
		If Not strExtensions Is Nothing Then
			For i As Integer = 0 To strExtensions.Length-1
				If Not String.Equals(Left(strExtensions(i), 1), ".") Then strExtensions(i) = "." & strExtensions(i)
			Next i
		End If
		
		'Loop through each file in the directory
		For Each myFI In myDir.GetFiles
			
			'If the file is visible or the user doesn't care
			If(blnIncludeHiddenFiles) OrElse ((myFI.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden)  Then
				
				'If the file extension matches or the user doesn't care
				If strExtensions Is Nothing OrElse Array.IndexOf(strExtensions, myFi.Extension) > -1 Then
					FnGetFilePaths.Add(myFI.FullName)
				End If
				
			End If        
		Next
	
	End Function
	
	
'==============================================================================================================
' FnParseInstruction
'
' 		This function parses the given Xscript instruction and populates the referenced command and the referenced
'		parameter list with the proper string values.  Should an error occur, the function itself will return
'		an error message string.
'
'		@strInstruction:	A full Xscript instruction
'		@strCmd:			The referenced string variable to populate with the derived command
'		@strParamList:		The referenced string list to populate with each derived paramater
'==============================================================================================================

	Public Function FnParseInstruction(ByVal strInstruction As String, ByRef strCmd As String, ByRef strParamList As List(Of Object)) As String

		'Declarations
		Dim intLeftParanIndex, intRightParanIndex As String
		Dim strUnsplitParams As String
		
		'Set defaults returns
		FnParseInstruction = ""
		strCmd = Nothing
		strParamList = New List(Of Object)
	
		'Determine the location of each parentheses
		intLeftParanIndex = strInstruction.IndexOf("(")
		intRightParanIndex = strInstruction.LastIndexOf(")")
		
		'Determine what the command and parameter values are
		If(intLeftParanIndex <> -1)
			strCmd = Left(strInstruction, intLeftParanIndex)
			strUnsplitParams = strInstruction.Substring(intLeftParanIndex+1, intRightParanIndex - intLeftParanIndex - 1)
		Else
			strCmd = strInstruction
			strUnsplitParams = Nothing
		End If
		
		'If there are parameters for this instruction
		If Len(strUnsplitParams) > 0 Then
			
			'Split the parameters by the comma character into a list.  The regex ensures that it only splits on commas outside of double-quotes.
			strParamList.AddRange(Regex.Split(strUnsplitParams, ",(?=(?:[^\""]*\""[^\""]*\"")*[^\""]*$)"))
							
			'Trim each value
			For i As Integer = 0 To strParamList.Count-1
				strParamList(i) = strParamList(i).Trim
			Next i
			
		End If
	
	End Function
	
	
'==============================================================================================================
' FnReverseSortDictionaryKeys
'
' 		Sorts the given dictionary of string/string by its keys.  The biggest keys will come first while the 
'		shortest keys will come last.  There is no native .NET method that does this.
'
'		This function is used by Xscript Excel during the FnUpdateObjectReferences function.  Because that 
'		function utilizes Excel's Replace routine, it's important that I start by finding and replacing the 
'		longest strings first so that shorter strings don't accidentally match part of a longer string and 
'		swap it out before the legitimate replacement can process it.
'
'		@strInputDic:	The dictionary of string/string to sort by its keys in reverse length
'==============================================================================================================

	Public Function FnReverseSortDictionaryKeys(ByRef strInputDic As Dictionary(Of String, String)) As Dictionary(Of String, String)
	
		'Declarations
		Dim intIndex, intMax, intIndexOfLongest As Integer
		Dim strKeys As New List(Of String)(strInputDic.Keys)
		
		'Initialize return object
		FnReverseSortDictionaryKeys = New Dictionary(Of String, String)
		
		'Cycle through the given dictionary until we've assembled every item into the return dictionary
		Do While FnReverseSortDictionaryKeys.Count < strInputDic.Count
						
			'If we haven't already added this key to the new dic, examine it for length
			If Not FnReverseSortDictionaryKeys.ContainsKey(strKeys.Item(intIndex)) Then
				
				'If the length of the key is greater than our current champion
				If Len(strKeys.Item(intIndex)) > intMax Then
					intIndexOfLongest = intIndex
					intMax = Len(strKeys.Item(intIndex))
				End If
				
			End If
			
			'If this is the end of the cycle, add the longest key we found to the dictionary
			If intIndex = strKeys.Count-1
				FnReverseSortDictionaryKeys.Add(strKeys.Item(intIndexOfLongest), strInputDic.Item(strKeys.Item(intIndexOfLongest)))
				intMax = 0
				intIndex = 0
			Else
				intIndex = intIndex + 1
			End If
			
		Loop
	
	End Function
	
	
'==============================================================================================================
' FnStripNonAlphaNumeric
'
' 		Strips all non-alphanumeric characters from a string and returns it.
'
' 		@strValue:	The string to process
'==============================================================================================================

	Public Function FnStripNonAlphaNumeric(ByVal strValue As String)
		
		FnStripNonAlphaNumeric = Regex.Replace(strValue, "[^a-zA-Z0-9]", "")
		
	End Function
	
	
'==============================================================================================================
' FnStripNonNumeric
'
' 		Strips all non-numeric characters from a string and returns it.
'
' 		@strValue:	The string to process
'==============================================================================================================

	Public Function FnStripNonNumeric(ByVal strValue As String)
		
		FnStripNonNumeric = Regex.Replace(strValue, "[^0-9]", "")
		
	End Function
	
	
'==============================================================================================================
' FnNumToLetter
'
' 		Converts a number to a string.  1 -> A, 26 -> Z, 27 -> AB, 53 -> AC, etc
'
' 		@intNum:	The number to convert
'==============================================================================================================

	Public Function FnNumToLetter(intNum As Integer) As String
		
		'Declare vars
		Dim intMod As Integer
		
		'Set vars
		FnNumToLetter = ""
	 
		'Loop once for every letter anticipated 
	    While intNum > 0
	        intMod = (intNum - 1) Mod 26
	        FnNumToLetter = Chr(65 + intMod) & FnNumToLetter
	        intNum = CInt((intNum - intMod) \ 26)
	    End While
	 
	End Function


'==============================================================================================================
' FnReplaceImproperSpaces
'
' 		Replaces 'blanks' with 'spaces'.  This can sometimes be necessary when you retrieve a web object's
'		property that contains &nbsp; and other mysterious blanks instead of normal space characters.  VB.NET 
'		treats them neither as spaces or the simple set of six characters "&nbsp;".  Instead, they will look 
'		likes spaces but not match spaces.  Anyway, this function strips them out for normal spaces.
'
'		@strInput:		The string to clean of blanks
'==============================================================================================================

	Public Function FnReplaceImproperSpaces(ByVal strInput As String) As String
		
		FnReplaceImproperSpaces = ""
		
		For Each charLetter As Char In strInput
			If charLetter.CompareTo(Chr(160)) = 0 OrElse charLetter.GetHashCode = 10485920 Then
				FnReplaceImproperSpaces = FnReplaceImproperSpaces & " "
			Else
				FnReplaceImproperSpaces = FnReplaceImproperSpaces & charLetter
			End If
		Next
		
	End Function

	
'==============================================================================================================
' SubSleep
'
' 		Pauses playback for the specified number of seconds.
'
' 		@dblNumSeconds:		Number of seconds to sleep
'==============================================================================================================

	Public Sub SubSleep(ByVal dblNumSeconds As Double)
		
		If(dblNumSeconds > 0) Then
			System.Threading.Thread.Sleep(dblNumSeconds * 1000)
		End If
		
	End Sub
	
	
'==============================================================================================================
' FnLowercaseList
'
' 		Convert every element in a list of strings to lowercase
'
' 		@listStrings:		The list of strings to lowercase
'==============================================================================================================

	Public Function FnLowercaseList(ByVal listStrings As List(Of String)) As List(Of String)
		
		FnLowercaseList = New List(Of String)
		
		For Each strElement As String In listStrings
            FnLowercaseList.Add(strElement.ToLower)
        Next
		
	End Function

	
'==============================================================================================================
' print
'
' 		Add the given string to a persistent list of strings.  This function is used as an alternative to 
'		Console.Writeline when debugging a script.  There are circumstances where the I need to debug a loop
'		that might cycle a million times, but calling Console.Writeline causes Silk to slow down to a crawl
'		after a while.  So instead, I add each thing I want to output to this list.  When the list is complete,
'		then I call the printout() function below which writes all the contents to a text file.
'
'		@str:	The output to add to the list
'==============================================================================================================
	
	Public Sub print(str)
		strPrintList.Add(str)
	End Sub
	
	
'==============================================================================================================
' printout
'
' 		This writes the contents of the persistent strPrintList C:\Log.txt.  It is used as an alternative to
'		Console.Writeline when we'd have to output a mammoth amount of data to the console.
'==============================================================================================================
	
	Public Sub printout()
		Dim sw As StreamWriter
		sw = My.Computer.FileSystem.OpenTextFileWriter("C:\Log.txt", False)
		For Each str As String In strPrintList
			sw.WriteLine(str)
		Next str
		sw.Close()
	End Sub
	
	
'==============================================================================================================
' FnGetAllFilePathsRecursively
'
' 		Given a directory, it returns full file paths to every contained file in that folder and all subfolders.
'
'		@strRootDir:	The path to the starting directory.  Example: "C:\Chuck\Test"
'==============================================================================================================
		
    Public Function FnGetAllFilePathsRecursively(ByVal strRootDir As String) As List(Of String)
		
		'Declarations
		Dim strFilePathsList As New List(Of String)
		Dim stkDirectories As New Stack(Of String)
		Dim strCurrentDir As String
		
		'Add the root directory to our stack of directories
		stkDirectories.Push(strRootDir)

		'Loop through every available directory
		Do While (stkDirectories.Count > 0)
		
			'Pluck a directory off the stack
			strCurrentDir = stkDirectories.Pop
		
			'Get all files in the current directory
			strFilePathsList.AddRange(Directory.GetFiles(strCurrentDir, "*"))
			
			'Loop through all the folders in the currrent directory and add each to the directory stack
			For Each strDirName As String In Directory.GetDirectories(strCurrentDir)
				stkDirectories.Push(strDirName)
			Next
		
		Loop

		Return strFilePathsList
		
    End Function
	
	
'==============================================================================================================
' SubDeleteEmptyFolders
'
' 		Given a starting folder, this routine deletes any empty folder appearing within it or any of its 
'		subfolders recursively.
'
'		@strRootFolder:					Full path to the starting directory
'		@blnDeleteRootFolderIfEmpty:	If TRUE, the routine will delete the root folder if it's empty
'==============================================================================================================

	Public Sub SubDeleteEmptyFolders(ByVal strRootFolder As String, Optional ByVal blnDeleteRootFolderIfEmpty As Boolean=False)
		
		Try
			
			'Assuming the root folder exists
			If Directory.Exists(strRootFolder) Then
				
				'Loop through each sub folder.
				For Each strCurrentDir As String In Directory.GetDirectories(strRootFolder)					
					SubDeleteEmptyFolders(strCurrentDir, True)
				Next
				
				'If the directory is empty, delete it
				If blnDeleteRootFolderIfEmpty AndAlso Directory.GetFiles(strRootFolder).Length = 0 AndAlso Directory.GetDirectories(strRootFolder).Length = 0 Then
					Directory.Delete(strRootFolder)
				End If
				
			End If
			
		'This catch is to prevent write-locks from haulting the execution of Xscript.  There is always a chance that someone is viewing the file we're trying to delete.
		Catch e As Exception
		End Try
		
	End Sub
	
'==============================================================================================================
' FnAddLeadingZeros
'
' 		Given an integer, this function tacks on zeros to the left side of it until it is the requested string
'		length.  This function returns a string.
'
'		@intInput:		The integer that needs zeros tacked onto it
'		@intNumDigits:	The number of digits the user wants to end up with
'==============================================================================================================
	
	Public Function FnAddLeadingZeros(ByVal intInput As Integer, ByVal intNumDigits As Integer) As String
		Dim str As String = intInput.ToString
		Do Until Len(str) >= intNumDigits
			str = "0" & str
		Loop
		Return str
	End Function	
	
	
'==============================================================================================================
' FnExists
'
' 		Simply tests for object existance.  This function is a little redundant.  It's meant to be here though
'		because the Interpreter makes use of it.  As a policy, I do not allow the Interpreter module to have
'		any interaction with Silk's Agent executable.
'
'		The Interpreter makes use of this function whenever a trailing hyphen is input by the user.  A trailing
'		hyphen indicates that the user wants Xscript to continue regardless if the step passes or fails.  A
'		common use for this is to handle a pop-up that only appears under certain circumstances.  If we do not
'		first check to see if the supplied object is on the screen in these instances, then Xscript will sit
'		and wait for a few seconds before throwing an exception - this slows down playback.  So for instructions
'		with a trailing hyphen, the Interpreter first calls FnExists to ensure it's not going to end up waiting 
'		for no reason.
'
'		Note: If the command is 'Wait', then FnExists is not called by the Interpreter because we may in fact
'		need to be waiting for the object to appear after all.
'
'		@strCtrlName:	The item to check for existance.
'==============================================================================================================
	
	Public Function FnExists(ByVal strCtrlName As String) As Boolean
		FnExists = _desktop.Exists(strCtrlName)
	End Function
	
	
'==============================================================================================================
' SubCaptureBitmap
'
' 		This function captures of image of the supplied test object and saves it to the specified location.  It
'		then returns a Bitmap object.

'		To do this, it is simply making use of Silk's native .CaptureBitmap method that is a member of every
'		single test object type.  However, .CaptureBitmap is retaining a write-lock on the files it creates.
'		This can become problematic for me because subsequent calls are frequently used to overwrite the previous
'		temp.bmp file.  When this happens, an IOException is thrown.
'
'		To mitigate the IOException, I am simply catching it and then saving the file in the same directory but
'		giving it a different file name.  The file name is generated by grabbing the current time all the way
'		down to a millionth of a second.  
'
'		@testObj:				The test object of which to get an image
'		@strFilePath:			The full file path in which to save the file
'		@blnOverwriteExisting:	If TRUE, existing bmp file is overwritten
'==============================================================================================================
	
	Public Function FnCaptureBitmap(ByRef testObj As TestObject, ByVal strFilePath As String, ByVal blnOverwriteExisting As Boolean) As Bitmap
	
		Dim dt As DateTime
	
		'If the destination file is present and the user wants to overwrite, try to first delete it.  Failing that, create a temp file name that is not already taken.
		If blnOverwriteExisting AndAlso File.Exists(strFilePath) Then
			Try
				File.Delete(strFilePath)
			Catch ioEx As IOException
				Do
					dt = DateTime.Now
					strFilePath = Path.GetDirectoryName(strFilePath) & "\" & dt.ToString("yyyyMMdd_HHmmss_fffffff")
				Loop Until Not File.Exists(strFilePath)
			End Try
		End If
		
		'Capture and return a Bitmap object based on the filepath
		FnCaptureBitmap = New Bitmap(testObj.CaptureBitmap(strFilePath))
		
	End Function
	
	
'==============================================================================================================
' FnImageContains
'
' 		Returns TRUE if the supplied Bitmap pattern object is present within the supplied source Bitmap object.
'
'		@bmpPattern:			A bitmap object of the pattern to find within the source bitmap object
'		@bmpSource:				The source bitmap object that needs to be searched for the pattern
'		@blnIgnoreRedPixels:	If TRUE, then red pixels in the pattern are ignored
'==============================================================================================================

	Public Function FnImageContains(ByRef bmpPattern As Bitmap, ByRef bmpSource As Bitmap, Optional ByVal blnIgnoreRedPixels As Boolean=False) As Boolean
	
		If FnGetContainedImagePoint(bmpPattern, bmpSource, blnIgnoreRedPixels) = Nothing Then
			FnImageContains = False
		Else
			FnImageContains = True
		End If
	
	End Function
	
'==============================================================================================================
' SubVerifyImage
'
' 		This routine is called when the user sends the VerifyImage command in Xscript.  It takes a screenshot
'		of the desktop and then searches for the specified bitmap within that image.  If it's found, then 
'		nothing happens and Xscript chugs merrily along.  If it's not found, then a bitmap failure is added
'		to the supplied Test object, and the desktop image along with the bitmap image are written to the
'		appropriate 'Bitmap check failures' folder.
'
'		@strPatternPath:			Full path to the pattern bitmap file
'		@dblTolerancePercentage:	Number 1-100 representing how far off each pixel color is allowed to be
'		@strIgnoreColor:			Pixels to ignore for matches.  Can be: "Red", "Green", or "Blue".
'		@test:						The currently exectuing Test object
'		@dtCurrentDateTime:			The datetime in which this Xscript run was kicked off
'		@dicConstants:				A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================
	
	Public Sub SubVerifyImage(ByVal strPatternPath As String, ByVal dblTolerancePercentage As Double, ByVal strIgnoreColor As String, ByRef test As Test, ByRef dtCurrentDateTime As DateTime, ByRef dicConstants As Dictionary(Of String, String))
		
		Dim bmpPattern As Bitmap
		Dim bmpSource As Bitmap
		Dim strBitmapCheckFailPath As String
		
		'Get the source and pattern bitmaps
		bmpPattern = New Bitmap(strPatternPath)
		bmpSource = FnCaptureBitmap(_desktop, dicConstants.Item("DirectoryToStoreTemporaryFiles") & "\Desktop.bmp", True)
		
		'Do the comparison.  If it failed, add it to the bitmap failure list and throw the bitmaps into the results folder.
		If FnGetContainedImagePoint(bmpSource, bmpPattern, strIgnoreColor, dblTolerancePercentage) = Nothing Then
			
			'Add the item to the bitmap failure list
			test.intBitmapFailureList.Add(test.intCurrentStepIndex)
			
			'Ensure the needed 'Bitmap check failure' folder exists
			strBitmapCheckFailPath = dicConstants.Item("ResultsLocation") & "\" & dicConstants.Item("ScreenshotsFolderName") & "\" & dtCurrentDateTime.ToString("yyyyMMdd_HHmmss") & "\" & _
									 test.strRelativeScreenShotPath & "\Bitmap check failures"
			If Not Directory.Exists(strBitmapCheckFailPath) Then Directory.CreateDirectory(strBitmapCheckFailPath)
			
			'Throw the bitmaps into the bitmap check failure directory
			bmpPattern.Save(strBitmapCheckFailPath & "\Bitmap check fail - step " & FnAddLeadingZeros(test.FnGetExcelPosition(test.intCurrentStepIndex), 4) & " - pattern.bmp")
			bmpSource.Save(strBitmapCheckFailPath & "\Bitmap check fail - step " & FnAddLeadingZeros(test.FnGetExcelPosition(test.intCurrentStepIndex), 4) & " - source.bmp")
			
		End If
		
	End Sub
	
	
'==============================================================================================================
' FnGetContainedImagePoint
'
' 		Returns the coordinate of a pattern match inside a larger bitmap.  If the pattern is not contained, then
'		Nothing is returned.
'
'		@src:					The source image as a bitmap object to scan through for the pattern
'		@bmp:					The pattern image as a bitmap object to search for within the source bitmap
'		@strIgnoreColor:		Color to always assume is a match.  Can be "red", "green" or "blue".
'		@dblTolerancePercent:	The percentage difference in color that each pixel is allowed to be a math
'==============================================================================================================
	
	Public Function FnGetContainedImagePoint(ByRef src As Bitmap, ByRef bmp As Bitmap, Optional ByVal strIgnoreColor As String=Nothing, Optional dblTolerancePercent As Double=0) As Point
        
		'Ensure the user actually passed data
        If src Is Nothing OrElse bmp Is Nothing Then Return Nothing
			
		'If the source is smaller than the pattern, obviously the pattern is not contained
		If src.Width < bmp.Width OrElse src.Height < bmp.Height Then Return Nothing
			
		'Convert the given color to ignore to a real RGB value
		Dim rIgnore, gIgnore, bIgnore As Integer
		If Not strIgnoreColor Is Nothing Then
			Select Case strIgnoreColor.ToLower
				Case "red"
					rIgnore = 255
				Case "green"
					gIgnore = 255
				Case "blue"
					bIgnore = 255
			End Select
		Else
			rIgnore = -1
			gIgnore = -1
			bIgnore = -1
		End If
			
		'Create rectangles that are as large as the two images
        Dim sr As New Rectangle(0, 0, src.Width, src.Height)
        Dim br As New Rectangle(0, 0, bmp.Width, bmp.Height)

		'This is an optimization which allows quicker access to bitmap pixel data
        Dim srcLock As BitmapData = src.LockBits(sr, Imaging.ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
        Dim bmpLock As BitmapData = bmp.LockBits(br, Imaging.ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)

		'Strides are the width of a single run of pixels rounded up to the nearest 4-byte boundary.  361=>1024, 160=>480.  I don't understand.
        Dim sStride As Integer = srcLock.Stride
        Dim bStride As Integer = bmpLock.Stride		
		
		'These are simply 'size' variables.  220052 and 17280.  I don't totally get this either.
        Dim srcSz As Integer = sStride * src.Height
        Dim bmpSz As Integer = bStride * bmp.Height
		
		'Optimization to store pixel data for the image.  Above we calculated the number of bytes needed to efficiently store each image in a variable.  That's what's happening now.
		'Array of bytes defined by how many bytes are needed.  Buffer is like a swappy storage location.
        Dim srcBuff(srcSz) As Byte
        Dim bmpBuff(bmpSz) As Byte

		'Optimization.  Not sure how this works.
        Marshal.Copy(srcLock.Scan0, srcBuff, 0, srcSz)
        Marshal.Copy(bmpLock.Scan0, bmpBuff, 0, bmpSz)

        'We don't need to lock the image anymore as we have a local copy
        bmp.UnlockBits(bmpLock)
        src.UnlockBits(srcLock)

		'Variable galore
        Dim x, y, x2, y2, sx, sy, bx, by, sw, sh, bw, bh As Integer
        Dim r, g, b As Byte
		Dim dblDiff As Double
		Dim intRsrc, intRbmp, intGsrc, intGbmp, intBsrc, intBbmp As Integer

		'Objects
        Dim p As Point = Nothing

		'Set variables
        bw = bmp.Width
        bh = bmp.Height
		
		'Limit scan to only what we need.  The extra corner point we need is taken care of in the loop itself.
        sw = src.Width - bw     
        sh = src.Height - bh   

		'Loop through every column of the source
        For y = 0 To sh
			
			'What is this?  Oh, sy is his 'working' coordinate probably in the source  
            sy = y * sStride
			
			'Loop through every row in this column
            For x = 0 To sw
				
				'I guess this gets the working x coorindates in the source
                sx = sy + x * 3
				
                'Get the RGB for this point within the source
                r = srcBuff(sx + 2)
                g = srcBuff(sx + 1)
                b = srcBuff(sx)
				
				'Get integer equivalents
				intRsrc = r
				intGsrc = g
				intBsrc = b
				intRbmp = bmpBuff(2)
				intGbmp = bmpBuff(1)
				intBbmp = bmpBuff(0)
				dblDiff = (Math.Abs(intRsrc-intRbmp) + Math.Abs(intGsrc-intGbmp) + Math.Abs(intBsrc-intBbmp)) / 768 * 100
								
				'If this source pixel matches the upper-left pixel of the bitmap (within tolerance), or the bitmap's upper left pixel is ignorable, we need to investigate further
				If (bmpBuff(2)=rIgnore AndAlso bmpBuff(1)=gIgnore AndAlso bmpBuff(0)=bIgnore) OrElse (dblDiff <= dblTolerancePercent) Then
					
					'Make a new drawing Point object for this pixel
                    p = New Point(x, y)
					
					'Loop through the sources's columns again
                    For y2 = 0 To bh - 1
						
						'Get the working y coordinate
                        by = y2 * bStride
						
						'Loop through the sources's rows again
                        For x2 = 0 To bw - 1
							
							'Get the working x coordinate
                            bx = by + x2 * 3
							
							'da fuq
                            sy = (y + y2) * sStride
                            sx = sy + (x + x2) * 3

							'Get the source's next pixel's RGB
                            r = srcBuff(sx + 2)
                            g = srcBuff(sx + 1)
                            b = srcBuff(sx)
							
							'Get integer equivalents
							intRsrc = r
							intGsrc = g
							intBsrc = b
							intRbmp = bmpBuff(bx + 2)
							intGbmp = bmpBuff(bx + 1)
							intBbmp = bmpBuff(bx)
							dblDiff = (Math.Abs(intRsrc-intRbmp) + Math.Abs(intGsrc-intGbmp) + Math.Abs(intBsrc-intBbmp)) / 768 * 100
							
							'If the corresponding pixel in the bmp is not ignoreable and it's not a match, disqualify the match
                            If Not (bmpBuff(bx + 2)=rIgnore AndAlso bmpBuff(bx + 1)=gIgnore AndAlso bmpBuff(bx)=bIgnore) AndAlso Not (dblDiff <= dblTolerancePercent) Then
								
								'Reset our hopes
                                p = Nothing
								
								'Pick up where we left off?
                                sy = y * sStride
								
								'Stop looping through the bmp's rows
                                Exit For
								
                            End If

                        Next
						
						'Stop looping through the bmp's columns
                        If p = Nothing Then Exit For
							
                    Next
				
				'End the region check
                End If

				'If a match was found, stop looping through each row in this column
                If p <> Nothing Then Exit For
					
			'Keep looping through each row in the source		
            Next
			
			'If a match was found, stop looping through the source's columns
            If p <> Nothing Then Exit For
				
		'Keep looping through each column in the source
        Next

		'Release
        bmpBuff = Nothing
        srcBuff = Nothing
	
		'Return
        Return p

    End Function
	
End Module