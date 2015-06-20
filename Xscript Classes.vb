'No references needed

Imports System.Text.RegularExpressions
	
'==============================================================================================================
' CLASS AssociativeArray2D
'
' 		This class stores tabular data in a single object while allowing direct reference to any particular 
'		cell/coordinate in the data without having to interate through the entire data structure.  The data is
'		simply stored in two nested dictionaries of string so that any object can be referenced by a pair of 
'		string typed keys.
'
'		Example #1
'
'			Imports CustomArrays
'
'			Public Module Main			
'				Public Sub Main()	
'					
'					Dim tdaa As AssociativeArray2D
'					tdaa = New AssociativeArray2D
'					
'					tdaa.Add("Charles Tronolone", "Nickname", "Chuck")
'					tdaa.Add("Charles Tronolone", "EyeColor", "Blue")
'					tdaa.Add("Charles Tronolone", "JobTitle", "Automation Engineer")
'					tdaa.Add("Jeffery Primmer", "Nickname", "Jeff")
'					tdaa.Add("Jeffery Primmer", "Location", "601 SW 2nd Ave")
'					tdaa.Add("Jeffery Primmer", "HireDate", "03/01/2011")
'					tdaa.Add("Jeffery Primmer", "JobTitle", "QA Lead")
'							
'					Console.Writeline(tdaa.Item("Charles Tronolone", "Nickname"))	'Chuck
'					Console.Writeline(tdaa.Item("Jeffery Primmer", "HireDate"))		'03/01/2011
'					Console.Writeline(tdaa.Exists("Jeffery Primmer", "EyeColor"))	'False
'					
'					For Each key1 As String In tdaa.Keys 
'						For Each key2 As String In tdaa.Keys(key1)
'							Console.Writeline(tdaa.Item(key1, key2))				'Outputs all data
'						Next
'					Next
'					
'				End Sub	
'			End Module
'==============================================================================================================
	
Public Class AssociativeArray2D
	
	'Declare the dictionary within a dictionary that will hold the entire data set
	'Note: I cannot use System.StringComparer.OrdinalIgnoreCase here because it won't compile when I apply it to
	'the inner dictionary.  I do not understand why.  This is a problem anytime we use the .Item method.  Oddly
	'enough though, the dictionary's .ContainsKey function is not case-sensitive.  See the link below for detail.
	'http://msdn.microsoft.com/en-us/library/system.collections.specialized.stringdictionary.containskey(v=vs.110).aspx
	'The workaround for .Item being case-sensitive is to pass all user-supplied keys into my FnConvertKeyToProperCase
	'function.  This function utilizes .ContainsKey to figure out if it exist; if it doesn't, then it loops through
	'all the keys and compares them in a case-insensitive manner in order to find if the given key does in fact exist
	'when you ignore the given case.  If it does exist, then it returns the properly-cased key and we proceed as normal.
	'Otherwise, an exception is simply thrown.
	Dim dicOuter = New Dictionary(Of String, Dictionary(Of String, Object))
	
	'Constructor
	Public Sub New()
	End Sub
	
	'Destructor
	Overrides Protected Sub Finalize()
	End Sub
		
'==============================================================================================================
' Add
'
' 		Sets a value within the 2D array.  If the specified key combination already exists, it overwrites the
'		stored value.  If the combination does not exist, it gets created.
'
' 		@strKey1:	The first key in the reference.  This can be conceptualized as a row number.
'		@strKey2:	The second key in the reference.  This can be conceptualized as a column number.
'		@objValue:	The value to set
'==============================================================================================================

	Public Sub Add(ByVal strKey1 As String, ByVal strKey2 As String, ByVal objValue As Object)
			
		'If the outer key does not exist, create a new outer and inner key with value
		If(dicOuter.ContainsKey(strKey1) = False) Then
			dicOuter.Add(strKey1, New Dictionary(Of String, Object) From {{strKey2, objValue}})	
			
		'If the outer key exists but the inner key is missing, add the inner key with value 
		Else If(dicOuter.Item(strKey1).ContainsKey(strKey2) = False) Then
			strKey1 = FnConvertKeyToProperCase(dicOuter, strKey1)
			dicOuter.Item(strKey1).Add(strKey2, objValue)
		
		'If the first and second order keys are present, then we need only update the value
		Else
			strKey1 = FnConvertKeyToProperCase(dicOuter, strKey1)
			strKey2 = FnConvertKeyToProperCase(dicOuter.Item(strKey1), strKey2)
			dicOuter.Item(strKey1).Item(strKey2) = objValue
		End If
	
	End Sub
		
'==============================================================================================================
' Item
'
' 		Retrieves the object at a specific location in the 2D associative array.
'
' 		@strKey1:	The first key in the reference.  This can be conceptualized as a row number.
'		@strKey2:	The second key in the reference.  This can be conceptualized as a column number.
'==============================================================================================================

	Public Function Item(ByVal strKey1 As String, Optional ByVal strKey2 As String=Nothing) As Object
	
		'Fix the casing
		strKey1 = FnConvertKeyToProperCase(dicOuter, strKey1)
		If Not strKey2 Is Nothing Then strKey2 = FnConvertKeyToProperCase(dicOuter.Item(strKey1), strKey2)
	
		'If the user has supplied both outer and inner keys, just return his object
		If Not strKey2 Is Nothing Then
			Item = dicOuter.Item(strKey1).Item(strKey2)
			
		'If the user wants the entire dictionary of String->Object pairs at a first-order key, send it
		Else 
			Item = dicOuter.Item(strKey1)
		End If
		
	End Function
	
	
'==============================================================================================================
' Update
'
' 		Update an object in the 2D associative array.  In the native .NET dictionary, we can simply write
'		myDictionary.Item("myKey") = "Elephant".  However, this is amateur hour over here and I don't know how
'		to set up my 2D associative array class in such a way that would let me use the equals sign so smoothly.
'		So instead, you get to use .Update.
'
'		You might be wondering why this function simply takes the supplied arguments and passes them directly to
'		the .Add function.  Right?  I mean, why not just call .Add directly.  Turns out, you can do exactly that
'		if you want.  The reason I created this .Update routine was because it implies that the key combination
'		already exists in the 2D associative array.  You can just as easily use .Add, but to the user, that really
'		seems like you're in danger of overwrite an existing key combination which could potentially throw an 
'		error.  So, this function is just to put your mind at ease and make it feel like it's doing what you want.
'
' 		@strKey1:	The first key in the reference
'		@strKey2:	The second key in the reference
'		@objValue:	The object to insert at the specified key pair
'==============================================================================================================

	Public Sub Update(ByVal strKey1 As String, ByVal strKey2 As String, ByVal objValue As Object)
		Add(strKey1, strKey2, objValue)
	End Sub
		
'==============================================================================================================
' ContainsKey
'
' 		Returns a boolean indicating whether the given key combo already exists.
'
' 		@strKey1:	The first key in the reference
'		@strKey2:	The second key in the reference
'==============================================================================================================

	Public Function ContainsKey(ByVal strKey1 As String, Optional ByVal strKey2 As String = "") As Boolean
	
		'Default return
		ContainsKey = False
	
		'If the key combo is valid on the outer
		If(dicOuter.ContainsKey(strKey1) = True) Then
			
			'If the user didn't supply strKey2, then we are done
			If(String.IsNullOrEmpty(strKey2)) Then
				ContainsKey = True
			
			'If the user DID supply strKey2, then find out if it exists
			Else If(dicOuter.Item(strKey1).ContainsKey(strKey2)) Then
				ContainsKey = True
			End If
			
		End If
	
	End Function

'==============================================================================================================
' Keys
'
' 		Returns a string List of keys in either the outer or inner dictionaries.  Conceptually, supplying a key 
'		to this function will return all the column names used by the specified row.  Omitting a key will simply
'		earn you all a list of all the rows being used.
'
' 		@strKey1:	The key of the "row" you want to retrieve every column name from.
'==============================================================================================================

	Public Function Keys(Optional ByVal strKey1 As String="") As List(Of String)
					
		'Set object
		Keys = New List(Of String)
		
		'If the user wants the outer keys
		If(String.IsNullOrEmpty(strKey1)) Then
			
			'I'm iterating and building my own list instead of returning dicOuter.Keys or a derrived list from
			'dicOuter.Keys because it's throwing unresolvable errors that I believe are due to the nested dictionary setup.
			For Each strOuterKey As String In dicOuter.Keys
				Keys.Add(strOuterKey)
			Next
			
		'If the user wants inner keys
		Else
			
			'This loop is for the same reason as above... something odd is happening that I cannot resolve.
			For Each strInnerKey As String In dicOuter.Item(strKey1).Keys
				Keys.Add(strInnerKey)
			Next
			
		End If
		
	End Function
	

'==============================================================================================================
' FnConvertKeyToProperCase
'
' 		These functions take a dictionary object and a key, then returns the key in the proper case assuming it
'		exists.  If the doesn't exist, an exception is thrown.
'
'		There is a good reason that this is split amongst three functions.  In a perfect world, it would just
'		be one function that accepted a list of keys and a particular key.  However, retrieving a list of keys
'		from a dictionary is harder than it looks.  The only way I've successfully done this is by writing
'		"Dim strKeyList As New List(Of String)(myDic.Keys)".  The dictionary's .Keys method doesn't actually
'		return a normal type of List; it returns some funky alternative to a list that I do not know how to 
'		iterate through as strings.  So, it ended up being easier to just pass the whole dictionary object
'		directly to FnConvertKeyToProperCase, then allow the function to derive the list of keys on its own by
'		using the aforementioned Dim statement.  Once the List is created, then it can be passed into the third
'		and proper form of FnConvertKeyToProperCase.  The reason there is 3 functions and not 2 is that sometimes
'		I need get the proper case for the outer key, and sometimes I need to get it for the inner key.  This 
'		means I need a function that accepts a 1D associative array, and another that accepts a 2D associative array.
'
'		@dic:		The 1D or 2D dictionary that we want to look through for our properly-cased key
'		@strKey:	The key that we want converted to its proper case
'==============================================================================================================
	
	Public Function FnConvertKeyToProperCase(ByRef dic As Dictionary(Of String, Dictionary(Of String, Object)), ByVal strKey As String) As String
		Dim strPrelimKeyList As New List(Of String)(dic.Keys)
		FnConvertKeyToProperCase = FnConvertKeyToProperCase(strPrelimKeyList, strKey)
	End Function
	
	Public Function FnConvertKeyToProperCase(ByRef dic As Dictionary(Of String, Object), ByVal strKey As String) As String
		Dim strPrelimKeyList As New List(Of String)(dic.Keys)
		FnConvertKeyToProperCase = FnConvertKeyToProperCase(strPrelimKeyList, strKey)
	End Function

	Public Function FnConvertKeyToProperCase(ByRef strKeyList As List(Of String), ByVal strKey As String) As String
		
		'Default return
		FnConvertKeyToProperCase = Nothing
		
		'If the case of the key is already valid, just return it
		If strKeyList.Contains(strKey) Then
			FnConvertKeyToProperCase = strKey
			
		'If the case of the key is invalid, iterate through all the keys and do a case-insensitive comparison until we find our key
		Else
			For Each strOneKey As String In strKeyList
				If String.Compare(strOneKey, strKey, True)=0 Then
					FnConvertKeyToProperCase = strOneKey
					Exit For
				End If
			Next strOneKey
			
			'If our key was never found, throw an exception
			If FnConvertKeyToProperCase Is Nothing Then
				Throw New Exception("Regardless of case, the given key '" & strKey & "' is not present in the dictionary.")
			End If
		End If
	End Function
	
End Class	
	
'==============================================================================================================
' Test Class
'
' 		The Test class holds all properties that pertain to a user's test.  Typically, a single test corresponds
'		to a single Excel file in Xscript.  However, for tests that utilize loops, each iteration of the loop
'		create a separate test object, too.
'
'		The Test object contains all sorts of useful properties.  Notably, it contains the corresponding object 
'		map, user variables, and results.  It also contains a reference to its parent if it has one, and a dictionary
'		of all its child tests.
'==============================================================================================================

Public Class Test
	
	'Private declarations
	Private _strUserVarDic As New Dictionary(Of String, Object)(System.StringComparer.OrdinalIgnoreCase)
	Private _strLocatorDic As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
	Private _strParentXpath As String
	Private _strInstruction As String
	Private _testRoot As Test
	Private _strRelativeScreenShotPath As String
	
	'Public objects
	Public strStepsList As New List(Of String)
	Public objFailureDic As New Dictionary(Of Integer, Object)
	Public strWarningList As New List(Of String)
	Public intBitmapFailureList As New List(Of Integer)
	Public strOutputTextList As New List(Of String)
	Public testChildDic As New Dictionary(Of String, Test)(System.StringComparer.OrdinalIgnoreCase)
	Public testParent As Test
	Public dataWorksheetDic As New Dictionary(Of String, AssociativeArray2D)
	
	'Simple properties	
	Public strTestName As String
	Public intCurrentStepIndex As Integer	
	Public strFinalMsg As String
	Public blnTestDone As Boolean
	Public strExcelFilePath As String
	Public strActiveChildTestName As String
	Public blnReportResults As Boolean
	Public blnAbortOnFail As Boolean
	Public intParentStartLoopIndex As Integer	
	Public intNumStepsAdded As Integer
	Public strLoopColumnHeader As String
	Public blnSimplyPassed As Boolean
	
	'Constructor
	Public Sub New(ByVal strNewTestName As String, ByVal strNewExcelFilePath As String)
		strTestName = strNewTestName
		intCurrentStepIndex=0
		strParentXpath=""
		strExcelFilePath=strNewExcelFilePath
		strActiveChildTestName=Nothing	
		blnReportResults = True
		blnAbortOnFail = True
	End Sub
	
	
'==============================================================================================================
' PROPERTY: testRoot
'
' 		Recursive property which will always end up returning the initial root Test object.
'==============================================================================================================

	Public Property testRoot As Test
		Get
			If Not testParent Is Nothing
				_testRoot = testParent.testRoot
			Else
				_testRoot = Me
			End If
			Return _testRoot
		End Get
		Set(ByVal testNewRoot As Test)
			_testRoot = testNewRoot
		End Set
	End Property
	
	
'==============================================================================================================
' PROPERTY: strInstruction
'
' 		This ensures that strInstruction always reflects the value corresponding to intCurrentStepIndex.
'==============================================================================================================

	Public Property strInstruction As String
		Get
			_strInstruction = strStepsList.Item(intCurrentStepIndex)
			Return _strInstruction
		End Get
		Set(ByVal strNewInstruction As String)
			_strInstruction = strNewInstruction
		End Set
	End Property
	
'==============================================================================================================
' PROPERTY: strRelativeScreenShotPath
'
' 		Returns the screenshot folder path of this test as it relates to the "time folder" for the Xscript run.
'==============================================================================================================

	Public Property strRelativeScreenShotPath As String
		Get
			_strRelativeScreenShotPath = Nothing
			If Not testParent Is Nothing Then
				_strRelativeScreenShotPath = testParent.strRelativeScreenShotPath & "\"
			End If
			_strRelativeScreenShotPath = _strRelativeScreenShotPath & Regex.Replace(If(strLoopColumnHeader Is Nothing, strTestName, strLoopColumnHeader), "[^a-zA-Z0-9\s_-]", "")
			Return _strRelativeScreenShotPath
		End Get
		Set(ByVal strNewRelativeScreenShotPath As String)
			_strRelativeScreenShotPath = strNewRelativeScreenShotPath
		End Set
	End Property
	
	
'==============================================================================================================
' PROPERTY: strUserVarDic
'
' 		Aggregates and returns a dictionary of user vars by combining and overwriting all "older" user vars.
'==============================================================================================================
	
	Public Property strUserVarDic As Dictionary(Of String, Object)	
		Get
			If Not Me.testParent Is Nothing
				For Each kvp As KeyValuePair(Of String, Object) In Me.testParent.strUserVarDic
					If Not Me._strUserVarDic.ContainsKey(kvp.Key) Then
						Me._strUserVarDic.Add(kvp.Key, kvp.Value)
					End If
				Next kvp
			End If
			Return Me._strUserVarDic
		End Get		
		Set(ByVal newDic As Dictionary(Of String, Object))
			Me._strUserVarDic = newDic
		End Set		
	End Property
	
	
'==============================================================================================================
' PROPERTY: strLocatorDic
'
' 		Aggregates and returns a dictionary of object xpaths by combining and overwriting all "older" xpaths.
'==============================================================================================================
	
	Public Property strLocatorDic As Dictionary(Of String, String)	
		Get
			If Not Me.testParent Is Nothing
				For Each kvp As KeyValuePair(Of String, String) In Me.testParent.strLocatorDic
					If Not Me._strLocatorDic.ContainsKey(kvp.Key) Then
						Me._strLocatorDic.Add(kvp.Key, kvp.Value)
					End If
				Next kvp
			End If
			Return Me._strLocatorDic
		End Get		
		Set(ByVal newDic As Dictionary(Of String, String))
			Me._strLocatorDic = newDic
		End Set		
	End Property
	
	
'==============================================================================================================
' PROPERTY: strParentXpath
'
' 		Returns a concatenated string containing all the "older" tests' xpaths.  Furthermore, setting this 
'		string will set the strParentXpath of all "older" tests, too.
'==============================================================================================================
		
	Public Property strParentXpath As String
		Get
			If Not Me.testParent Is Nothing
				Me._strParentXpath = Me.testParent.strParentXpath & Me._strParentXpath
			End If
			Return Me._strParentXpath
		End Get		
		Set(ByVal strNewParentXpath As String)
			Me._strParentXpath = strNewParentXpath
			If Not Me.testParent Is Nothing
				Me.testParent.strParentXpath = strNewParentXpath
			End If
		End Set		
	End Property
	
	
'==============================================================================================================
' METHOD: FnGetExcelPosition
'
' 		Returns the row number that corresponds to the given step index.  This method takes into account all
'		ancestor positions and their uses of InsertSteps.
'==============================================================================================================
	
	Public Function FnGetExcelPosition(ByVal intThisTestIndex As Integer) As Integer
		If Me.testParent Is Nothing Then
			FnGetExcelPosition = 2 - Me.intNumStepsAdded + intThisTestIndex
		Else
			FnGetExcelPosition = 1 + intThisTestIndex - Me.intNumStepsAdded + testParent.FnGetExcelPosition(Me.intParentStartLoopIndex)
		End If
	End Function
	
	
'==============================================================================================================
' METHOD: SubBuildResult
'
' 		Sets the value for this test's strFinalMsg string.  This is a very tricky method that takes into account
'		the status of all parents and whether a child test is meant to report its own results.
'
'		@blnIncludeStackTrace:	If TRUE, any available stacktrace will be thrown into the strFinalMsg.
'==============================================================================================================
	
	Public Sub SubBuildResult(ByVal blnIncludeStackTrace As Boolean)
		
		Dim blnWeKnowAtLeastOneChildFailed As Boolean
		strFinalMsg = Nothing
		
		'If a failure occurred
		If objFailureDic.Count > 0 Then
			For Each kvp As KeyValuePair(Of Integer, Object) In objFailureDic
				If String.IsNullOrEmpty(strFinalMsg)=False Then strFinalMsg = strFinalMsg & Chr(10)
				strFinalMsg = strFinalMsg & "Failure at step #" & FnGetExcelPosition(kvp.Key) & ": " & strStepsList.Item(kvp.Key) & ".  " & _
							  If(TypeOf kvp.Value Is System.Exception, "Exception thrown: " & kvp.Value.GetType.ToString & ".  Message: " & kvp.Value.Message, "Message: " & kvp.Value) & _
							  If(TypeOf kvp.Value Is System.Exception And blnIncludeStackTrace, Chr(10) & kvp.Value.StackTrace, "")
			Next kvp
		End If
		
		'If a bitmap check failed
		For Each intOneFailStepNum As Integer In intBitmapFailureList
			If String.IsNullOrEmpty(strFinalMsg)=False Then strFinalMsg = strFinalMsg & Chr(10)
			strFinalMsg = strFinalMsg & "Bitmap check failed at step #" & FnGetExcelPosition(intOneFailStepNum) & ": " & strStepsList.Item(intOneFailStepNum)
		Next intOneFailStepNum
		
		'If no errors occurred
		If String.IsNullOrEmpty(strFinalMsg) Then 
			If strOutputTextList.Count = 0 Then
				strFinalMsg = "PASS"
				blnSimplyPassed = True
			Else
				strFinalMsg = "PASS."
			End If
		End If
		
		'If there was output text
		If strOutputTextList.Count > 0 Then
			strFinalMsg = strFinalMsg & "  OUTPUT TEXT: "
			For Each strOneOutput As String In strOutputTextList
				strFinalMsg = strFinalMsg & "[" & strOneOutput & "] "
			Next strOneOutput
		End If
		
		'Combine this test's results with its children
		For Each kvpChild As KeyValuePair(Of String, Test) In Me.testChildDic
			
			'If the child didn't report himself
			If kvpChild.Value.blnReportResults = False Then
				
				'Build the child's final message
				kvpChild.Value.SubBuildResult(blnIncludeStackTrace)
				
				'If this is the first child test we've examined that has not simply passed, either wipe out the parent's value (because it actually simply passed itself) or give the parent a name tag
				If Not String.Equals(kvpChild.Value.strFinalMsg, "PASS") AndAlso Not blnWeKnowAtLeastOneChildFailed Then
					
					'Set a flag to indicate that we already know at least one child failed
					blnWeKnowAtLeastOneChildFailed = True
					
					'If the parent simply passed, wipe out the parent
					If Me.blnSimplyPassed Then
						Me.strFinalMsg = Nothing
					
					'Else, the parent did not simply pass, so add a name tag
					Else
						Me.strFinalMsg = "[" & Me.strTestName.ToUpper & "]" & Chr(10) & Me.strFinalMsg
					End If
					
				End If
					
				'If the child isn't blank and it has text other than 'PASS'
				If Not String.IsNullOrEmpty(kvpChild.Value.strFinalMsg) AndAlso Not String.Equals(kvpChild.Value.strFinalMsg, "PASS") Then
					
					'If this test's final message is already set to something, insert a couple line breaks
					If Not String.IsNullOrEmpty(Me.strFinalMsg) Then
						Me.strFinalMsg = Me.strFinalMsg & Chr(10) & Chr(10)
					End If
					
					'If this child didn't simply pass, put its name tag within the parent
					If Not kvpChild.Value.blnSimplyPassed Then 
						Me.strFinalMsg = Me.strFinalMsg & "[" & kvpChild.Value.strTestName.ToUpper & "]" & Chr(10)
					End If
					
					'Write the child's message into the parent
					Me.strFinalMsg = Me.strFinalMsg & kvpChild.Value.strFinalMsg	
					
				End If
				
			End If
			
		Next kvpChild
		
	End Sub	
	
End Class