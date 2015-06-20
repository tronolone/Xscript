'Add reference: [any custom .NET script you want to utilize within Xscript]

Imports System.Reflection

Public Module CallerModule		
	
'==============================================================================================================
' FnCallCustomMethod
'
' 		Allows the user to call non-Xscript functions via the Xscript framework.
'
'		The .NET script containing the user's function must be referenced via the 'Properties' tab.  The function
'		must be contained in a module.  The module must have a unique name such as "CommonFunctionsModule".
'		The contained function must be declared as Public.
'
'		FnCallCustomMethod will return TRUE if the user's function was successfully located.
'
' 		@strMethodName:	The name of the user's function he wishes to invoke
' 		@objArgList:	The optional list of arguments to send to the user's function
'==============================================================================================================

	Public Function FnCallCustomMethod(ByVal strMethodName As String, Optional ByRef objArgList As List(Of Object)=Nothing) As Boolean
		
		'Declarations
		Dim moduleTypeList As New List (Of Type)
		Dim methodList As New List(Of MethodInfo)
		
		'Default return
		FnCallCustomMethod = False
				
		'**** ADD THE NAME OF YOUR MODULE HERE.  IT MUST BE A UNIQUE NAME. ****
		'moduleTypeList.Add(GetType(CommonFunctionsModule))
		
		'Cycle through the modules
		For Each oneModuleType As Type In moduleTypeList
			
			'Cycle through the methods
			For Each oneMethodInfo As MethodInfo In oneModuleType.GetMethods()
				
				'If this is the corresponding method, invoke it
				If String.Compare(oneMethodInfo.Name, strMethodName, True)=0 Then
					Try
						oneMethodInfo.Invoke(Nothing, If(objArgList.Count=0, Nothing, objArgList.ToArray))
					Catch eTI As TargetInvocationException
						Throw eTI.GetBaseException
					End Try
					FnCallCustomMethod = True
					Exit For
				End If
				
			Next oneMethodInfo
			
			If FnCallCustomMethod Then Exit For
			
		Next oneModuleType
		
	End Function
	
End Module