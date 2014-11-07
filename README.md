Test2
=====
Dim objFlowExcel
Dim objFlowWorkBook
Dim objWorkSheet
Dim strCurModule
Dim intCurrRow
Dim TestLocation
Dim oFso
Dim arrModules
Dim strColName
Dim TestCaseFlag
Dim PerformTest


'  Following lines are coded to get the relative path of the test placed 
Set oFso = CreateObject( "Scripting.FileSystemObject" )
Environment( "TestLocation" ) = oFso.GetParentFolderName( Environment( "TestDir" ) )

'Loading Functional Libraries.
LoadFunctionLibrary ( Environment( "TestLocation" ) &"\Functions\Framework_Specific.vbs" )
LoadFunctionLibrary ( Environment( "TestLocation" ) &"\Functions\GenericVBscript.vbs" )
LoadFunctionLibrary ( Environment( "TestLocation" ) &"\Functions\ApplicationSpecific.vbs" )

'  Fn_Interface function is located Framework_Specific function file which is used to get the input from user on running the test ( HTML Popup will be displayed )
Call Fn_Interface (  )

If Environment( "Modules" ) <> "" Then
'	  Fn_Config  is located Framework_Specific function file which is used to import Object Repository and external environment variables defined in XML file
	Call Fn_Configuration ( "SetUp" )
	
'	 Open only  if the Datatable exist in the location specified
	If oFso.FileExists( Environment( "DatatableLocation" ) ) = True Then
	
'	 Create an object to open the  datatable
		Set objFlowExcel = CreateObject( "Excel.Application" )
		Environment("objFlowExcel") = objFlowExcel
		Set objFlowWorkBook = Environment( "objFlowExcel" ).Workbooks.Open ( Environment( "DatatableLocation" ) )
		Environment( "objFlowWorkBook" ) = objFlowWorkBook
		Environment("objFlowExcel").Visible = True
'	  Following line execute the macro defined in the datatable . This macro will set the number of test cases that needs to run on the module that is been executed.
		PerformTest = Fn_TestCaseCollection( objFlowWorkBook )
	
'	 IIterate the number of modules defined 
		arrModules = Split(Environment( "Modules" ), "@@", -1, 1)
	
'	  Fn_CreateResultFile is located  Generic VBscript  function file which is used to create the output file in HTML/Excel format 
		Call Fn_CreateResultFile( )
	
'	 ************************************************ Iteration starts at module level  ( List of modules are available in interface ) ***********************************************************************************************
		For i = 0 to Ubound(arrModules)
			If Trim( arrModules ( i ) ) <> ""  Then	
	
				Environment( "strCurModule" ) = Trim( arrModules ( i ) )
				Environment( "intCurModule" ) = i
				Environment("Driver") = True
				Call Fn_Configuration ( "ModuleStart" )
	
				Set objWorkSheet = objFlowWorkBook.Worksheets( Trim ( arrModules ( i ) ) ) 
				Environment("CurrModuleSheet") = objWorkSheet
	
'	  Following lines will fetch the Used rows and columns count in the module sheet  and assigned them to an environemnt variable
				intCurModRowCnt = objWorkSheet.UsedRange.Rows.Count
				intCurModColCnt = objWorkSheet.UsedRange.Columns.Count
				Environment("intCurModRowCnt") = intCurModRowCnt
				Environment("intCurModColCnt") = intCurModColCnt 
	
'	 ************************************************   Iteration starts at row level ( Each row is a test case ) ***********************************************************************************************************************
				For rows = 2 to intCurModRowCnt
				
					Environment("TestCaseFlag") = True
					Environment("CurrentRow") = rows
					
					For intI = 1 To PerformTest(i).Count
						TestCaseFlag = False
	
'	 ************************************************   Environment  varaible started storing  in Environment header ( Row 1 of Data sheet is header and  Test case row is varaiable ) ************************************
						For cols = 1 to intCurModColCnt
	
'	 If the test case row is not required to run stop assigning environment variables for that test case
							If  Trim( objWorkSheet.Rows( rows ).Columns( cols ).Value ) <> Trim( PerformTest(i).Item( "Case"& intI ) ) And cols = 1 Then
								TestCaseFlag = False
								Exit For
							End If
							strColName = Trim( objWorkSheet.Rows( 1 ).Columns( cols ).Value )
							If strColName <> "" Then
								Environment( strColName ) = Trim( objWorkSheet.Rows( rows ).Columns( cols ).Value )
							End If
							TestCaseFlag = True
					
						Next
						
'	 ************************************************   Environment  varaible completed storing  in  Environment header ( Row 1 of Data sheet is header and  Test case row is varaiable ) *****************************
	
'	 Start executing the script for the cases specified as Run
						If TestCaseFlag = True Then
						
							Call Fn_TableName ( Fn_ResultGenerator( "Header1" ), "Header1", "Report" )
							Call Fn_TableName ( Fn_ResultGenerator( "Header2" ), "Header2", "Report" )
							
'	 Fn_Module  is located Framework_Specific function file which actually taken the function name from the driver sheet and start executing them
							Call Fn_Module ( Trim( arrModules ( i ) ) )
							Environment( "Driver" ) = False
						End If
					
					Next
					
				Next
'	 Fn_Configuration   is located Framework_Specific function file which will close all open connections
				Call Fn_Configuration ( "ModuleEnd" )
	
'	 ************************************************   Iteration ends at row level ( Each row is a test case ) **********************************************************************************************************************
			End If
			Set objWorkSheet = Nothing
		Next
'	 ************************************************ Iteration ends at module level  ( List of modules are available in interface ) ************************************************************************************************	
		objFlowWorkBook.Close
		Call Fn_Configuration ( "End" )
		Set PerformTest = Nothing
	
'	 Throw an error if the path speciified is not correct
	Else
		Msgbox "Unable to find the Datatable in the path"& Chr( 10 ) & Environment( "DatatableLocation" ),,""& Chr( 10 ) &"Test Run Stopped"
	End If
End If

'Close all Objects created
Set objFlowWorkBook = Nothing
Set objFlowExcel = Nothing
Set oFso = Nothing
'----------------------------------------------------------------End Of Driver Script---------------------------------------------------------------------------------------
