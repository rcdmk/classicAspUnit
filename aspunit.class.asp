<%
' aspUnit testing framework class
' By RCDMK - rcdmk@rcdmk.com

' The MIT License (MIT)
' Copyright (c) 2012 RCDMK - rcdmk@rcdmk.com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software
' and associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute,
' sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
' BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


' Constants
const AU_ASSERT_EXISTS = 1
const AU_ASSERT_NULL = 2
const AU_ASSERT_EMPTY = 3
const AU_ASSERT_EQUALS = 4
const AU_ASSERT_NOT_EQUALS = 5
const AU_ASSERT_IS_ARRAY = 6
const AU_ASSERT_IS_OBJECT = 7
const AU_ASSERT_IS_A = 8


' Classes
class aspUnit
	' fields
	dim dTestCases
	
	
	' constructor and destructor
	private sub class_initialize()
		set dTestCases = createObject("Scripting.Dictionary")
	end sub
	
	private sub class_terminate()
		dim testCase
		for each testCase in dTestCases.keys
			set dTestCases(testCase) = nothing
			dTestCases.remove testCase
		next
		
		set dTestCases = nothing
	end sub
	
	
	' public methods
	public function AddTestCase(byval name)
		dim testCase
		set testCase = new aspUnitTestCase
		testCase.Name = name
		
		dTestCases.Add name, testCase
		
		set AddTestCase = testCase
	end function
	
	public function Run()
		dim results, testCase
		
		set results = new aspUnitTestResult
		
		for each testCase in dTestCases.Items
			testCase.Run
			results.TestCases.Add testCase
		next
		
		set Run = results
	end function
end class


class aspUnitTestCase
	' fields
	dim sName, sStatus
	dim dTests
	dim setupCode, tearDowncode
	
	
	' properties
	public property get Name()
		Name = sName
	end property
	
	public property let Name(value)
		sName = value
	end property

	public property get Status()
		Status = sStatus
	end property
	
	public property get Tests()
		set Tests = dTests
	end property
	

	' constructor and destructor
	private sub class_initialize()
		set dTests = createObject("Scripting.Dictionary")
		sStatus = "Inconclusive"
	end sub
	
	private sub class_terminate()
		dim test
		for each test in dTests.keys
			set dTests(test) = nothing
			dTests.remove test
		next
		
		set dTests = nothing
	end sub

	
	
	' public methods
	public sub Setup(byval setupCallbackCode)
		setupCode = setupCallbackCode
	end sub

	public sub Teardown(byval terardownCallbackCode)
		tearDowncode = terardownCallbackCode
	end sub
	
	public function AddTest(byval testName)
		dim test
		set test = new aspUnitTestMethod
		test.Name = testName
		
		dTests.add testName, test
		
		set AddTest = test
	end function
	
	public sub Run()
		dim passed
		passed = false
		
		if dTests.Count > 0 then
			sStatus = "Passed"
			
			for each test in dTests.Items
				execute setupCode
				
				if not test.Run() then
					sStatus = "Failed"
				end if
				
				execute tearDowncode
			next
		end if
	end sub
end class



class aspUnitTestMethod
	' fields
	dim sName, sStatus
	dim cAssertions, cErrors
	
	' properties
	public property get Name()
		Name = sName
	end property
	
	public property let Name(value)
		sName = value
	end property

	public property get Status()
		Status = sStatus
	end property
	
	public property get Output()
		dim sOutput
		
		if cAssertions.Count > 0 then
			if cErrors.Count > 0 then
				sOutput = "<li>" & join(cErrors, "</li><li>") & "</li>"
			else
				sOutput = "OK"
			end if
		else
			sOutput = "Untested"			
		end if
		
		Output = sOutput
	end property

	
	' constructor and destructor
	private sub class_initialize()
		sStatus = "Inconclusive"
		
		set cAssertions = new aspUnitCollection
		set cErrors = new aspUnitCollection		
	end sub
	
	private sub class_terminate()
		cAssertions.Clear
		cErrors.Clear
		
		set cAssertions = nothing
		set cErrors = nothing
	end sub
	
	
	' public methods
	public sub AssertExists(byref obj, byval message)
		addAssertion AU_ASSERT_EXISTS, obj, null, message
	end sub
	
	public sub AssertIsA(byref obj, byval typeName, byval message)
		addAssertion AU_ASSERT_IS_A, obj, typeName, message
	end sub
	
	public sub AssertEquals(byref obj, byref obj2, byval message)
		addAssertion AU_ASSERT_EQUALS, obj, obj2, message
	end sub
	
	
	public function Run()
		dim assertion, passed, msg
		passed = false
		
		if cAssertions.Count > 0 then
			passed = true
			sStatus = "Passed"
			
			for each assertion in cAssertions.Collection
				if not assertion.Run() then
					passed = false
					sStatus = "Failed"
					cErrors.Add assertion.Message
				end if
			next
		end if
		
		Run = passed
	end function
	
	
	' private methods
	private sub addAssertion(byval mode, byref obj1, byref obj2, byval msg)
		dim assertion		
		set assertion = new aspUnitAssertion
		
		assertion.Mode = mode
		assertion.Message = msg
		
		if isObject(obj1) then
			set assertion.Obj1 = obj1
		else
			assertion.Obj1 = obj1
		end if
		
		if isObject(obj2) then
			set assertion.Obj2 = obj2
		else
			assertion.Obj2 = obj2
		end if		
		
		cAssertions.Add assertion
	end sub
end class


class aspUnitTestResult
	' fields
	dim iTests, iPassed, iFailed, iErrors
	dim cTestCases
	
	' properties
	public property get Tests()
		Tests = cTestCases.Count
	end property
	
	public property get Passed()
		Passed = iPassed
	end property
	
	public property get Failed()
		Failed = iFailed
	end property
	
	public property get Errors()
		Errors = iErrors
	end property
	
	public property get TestCases()
		set TestCases = cTestCases
	end property
	
	
	' constructor and desctructor
	private sub class_initialize()
		iTests = 0
		iPassed = 0
		iFailed = 0
		iErrors = 0
		set cTestCases = new aspUnitCollection
	end sub
	
	private sub class_terminate()
		cTestCases.clear()
		set cTestCases = nothing
	end sub
end class


class aspUnitCollection
	' fields
	dim aCollection()
	
	
	' properties
	public default property get Collection()
		Collection = aCollection
	end property

	public property get Count()
		Count = ubound(aCollection) + 1
	end property
	
	
	' constructor
	private sub class_initialize()
		redim aCollection(-1)
	end sub
	
	private sub class_terminate()
		Clear
	end sub
	
	
	' public methods
	public sub Add(byref value)
		redim preserve aCollection(ubound(aCollection) + 1)
		if isobject(value) then
			set aCollection(ubound(aCollection)) = value
		else
			aCollection(ubound(aCollection)) = value
		end if
	end sub
	
	public function Remove(obj)
		dim i, index, total, result
		i = 0
		total = ubound(aCollection)
		result = false
		
		index = getIndex(obj)
		
		' If the object was found
		if index >= 0 then
			'Destroy the object
			set aCollection(index) = nothing
			
			' Shifts the objecs above this index one index less
			for i = index to total
				set aCollection(i) = aCollection(i + 1)
			next
			
			' Destroy the las item to be removed
			set aCollection(total) = nothing
			
			' Shorten the array, removing the last item
			redim preserve aCollection(total - 1)
			
			result = true
		End If
		
		Remove = result
	end function	
	
	public sub Clear()
		dim obj
		for each obj in aCollection
			Remove obj
		next
	end sub
	
	' private methods
	private function getIndex(obj)
		dim i, index
		index = -1
		
		do while i < total
			if aCollection(i) = obj then
				index = i
				exit do
			end if
			
			i = i + 1
		loop
		
		getIndex = index
	end function
end class


class aspUnitAssertion
	' fields
	dim iMode, sMessage, oObj1, oObj2
	
	
	' properties
	public property get Mode()
		Mode = iMode
	end property
	
	public property let Mode(value)
		iMode = value
	end property
	
	public property get Message()
		Message = sMessage
	end property
	
	public property let Message(value)
		sMessage = value
	end property	
	
	public property get Obj1()
		if isObject(oObj1) then
			set Obj1 = oObj1
		else
			Obj1 = oObj1
		end if
	end property
	
	public property let Obj1(value)
		oObj1 = value
	end property	
	
	public property set Obj1(value)
		set oObj1 = value
	end property	
	
	public property get Obj2()
		if isObject(oObj2) then
			set Obj2 = oObj2
		else
			Obj2 = oObj2
		end if
	end property
	
	public property let Obj2(value)
		oObj2 = value
	end property	
	
	public property set Obj2(value)
		set oObj2 = value
	end property
	
	
	' public methods
	public function Run()
		dim passed, msg, val1, val2
		
		val1 = objectValue(oObj1)
		val2 = objectValue(oObj2)
		
		passed = false
		
		select case iMode
			case AU_ASSERT_EXISTS:
				if isObject(oObj1) then
					if typeName(oObj1) <> "Nothing" then
						passed = true
					end if
					
				elseif not isnull(oObj1) then
					if oObj1 <> "" then passed = true
				end if
				
				if not passed then
					if sMessage = "" or isnull(sMessage) then
						msg = "Object doesn't exists (" & val1 & ")"
					else
						msg = replace(sMessage, "{1}", val1)
					end if
				end if
			
			case AU_ASSERT_IS_A:
				if typeName(oObj1) = oObj2 then
					passed = true					
				
				else
					if sMessage = "" or isnull(sMessage) then
						msg = "Object " & val1 & " is not of type " & val2
					else
						msg = replace(replace(sMessage, "{1}", val1), "{2}", val2)
					end if
				end if
			
			case AU_ASSERT_EQUALS:
				if oObj1 = oObj2 then
					passed = true
				else
					if sMessage = "" or isnull(sMessage) then
						msg = val1 & " is not equals " & val2
					else
						msg = replace(replace(sMessage, "{1}", val1), "{2}", val2)
					end if
				end if
			
			case default
				msg = "Invalid assertion mode"
		end select
		
		if not passed then sMessage = msg
		Run = passed
	end function
	
	
	private function objectValue(byref obj)
		dim name
		name = typeName(obj)
		
		if isObject(obj) or name = "Empty" then
			objectValue = name
			
		elseif name = "Array" then
			objectValue = "[" & join(obj, ", ") & "]"
		
		else
			objectValue = obj
		end if
	end function
end class
%>