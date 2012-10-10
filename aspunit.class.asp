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
		
		set AddTestCase = testCase
	end function
	
	public function Run()
		dim result, testCase
		
		set result = new aspUnitTestResult
		
		for each testCase in dTestCases
			results.TestCases.Add testCase
		next
		
		set Run = result
	end function
end class


class aspUnitTestCase
	' fields
	dim sName, dTests
	dim setupCode, tearDowncode
	
	
	' properties
	public property get Name()
		Name = sName
	end property
	
	public property let Name(value)
		sName = value
	end property


	' constructor and destructor
	private sub class_initialize()
		set dTests = createObject("Scripting.Dictionary")
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
end class

class aspUnitTestMethod
	' fields
	dim sName
	
	' properties
	public property get Name()
		Name = sName
	end property
	
	public property let Name(value)
		sName = value
	end property

	
	' public methods
	public sub AssertExists(byref obj)
		
	end sub
	
	public sub AssertIsA(byref obj, byval typeName, byval message)
		
	end sub
end class


class aspUnitTestResult
	' fields
	dim cTests, cPassed, cFailed, cErrors
	dim cTestCases
	
	' properties
	public property get Tests()
		set Tests = cTests
	end property
	
	public property get Passed()
		set Passed = cPassed
	end property
	
	public property get Failed()
		set Failed = cFailed
	end property
	
	public property get Errors()
		set Errors = cErrors
	end property
	
	public property get TestCases()
		set TestCases = cTestCases
	end property
	
	
	' constructor and desctructor
	private sub class_initialize()
		set cTests = new aspUnitCollection
		set cPassed = new aspUnitCollection
		set cFailed = new aspUnitCollection
		set cErrors = new aspUnitCollection
		set cTestCases = new aspUnitCollection
	end sub
	
	private sub class_terminate()
		cTests.clear()
		cPassed.clear()
		cFailed.clear()
		cErrors.clear()
		cTestCases.clear()
	end sub
end class


class aspUnitCollection
	' fields
	dim aCollection()
	
	
	' properties
	public property get Collection()
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
%>