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
	
	public sub Run()
	end sub
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
%>