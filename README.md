Classic ASPUnit
===============

A classic ASP unit framework for helping in testing classic asp code.

# Usage
<!-- languages: vbscript -->
	
Instantiate the context:

    set testContext = new aspUnit
	
Create a test case:

	set oTest = testContext.addTestCase("User Administration")

Make assertions:

	oTestMethod.AssertExists usersDB, "optional message override: {1}" ' accepts a wildcard marks for the parammeters
	oTestMethod.AssertIsA usersDB, "testDB", "" ' leave blank for default message

You can also create test setups and teardowns to be executed before and after each test for a `Test Case`:

	sub testSetup()
		set usersDB = new testDB

		usersDB.TableName = "users"
		
		set newUser = new User
		newUser.id = 1
		newUser.name = "Bob"
		
		usersDB.add newUser
	end sub

	sub testTeardown()
		set usersDB = nothing
	end sub
	
... and then pass the method names for the Test Case:

	oTest.Setup("testSetup")
	oTest.Teardown("testTeardown")
	
This would work too:

	oTest.Setup("myGlobalObject.MyMethod(1, ""param2"", true)")
	
> **Warning:** This uses `Execute` to run the code and will accpect any executable code string like `"myVar = 1"` or `"myFunction() : myOtherFunction()"`


To run and get the results of the tests:

	set results = testContext.run
	results.Update ' This will update the test counters for passed, failed and errors

Then you can have access to the results and write any view you want:

	Response.Write "Test Cases: " & results.TestCases.Count & "<br>"
	Response.Write "Tests runned: " & results.Tests & ", "
	Response.Write "Tests passed: " & results.Passed & ", "
	Response.Write "Tests failed: " & results.Failed & ", "
	Response.Write "Tests errored: " & results.Errors & "<br><br>"

	' loop the testCases
	for each testCase in result.TestCases.Items
		Response.Write "-> Test Case: " & testCase.Name & "(" & testCase.Status & ")<br>"
		
		' loop the tests
		for each test in testCase.Tests.Items
			Response.Write "--> Test: " & test.Name & "<br>"
			Response.Write "----> " & test.Output & "(" & test.Status & ")<br>"
		next
	next
	
>There is a template view with the source in the test folder.