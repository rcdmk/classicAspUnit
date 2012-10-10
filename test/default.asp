<% option explicit %>
<!--#include file="../aspUnit.class.asp" -->
<!--#include file="testDB.class.asp" -->
<%
' Helpers
function createUser(byval id, byval name, byval email)
	dim user(2,0)
	
	user(0, 0) = id
	user(1, 0) = name
	user(2, 0) = email
	
	createUser = user
end function


' Tests
dim usersDB, testContext, results

set testContext = new aspUnit

sub testSetup()
	set usersDB = new testDB

	usersDB.TableName = "users"
end sub

sub testTeardown()
	set usersDB = nothing
end sub


testSetup

dim otest
set otest = testContext.addTestCase("User CRUD")

otest.Setup("testSetup")
otest.Teardown("testTeardown")

' #####
dim oTestMethod
set oTestMethod = otest.addTest("UserDB is a testDB instance")

oTestMethod.AssertExists usersDB, ""
oTestMethod.AssertIsA usersDB, "testDB", ""


' #####
set oTestMethod = otest.addTest("UserDB's Table Name is set to users")
oTestMethod.AssertEquals usersDB.TableName, "users", ""


' #####
set oTestMethod = otest.addTest("UserDB adds a user")

dim oldCount, newUser
oldCount = usersDB.Count

newUser = createUser(1, "Jhon", "jhon@domain.com")

usersDB.Add newUser

oTestMethod.AssertEquals usersDB.Count, oldCount + 2, ""


set results = testContext.run

set oTestMethod = nothing
set otest = nothing
set usersDB = nothing
set testContext = nothing

dim testCase, test, testMethod
dim testCases, tests, testMethods

set testCases = results.TestCases
%>
<!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="UTF-8">
	<title>Classic ASP Unit Testing Framework</title>
	<link rel="stylesheet" href="css/style.css"/>
	<script type="text/javascript" src=""></script>
</head>
<body>
	<h1>Classic ASP Unit Testing Framework</h1>
	<h2>Test Cases: <%= testCases.Count %>, Tests: <%= results.Tests %>, Passed: <%= results.Passed %>, Failed: <%= results.Failed %>, Error: <%= results.Errors %></h2>
	
	<table>
		<tr>
			<th colspan="3">Test Cases</th>
		</tr>
		<%
		for each testCase in testCases.Collection
			%>
			<tr>
				<td colspan="2"><%= testCase.Name %></td>
				<td class="status <%= testCase.Status %>"><%= testCase.Status %></td>
			</tr>
			<tr>
				<td colspan="3" class="subtitle">Tests:</td>
			</tr>
			<%
			for each test in testCase.Tests.Items
				%>
				<tr>
					<td><%= test.Name %></td>
					<td><%= test.Output %></td>
					<td class="status <%= test.Status %>"><%= test.Status %></td>
				</tr>
				<%
			next
		next
		%>
	</table>	
</body>
</html>