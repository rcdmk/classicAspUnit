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


testSetup()

dim oTest
set oTest = testContext.addTestCase("User Administration")

'oTest.Setup("testSetup")
'oTest.Teardown("testTeardown")

' #####
dim oTestMethod
set oTestMethod = oTest.addTest("UserDB is a testDB instance")

oTestMethod.AssertExists usersDB, ""
oTestMethod.AssertIsA usersDB, "testDB", ""


' #####
set oTestMethod = oTest.addTest("UserDB's Table Name is set to users")
oTestMethod.AssertEquals usersDB.TableName, "users", ""


' #####
set oTestMethod = oTest.addTest("UserDB adds a user")

dim oldCount, newUser
oldCount = usersDB.Count

newUser = createUser(1, "Jhon", "jhon@domain.com")

usersDB.Add newUser

oTestMethod.AssertEquals usersDB.Count, oldCount + 1, ""

' #####
set oTestMethod = oTest.addTest("UserDB returns the added user")

newUser = createUser(2, "Joe", "joe@domain.com")

usersDB.Add newUser

' usersDB.GetOne accepts the record index as a parameter
oTestMethod.AssertEquals usersDB.GetOne(1), newUser, ""


set oTest = testContext.addTestCase("User Login")


testTeardown()

set results = testContext.run
results.Update

set oTestMethod = nothing
set oTest = nothing
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
			<th class="right">Status</th>
		</tr>
		<%
		if testCases.Count > 0 then
			for each testCase in testCases.Collection
				%>
				<tr class="title">
					<td colspan="3"><%= testCase.Name %></td>
					<td class="status <%= testCase.Status %>"><%= testCase.Status %></td>
				</tr>
				<tr>
					<td class="indent">&nbsp;</td>
					<td colspan="3" class="subtitle">Tests</td>
				</tr>
				<%
				if testCase.Tests.Count > 0 then
					for each test in testCase.Tests.Items
						%>
						<tr>
							<td class="indent">&nbsp;</td>
							<td><%= test.Name %></td>
							<td><%= test.Output %></td>
							<td class="status <%= test.Status %>"><%= test.Status %></td>
						</tr>
						<%
					next
				else
					%>
					<tr>
						<td class="indent">&nbsp;</td>
						<td colspan="2">No tests</td>
						<td class="status Inconclusive">Inconclusive</td>
					</tr>
					<%
				end if
				%>
				<tr>
					<td colspan="4" class="indent">&nbsp;</td>
				</tr>
				<%
			next
		else
			%>
			<tr>
				<td colspan="3">No test cases</td>
				<td>&nbsp;</td>
			</tr>
			<%
		end if
		%>
	</table>	
</body>
</html>