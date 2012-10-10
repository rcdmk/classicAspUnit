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


dim otest, otestMethod

set otest = testContext.addTestCase("User CRUD")

otest.Setup("testSetup")
otest.Teardown("testTeardown")

set otestMethod = otest.addTest("UserDB is a testDB instance")

otestMethod.AssertExists usersDB
otestMethod.AssertIsA usersDB, "testDB", ""


set results = testContext.run

set otestMethod = nothing
set otest = nothing
set usersDB = nothing
set testContext = nothing

dim testCase, test, testMethod
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
	<h2>Test Cases: <%= results.TestCases.Count %>, Tests: <%= results.Tests.Count %>, Passed: <%= results.Passed.Count %>, Failed: <%= results.Failed.Count %>, Error: <%= results.Errors.Count %></h2>
	
	<table>
		<%
		for each testCase in results.TestCases
			%>
			<tr>
				<th colspan="2"><%= testCase.Name %></th>
				<th class="<%= testCase.Status %>"><%= testCase.Status %></th>
			</tr>
			<%
			for each test in results.Tests
				%>
				<tr>
					<th colspan="2"><%= test.Name %></th>
					<th class="<%= test.Status %>"><%= test.Status %></th>
				</tr>
				<%
				for each testMethod in test.TestMethods
					%>
					<tr>
						<td><%= testMethod.Name %></td>
						<td><%= testMethod.Output %></td>
						<td class="<%= testMethod.Status %>"><%= testMethod.Status %></td>
					</tr>
					<%
				next
			next
		next
		%>
	</table>	
</body>
</html>