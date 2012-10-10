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


dim test, testMethod

set test = testContext.addTestCase("User CRUD")

test.Setup("testSetup")

set testMethod = test.addTest("UserDB is a testDB instance")

testMethod.AssertExists usersDB
testMethod.AssertIsA usersDB, "testDB", ""

test.Teardown("testTeardown")

set results = testContext.run

set usersDB = nothing
set testContext = nothing
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
	<h2>Tests: <%= results.Tests.Count %>, Passed: <%= results.Passed.Count %>, Failed: <%= results.Failed.Count %>, Error: <%= results.Errors.Count %></h2>
	
	
</body>
</html>