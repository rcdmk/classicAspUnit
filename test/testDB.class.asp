<%
' Emulates a database class for user persistence

class testDB
	dim oConn()
	
	private sub class_initialize()
		redim oConn(-1)
	end sub
	
	private sub class_terminate()
	end sub
end class
%>