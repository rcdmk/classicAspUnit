<%
' Emulates a database class for user persistence

class testDB
	' Fields
	dim aConn()
	dim sTableName
	
	
	' Properties
	public property get TableName()
		TableName = sTableName
	end property
	
	public property let TableName(value)
		sTableName = value
	end property
	
	
	public property get Count()
		Count = recordCount()
	end property
	
	
	' Constructor and destructor
	private sub class_initialize()
		redim aConn(-1)
	end sub
	
	private sub class_terminate()
		redim aConn(-1)
	end sub
	
	
	' Public methods
	public function GetAll()
		GetAll = aConn
	end function
	
	
	public function GetOne(iIndex)
		if isObject(aConn(iIndex)) then
			set GetOne = aConn(iIndex)
		else
			GetOne = aConn(iIndex)
		end if
	end function
	
	
	public function Add(obj)
		arrayPush aConn, obj
		Add = true
	end function
	
	
	public function Update(obj)
		dim index, result
		
		index = getIndex(obj)
		result = false
		
		if index >= 0 then
			setObject index, obj
			
			result = true
		end if
		
		Update = result
	end function
	
	
	public function Remove(obj)
		dim i, index, total, result
		i = 0
		total = ubound(aConn)
		result = false
		
		index = getIndex(obj)
		
		' If the object was found
		if index >= 0 then
			'Destroy the object
			set aConn(index) = nothing
			
			' Shifts the objecs above this index one index less
			for i = index to total
				set aConn(i) = aConn(i + 1)
			next
			
			' Destroy the las item to be removed
			set aConn(total) = nothing
			
			' Shorten the array, removing the last item
			redim preserve aConn(total - 1)
			
			result = true
		End If
		
		Remove = result
	end function
	
	
	public sub Clear()
		for each obj in aConn
			Remove obj
		next
	end sub
	
	
	
	' Private methods
	private function recordCount()
		recordCount = ubound(aConn) + 1
	end function
	
	
	private function getIndex(obj)
		dim i, index
		index = -1
		
		do while i < total
			if aConn(i) = obj then
				index = i
				exit do
			end if
			
			i = i + 1
		loop
		
		getIndex = index
	end function
	
	
	private sub setObject(index, obj)
		if isObject(obj) then
			set aConn(index) = obj
		else
			aConn(index) = obj
		end if
	end sub
	
	
	' Pushes (adds) a value to an array, expanding it
	private function arrayPush(byref arr, byref value)
		redim preserve arr(ubound(arr) + 1)
		if isobject(value) then
			set arr(ubound(arr)) = value
		else
			arr(ubound(arr)) = value
		end if
		ArrayPush = arr
	end function
end class
%>