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
		redim aConn(2, -1)
	end sub
	
	private sub class_terminate()
		redim aConn(2, -1)
	end sub
	
	
	' Public methods
	public function GetAll()
		GetAll = aConn
	end function
	
	
	public function GetOne(id)
		dim i, j, item(2, 0), result, total
		
		total = recordCount()
		
		do while i < total
			if aConn(0, i) = id then
				for j = 0 to 2
					item(j, 0) = aConn(j, i)
				next
				
				result = item				
				exit do
			end if
			
			i = i + 1
		loop
		
		GetOne = result
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
	
	
	public function Remove(id)
		dim obj, i, j, index, total, cols, result
		result = false
		
		total = recordCount()
		cols = 2
		
		obj = GetOne(id)
		
		index = getIndex(obj)

		' If the object was found
		if index >= 0 then
			if total > 1 then
				' Shifts the objecs above this index one index less
				for i = index to total - 1
					for j = 0 to cols
						aConn(j, i) = aConn(j, i + 1)
					next
				next
			end if
			
			' Shorten the array, removing the last item
			redim preserve aConn(cols, total - 2)
			
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
		recordCount = ubound(aConn, 2) + 1
	end function
	
	
	private function getIndex(obj)
		dim i, index, total
		i = 0
		index = -1
		
		total = recordCount()
		
		do while i < total
			if aConn(0, i) = obj(0, 0) then
				index = i
				exit do
			end if
			
			i = i + 1
		loop
		
		getIndex = index
	end function
	
	
	private sub setObject(index, obj)
		dim i, cols
		cols = ubound(aConn, 1)
		
		for i = 0 to cols
			aConn(i, index) = obj(i, 0)
		next
	end sub
	
	
	' Pushes (adds) a value to an array, expanding it
	private function arrayPush(byref arr, byref value)
		dim i, rows
		rows = ubound(arr, 2) + 1
		redim preserve arr(2, rows)
		
		for i = 0 to ubound(arr, 1)
			arr(i, rows) = value(i, 0)
		next
		
		ArrayPush = arr
	end function
end class
%>