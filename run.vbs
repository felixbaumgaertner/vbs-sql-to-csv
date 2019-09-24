' Configure variables here
sql = "SELECT * FROM your_table"
csvFilePath = "path\to\destination\results.csv"
seperator = ";"
server = "192.168.178.100"
database = "your_database"
user = "your_user"
password = "your_password"



Set connect = CreateObject("ADODB.Connection")
connect.ConnectionString = "Provider=SQLOLEDB.1;Data Source=" & server & ";Initial Catalog=" & database & ";user id ='" & user & "';password='" & password & "';Trusted_Connection=True;"
connect.Open
Set resultSet = connect.Execute(sql)


' Create new CSV file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCSVFile = objFSO.CreateTextFile(csvFilePath, True)

' Get all column names for reference
colCounter = resultSet.Fields.Count
i = 1
For each field in resultSet.Fields
	If i = colCounter Then
		objCSVFile.Write field.Name
	Else
		objCSVFile.Write field.Name & seperator
	End If
	i = i + 1
Next
objCSVFile.WriteLine

' Loop through query result
While NOT resultSet.EOF
	j = 1
    For Each field In resultSet.Fields
		If j = colCounter Then
			objCSVFile.Write field.Value
		Else
			objCSVFile.Write field.Value & seperator
		End If
		j = j + 1
    next

    resultSet.MoveNext
    objCSVFile.WriteLine
Wend

' Clean script end
resultSet.Close
Set resultSet = Nothing
connect.Close
Set connect = Nothing
