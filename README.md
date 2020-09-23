<div align="center">

## A Connectionless Recordset


</div>

### Description

This code demonstrates

how to use a connectionless recordset.
 
### More Info
 
The program uses ADO.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Syrob](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syrob.md)
**Level**          |Intermediate
**User Rating**    |3.0 (15 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syrob-a-connectionless-recordset__1-32333/archive/master.zip)





### Source Code

```
Option Explicit
'an object variable
Dim RsD As ADODB.Recordset
'a boolean variable
Dim f As Boolean
Public Function GetNames() As String()
'an object variable
Dim rs As ADODB.Recordset
'an object variable to hold a reference
'to an instance of the clsDB class
Dim objDB As clsDB
'a dynamic array
Dim a() As String
'an integer variable
Dim i As Integer
'a string variable
Dim l As String
'get a reference to the objDB object
Set objDB = New clsDB
'execute GetData function on the objDB
'object to get a reference to the recordset
Set rs = objDB.GetData
'extract the data from the recordset
'and applay to them a business rule
Do Until rs.EOF
'build a text
 l = rs!titleofcourtesy
 l = l & Left(Trim(rs!FirstName), 1) & "."
 l = l & rs!lastname
'resize the array
 ReDim Preserve a(i)
'populate the array
 a(i) = l
'set a size of the array
 i = i + 1
'call MakeRs method
 MakeRs rs!Notes, rs!employeeid
'move to the next record
 rs.MoveNext
Loop
'assign the array to a GetNames function
GetNames = a
End Function
Private Sub MakeRs(strData As String, intID As Integer)
'if the RsD recordset does not exist
'then create the recordset
If Not f Then
Set RsD = New Recordset
 With RsD
 .Fields.Append "ID", adInteger
 .Fields.Append "Notes", adBSTR
 .Open
 End With
'set a flag to indicate that
'the recordset exists
 f = True
End If
'dump the data into the recordset
With RsD
 .AddNew
 .Fields("ID") = intID
 .Fields("Notes") = strData
 .Update
End With
End Sub
Public Property Get Notes() As ADODB.Recordset
'assign the recordset to the property
 Set Notes = RsD
End Property
Public Function SaveData(rs As ADODB.Recordset, f() As Integer) As Boolean
'an object variable
Dim objDB As clsDB
'a string variable
Dim l As String
'an integer variable
Dim i As Integer
'error handler
On Error GoTo ErrSave
'get a reference to the objDB object
Set objDB = New clsDB
'make sure that we start from the first record
rs.MoveFirst
'extract the updated data from
'the recordset and pass them
'to UpdateRecords function
Do Until rs.EOF
For i = 0 To UBound(f)
'check if the data were changed
 If f(i) = rs!ID Then
 l = rs!Notes
'applay a business rule
 If l = "" Then
 l = "Notes deleted " & Date
 End If
 If Not objDB.UpdateRecords(rs!ID, l) Then
'indicate a failure
 SaveData = False
 Exit Function
 End If
 End If
 Next i
 rs.MoveNext
Loop
'put an end to the object
Set rs = Nothing
'we are happy
SaveData = True
Exit Function
ErrSave:
'indicate a failure
SaveData = False
End Function
```

