VERSION 5.00
Begin VB.UserControl sqlSDBC 
   BackColor       =   &H00000000&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   Palette         =   "UserControl1.ctx":0000
   Picture         =   "UserControl1.ctx":3556
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "UserControl1.ctx":3B98
   Begin VB.Image Image1 
      Height          =   915
      Left            =   -4680
      Picture         =   "UserControl1.ctx":3EAA
      Top             =   -2880
      Width           =   1740
   End
End
Attribute VB_Name = "sqlSDBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SQL Server Database Control 1.0 for VB 6.0
'(If you need VB.NET version, contact me
'See the Ream Me.txt for more details
''''''''''''''''''''''''''''''''''''''''''''
'This code written by Issam Hijazi''''''''''
'It help for projects that use to connect  '
'to MS SQL Server and work with database   '
'stuff! Everything as I guess you'll find  '
'here! If there is anything you can't find '
'please contact me! See the Read Me.txt    '
'for more information                      '
''''''''''''''''''''''''''''''''''''''''''''
'You can go now to File => Make sqlSDBC.ocx
'to start using it imeditly in your project
'after you add it to your project after you
'make it .ocx file or read the following
'code carefully and try to understand every
'little char in it, it helps a lot!
''''''''''''''''''''''''''''''''''''''''''''
'PLEASE PLEASE PLEASE
'Rate My Control (Vote For It!)
'PLEASE PLEASE PLEASE
'''''''''''''''''''''''''''''''''''''''''''''


'The folowing publics for easy when use the .ocx file
'you will see them later in the code
Public Enum Way
UseLike = 1
UseEquel = 2
End Enum

'Just to make easy choice, will help when use .ocx file
'you will see them later in the code
Public Enum MoveWay
MoveNext = 1
MoveBack = 2
MoveFirst = 3
MoveLast = 4
End Enum

'this will control the sql server
Dim SQLS As New SQLDMO.SQLServer
'this will control the sql server
Dim SQLS3 As New SQLDMO.SQLServer
'this will allow you to access to table and do some commands there
Dim RecordsetT As New ADODB.Recordset
'this will all you to access the database which will allow you to access the tables
Dim DatabaseT As New ADODB.Connection
'this will help us to upload the files and pictures to the table fields
Dim MStream As New ADODB.Stream
'just some variables will used
Dim ConnectionString As String
Dim ErrorNumber
Dim ErrorDescription
Dim User As String, Pass As String, SrvName As String, DBName As String, SQLSta As String

'here we open the connection to the database
Public Function OpenConnection(Username As String, Password As String, ServerName As String, DatabaseNameIs)
On Error GoTo BestHandler

'this variable we write it in the first, will used later and here
ConnectionString = "Server=" & ServerName & ";Provider=SQLOLEDB;UID=" & Username & ";PWD=" & Password & ";database=" & DatabaseNameIs & ";"

'this going to open the connection by the sentence above
DatabaseT.Open ConnectionString

'variables
User = Username
Pass = Password
SrvName = ServerName
DBName = DatabaseTName

'you will see much, its for handeling errors without popup msg boxes, just as you need, with numbers to specify yours!
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'this will find any record in the table, after you give the column name and the way of finding (= or like) and tell what you want to find
Public Function FindRecord(ColName As String, ByVal FindWay As Way, Text As String)
On Error GoTo BestHandler

'this line prevent from errors when searching for records contains (') in any part, try to find a record with (') like Moh'd and see what will happen!, this will prevent the errors!
Text = Replace(Text, "'", "''")

'its the find way, remember these publics, return to top to see them
If FindWay = UseEquel Then
'you know, its the certeria used in sql, we use [] if we have column caption from two words like (Full Name), if you don't use the [] in this case, error will occured
RecordsetT.Find "[" & ColName & "]" & " ='" & Text & "'", 0, adSearchForward, 1

Else

RecordsetT.Find "[" & ColName & "]" & " Like '" & Text & "%'", 0, adSearchForward, 1
'''
'Note: For your information, if you need to find records
'between two things, I mean by things like (members ranks
'between 5 to 100) you can use the following certeria:
'("where Xfield between '" & Xtext1 & "' and '" & Xtext2 &"'")
'thats just for your information!
'''

End If
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'here we change any field data, like update!
Public Function ChangeFieldData(FieldIndexOrName, NewData)
On Error GoTo BestHandler
RecordsetT.Update FieldIndexOrName, NewData
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function
'here we add new record
Public Function AddNewRecord()
On Error GoTo BestHandler
RecordsetT.AddNew
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function
'here we cancel any active process, like updating
Public Function CancelOperation()
On Error GoTo BestHandler
RecordsetT.CancelUpdate
RecordsetT.Cancel
RecordsetT.CancelBatch adAffectCurrent
RecordsetT.Requery -1
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description


End Function
'here we delete record after we select it by find or anything you use for that
Public Function DeleteRecord()
On Error GoTo BestHandler
RecordsetT.Delete adAffectCurrent
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function
'here we move the selector (back, next, first or last), return to top to see the publics
Public Function MoveRecord(ByVal Move As MoveWay)
On Error GoTo BestHandler
If Move = MoveFirst Then
RecordsetT.MoveFirst
ElseIf Move = MoveLast Then
RecordsetT.MoveLast
ElseIf Move = MoveNext Then
RecordsetT.MoveNext
ElseIf Move = MoveBack Then
RecordsetT.MovePrevious
End If
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'we use this to open a database and start our commands like delete and add new and find...
'this done after opening a connection!
Public Function OpenRecordset(SQLStatment As String)
On Error GoTo BestHandler



RecordsetT.Open SQLStatment, DatabaseT, adOpenKeyset, adLockOptimistic

SQLSta = SQLStatment

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'refresh, use it after add new record or delete, anywhere!
Public Function Refresh()
On Error GoTo BestHandler
RecordsetT.Requery -1
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'here is the good stuff start, upload file to database table into binary field!
Public Function SaveFileToDB(FilePath As String, FieldIndexOrName)
On Error GoTo BestHandler
       
       'this line specify that the field is binary, which used for files!
       MStream.Type = adTypeBinary
              
       'open the streem, don't work without open it!
       MStream.Open
       
       'we first load the file, the FilePath is specified by the user like "C:\doc\new.xls"
       MStream.LoadFromFile (FilePath)
       
       'here we specify which field will be used for saving the file, by name or index!
       RecordsetT.Fields(FieldIndexOrName).Value = MStream.Read
       
       'here we save it database!
       RecordsetT.Update
       
       'here we finish the work, it must be closed for the next job!
     MStream.Close
     
       Refresh
       
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'we upload file in the previos public, right? now how we download it? here we go!
Public Sub SaveDBToFile(FieldIndexOrName, FilePath As String)
On Error GoTo BestHandler

'as above
MStream.Type = adTypeBinary

'as above
MStream.Open

'no we going to write out the file from the field, unlike above!
MStream.Write RecordsetT.Fields(FieldIndexOrName).Value

'select out desire path and save it!
MStream.SaveToFile FilePath, adSaveCreateOverWrite

'as above
MStream.Close

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Sub

'Ok, you want to load the file from DB directly to picture object,
'oops, you use NTFS and have no access to disk? how will you download
'the file to disk then load it to the picture box! thats a real problem
'but we here have good solution!

Public Sub LoadDBPicToObject(FieldIndexOrName, ObjectName As Object)
On Error GoTo BestHandler

'check if the stream is opened or not
If MStream.State = adStateOpen Then MStream.Close
'as obove
MStream.Type = adTypeBinary

'as above
MStream.Open

'as above
MStream.Write RecordsetT.Fields(FieldIndexOrName).Value

'not as above! why? because we used here Environ$ which is solution for NTFS if you don't have a permission on the disk!
'search the internet to learn more about Environ
MStream.SaveToFile Environ$("TEMP") & "\temp", adSaveCreateOverWrite

'he we load out picture from Environ$ with no problem
ObjectName.Picture = LoadPicture(Environ$("TEMP") & "\temp")

'as above
MStream.Close
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Sub

'thats for cool look! lol! just if you make it .ocx and used it in your project
Private Sub UserControl_Resize()
On Error GoTo BestHandler
Height = 420
Width = 420
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Sub

'here we close our tables
Public Function CloseRS()
On Error Resume Next

'check if its open first, thats make problem if we called the close the recordest twice without open between them, so this is solution
If RecordsetT.State = adStateOpen Then

RecordsetT.Close

End If

If DatabaseT.State = adStateOpen Then

DatabaseT.Close

End If


ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'here is the DMO library control work starts!
'you can get the MS SQL Server state from thiese lines
Public Property Get SQLServerStatus(SQLServerName) As Variant
On Error GoTo BestHandler
SQLS3.Name = SQLServerName
If SQLS3.Status = SQLDMOSvc_Running Then
SQLServerStatus = "Running"
ElseIf SQLS3.Status = SQLDMOSvc_Paused Then
SQLServerStatus = "Paused"
ElseIf SQLS3.Status = SQLDMOSvc_Stopped Then
SQLServerStatus = "Stopped"
ElseIf SQLS3.Status = SQLDMOSvc_Unknown Then
SQLServerStatus = "Unknown"
ElseIf SQLS3.Status = SQLDMOSvc_Continuing Then
SQLServerStatus = "Continuing"
ElseIf SQLS3.Status = SQLDMOSvc_Pausing Then
SQLServerStatus = "Pausing"
ElseIf SQLS3.Status = SQLDMOSvc_Starting Then
SQLServerStatus = "Starting"
ElseIf SQLS3.Status = SQLDMOSvc_Stopping Then
SQLServerStatus = "Stopping"
End If
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Property

'you know EOF & BOF stuff, to see if you are out of the records
Public Property Get IfBOForEOF() As Boolean
On Error GoTo BestHandler
If RecordsetT.BOF = True Or RecordsetT.EOF = True Then
IfBOForEOF = True
End If
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Property

'load the field text to any object or variable
Public Property Get GetFieldData(FieldIndexOrName) As Variant
On Error GoTo BestHandler
GetFieldData = RecordsetT.Fields(FieldIndexOrName)
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Property

'do you to check if the account you use is working on the SQL Server or not? here is little check up
Public Function CheckAccount(ServerName As String, Username As String, Password As String) As Boolean
On Error GoTo BestHandler
Dim SQLS2 As New SQLDMO.SQLServer

On Error GoTo BestHandler
SQLS2.Name = ServerName

On Error Resume Next

SQLS2.Connect ServerName, Username, Password


CheckAccount = True

Stop
BestHandler:
CheckAccount = False
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Start the SQL Server
Public Function StartSQLServer(ServerName As String, Username As String, Password As String)
On Error GoTo BestHandler
SQLS.Start False, ServerName, Username, Password
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Pause the SQL Server
Public Function PauseSQLServer(ServerName As String)
On Error GoTo BestHandler
SQLS.Name = ServerName
SQLS.Pause
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Continue the SQL Server
Public Function ContinueSQLServer(ServerName As String)
On Error GoTo BestHandler
SQLS.Name = ServerName
SQLS.Continue
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Stop the SQL Server
Public Function StopSQLServer(ServerName As String)
On Error GoTo BestHandler
SQLS.Name = ServerName
SQLS.Stop
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Ye this delete any database on the SQL Server!, but done after you are connectd to the server
Public Function DeleteDatabase(DatabaseTName As String)
On Error GoTo BestHandler
SQLS.KillDatabase DatabaseTName
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Add new dattbase to SQL Server, but the file must be ended with .MDF extension
Public Function AddDatabase(DatabaseTName As String, DatabaseTFileMDF As String)
On Error GoTo BestHandler
SQLS.AttachDBWithSingleFile DatabaseTName, DatabaseTFileMDF
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'Here we connect to SQL Server to do some commands like Delete Datbase and Add Database
Public Function ConnectToSQLServer(ServerName As String, Username As String, Password As String)
On Error GoTo BestHandler
SQLS.Connect ServerName, Username, Password
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'This will disconnect you from the SQL Server, so you can't add new database or delete one...
Public Function DisconnectFromSQLServer()
On Error GoTo BestHandler
SQLS.Disconnect
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'You not sure if you connected or not, check your connection here
'Its not really made for testing the connection
'but it works perfectly (its just idea from me)
Public Property Get IsConnected() As Boolean
On Error GoTo BestHandler
If SQLS.IsPackage = SQLDMO_Unknown Then Width = 735 Else Height = 255


IsConnected = True

If Err.Number = -2147201022 Then IsConnected = False Else IsUserLogin = True
BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Property

'Repair your damaged database, make access to faster!
Public Function RepairDatabase(DatabaseTName As String)
On Error GoTo BestHandler
SQLS.Databases(DatabaseTName).CheckAllocations SQLDMORepair_None

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'The following lines backup your database to any path you desire
Public Function BackupDatabaseToFile(DatabaseTName As String, Path As String)
On Error GoTo BestHandler
On Error Resume Next
'this used to creat the back utility by DMO
Dim BackMeUp As SQLDMO.Backup
'the following line must entered, don't work without it
Set BackMeUp = New SQLDMO.Backup
'variables
Dim DatabaseTFileName As String


'the above variable which will used for the backed up file
DatabaseTFileName = Environ$("TEMP") & "\" & DatabaseTName & ".bak"

'here we select which database
BackMeUp.Database = DatabaseTName

'the file path is selected here
BackMeUp.Files = DatabaseTFileName

'start the backing up, SQLS is used as the connection, you must be connected!
BackMeUp.SQLBackup SQLS


'move the set to your desire location
FileCopy DatabaseTFileName, Path & "\" & DatabaseTName & ".bak"
Kill Environ$("TEMP") & "\" & DatabaseTName & ".bak"


BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'we backed up a database before, right? how we restore it? here we go
Public Function RestoreDatabaseFromFile(DatabaseTName As String, Path As String)
On Error GoTo BestHandler
'the resote utility as object
Dim oRestore As SQLDMO.Restore
'the following line must be entered, don't work without it
Set oRestore = New SQLDMO.Restore

'get the file we wanna restore
FileCopy Path & "\" & DatabaseTName & ".bak", Environ$("TEMP") & "\" & DatabaseTName & ".bak"

'enter database name
oRestore.Database = DatabaseTName

'the file path set here
oRestore.Files = Environ$("TEMP") & "\" & DatabaseTName & ".bak"

'start the resorting up, the SQLS is our connection as object, you must be connected!
oRestore.SQLRestore SQLS

'clean out work
Kill Environ$("TEMP") & "\" & DatabaseTName & ".bak"

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description
End Function

'this is the error handler by number
'this help you out like this:
'[
'If sqlSDBC1.ErrorNum = -54845484 Then
'Msgbox "Wrong Username!"
'End if
']
'its just example

Public Property Get ErrorNum() As Variant
ErrorNum = ErrorNumber
End Property

'this is error handler by name
'so you don't have to describe you problem
'in the msgbox or whereever you show your error
Public Property Get ErrorDes() As Variant
ErrorDes = ErrorDescription
End Property

'this show about me dialog box!
Public Function AboutMe()
Dim oFrm As About
Set oFrm = New About
oFrm.Show vbModal

End Function

'if you want to bind the table to MSFLEXGRID control, this is what you need
'Note: If you need to use FlexGrid control its ok,
'use the same code, but the table must have primary key
'to be able to bound to FlexGrid control

Public Function BindToMSHFlexGrid(ObjectName As Object)
On Error GoTo BestHandler

RecordsetT.Close

RecordsetT.Open SQLSta, DatabaseT, adOpenKeyset, adLockOptimistic

Set ObjectName.DataSource = RecordsetT
ObjectName.Refresh
RecordsetT.Requery -1

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description
End Function

'This will bind any field to object like TextBox or Label
Public Function BindToObject(ObjectName As Object, DataFieldName As String)
On Error GoTo BestHandler


Set ObjectName.DataSource = RecordsetT

ObjectName.DataField = DataFieldName
ObjectName.Refresh

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description

End Function

'If you wanna list all the databases in SQL Server this is what you need
'it works with object support the .AddItem future

Public Function ListDatabases(ObjectName As Object)
On Error GoTo BestHandler
Set RecordsetT = DatabaseT.Execute("sp_databases")
Do Until RecordsetT.EOF
ObjectName.AddItem (RecordsetT.Fields("Database_Name"))
RecordsetT.MoveNext
Loop

RecordsetT.Close

RecordsetT.Open SQLSta, DatabaseT, adOpenKeyset, adLockOptimistic
RecordsetT.Requery -1

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description
End Function

'as above, but list the tables in database
Public Function ListTables(ObjectName As Object)
On Error GoTo BestHandler

On Error Resume Next
RecordsetT.Close
ErrorNumber = ""
ErrorDescription = ""

RecordsetT.Open "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND OBJECTPROPERTY(OBJECT_ID(TABLE_NAME), 'IsMSShipped') = 0", DatabaseT, adOpenKeyset, adLockOptimistic

RecordsetT.Requery -1

Do Until RecordsetT.EOF
ObjectName.AddItem RecordsetT.Fields("TABLE_NAME")
RecordsetT.MoveNext
Loop

RecordsetT.Close



If SQLSta = "" Then
Else
RecordsetT.Open SQLSta, DatabaseT, adOpenKeyset, adLockOptimistic
RecordsetT.Requery -1
End If

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description
End Function

'as above, but list fields in table
Public Function ListFields(ObjectName As Object, TableName As String)
On Error GoTo BestHandler
Dim nulls As String
Dim cnt As Integer



Set RecordsetT = DatabaseT.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))

Do Until RecordsetT.EOF

cnt = cnt + 1


ObjectName.AddItem RecordsetT!column_name
RecordsetT.MoveNext
Loop

RecordsetT.Close

If SQLSta = "" Then
Else
RecordsetT.Open SQLSta, DatabaseT, adOpenKeyset, adLockOptimistic
RecordsetT.Requery -1
End If

BestHandler:
ErrorNumber = Err.Number
ErrorDescription = Err.Description
End Function

''''''''''''''''''''''''
'I'm sorry for any type error, I didn't check my notes
'back. Anyway, if you have any problem please contact me
'See the Ream Me.txt file first!
''''''''''''''''''''''''
