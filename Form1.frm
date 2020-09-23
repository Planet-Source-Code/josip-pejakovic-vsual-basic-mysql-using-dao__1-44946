VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "MySQL with DAO - by Josip Pejakovic - jpejakovic@yahoo.com - http://jp-net.web1000.com"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Showing records"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4575
      Begin MSComctlLib.ListView Lv 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton delete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adding records"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Add record"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Connecting to MySQL database using myODBC 3.5 & DAO 3.51
'by Josip Pejakovic - jpejakovic@yahoo.com - http://jp-net.web1000.com - ICQ# 127475388

'THIS CODE IS CREATION OF CROATIAN PROGRAMMER

'Instructions:
' 1. create database called dbTest
' 2. inside dbTest create table called TestTable
' 3. create 2 fields: Name (varchar 50) and Address (varchar 100)
' 4. open Control Panel - ODBC Data Sources (32bit)
' 5. create new Data Source and name it to dbTest (DSN), don't put any data in fields when you creating DSN, leave it empty


Dim ws As Workspace
Dim conn As Connection
Dim rs As Recordset
Dim lv_item As ListItem
Private Sub Command1_Click()

'This is for adding records in table
Set rs = conn.OpenRecordset("TestTable")

With rs
rs.AddNew
rs!Name = Text1
rs!Address = Text2
rs.Update
MsgBox "Record saved!!", vbInformation
End With

Command2_Click 'refresh listview
End Sub

Private Sub Command2_Click()
'This is for reading records form our database and putting them into listview
Lv.ListItems.Clear
Set rs = conn.OpenRecordset("SELECT * FROM TestTable")
Do Until rs.EOF
Set lv_item = Lv.ListItems.Add(, , rs!Name)
lv_item.SubItems(1) = rs!Address
rs.MoveNext
Loop
End Sub

Private Sub delete_Click()
'This is for deleting records from table
'In this case we will only delete records that we have selected from ListView
Set rs = conn.OpenRecordset("DELETE FROM TestTable WHERE Name = '" & Lv.SelectedItem & "' AND Address = '" & Lv.SelectedItem.SubItems(1) & "'")
Command2_Click
End Sub

Private Sub Form_Load()
Dim str As String

Set ws = DBEngine.CreateWorkspace("", "root", "", dbUseODBC)
str = "odbc;DRIVER={MySQL ODBC 3.51 Driver};" _
& "SERVER=localhost;" _
& " DATABASE=dbTest;" _
& "UID=root;PWD=; OPTION=35"

'SERVER - host name
'DATABASE - name of your database (in this case the database is dbTest)
'UID - user name for accessing, default is root
'PWD - password for accessing to database (in this case the password is empty because I don't define any passwords

Set conn = ws.OpenConnection("dbTest", dbDriverNoPrompt, False, str) 'here we opening connection

'now we will create column headers for listview
With Lv.ColumnHeaders
.Add , , "Name"
.Add , , "Address", Lv.Width * 0.7
End With

'after that we will read data from our database
Command2_Click
End Sub
