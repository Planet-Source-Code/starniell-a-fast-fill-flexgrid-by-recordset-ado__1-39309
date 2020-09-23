VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTestFastFillFlexGrid 
   Caption         =   "Test Fast Fill FlexGrid"
   ClientHeight    =   5715
   ClientLeft      =   1725
   ClientTop       =   1935
   ClientWidth     =   9855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9855
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   8640
      Picture         =   "Form1.frx":2AFA
      ScaleHeight     =   705
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   4890
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Grid"
      Height          =   585
      Left            =   4110
      TabIndex        =   3
      Top             =   450
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fill with FlexGrid.Additem metod"
      Height          =   585
      Left            =   6090
      TabIndex        =   2
      Top             =   450
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill with Rcs.getString() and FlexGrid.Clipmetod"
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   3765
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3645
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Tested by StarNiell®2002 - Naples - Italy - adinardo@libero.it"
      Height          =   345
      Left            =   4230
      TabIndex        =   6
      Top             =   5310
      Width           =   4395
   End
   Begin VB.Label Label2 
      Caption         =   "Very Slow method !!!!!!!!!!"
      Height          =   255
      Left            =   6270
      TabIndex        =   5
      Top             =   150
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Very Fast method !!!!! -  About 230%  + Fast!!!!"
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   180
      Width           =   3345
   End
End
Attribute VB_Name = "frmTestFastFillFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim lTimer As Long

Screen.MousePointer = vbHourglass
Command3_Click
MSFlexGrid1.Refresh
lTimer = Timer

MSFlexGrid1.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Test.mdb;Persist Security Info=False"
rs.Open "SELECT * FROM COMUNI", db, adOpenStatic, adLockReadOnly
rs.MoveFirst

MSFlexGrid1.Rows = rs.RecordCount + 1
MSFlexGrid1.Cols = rs.Fields.Count - 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
MSFlexGrid1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
MSFlexGrid1.Row = 1
MSFlexGrid1.Visible = True

Set rs = Nothing
Set db = Nothing

Screen.MousePointer = vbDefault

MsgBox "Execution time: " & Timer - lTimer & " sec." & vbCr & "of " & MSFlexGrid1.Rows - 1 & " record"

End Sub

Private Sub Command2_Click()

Dim db As New ADODB.Connection
Dim rcs As New ADODB.Recordset
Dim lTimer As Long

Screen.MousePointer = vbHourglass
Command3_Click
MSFlexGrid1.Refresh
lTimer = Timer

MSFlexGrid1.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Test.mdb;Persist Security Info=False"
rcs.Open "SELECT * FROM COMUNI", db, adOpenStatic, adLockReadOnly
rcs.MoveFirst

MSFlexGrid1.Rows = 0
MSFlexGrid1.Cols = rcs.Fields.Count - 1
Do Until rcs.EOF
    MSFlexGrid1.AddItem rcs(0) & vbTab & rcs(1) & vbTab & rcs(2) & vbTab & rcs(3) & vbTab & rcs(4) & vbTab & rcs(5) & vbTab & rcs(6) & vbTab & rcs(7)
    rcs.MoveNext
Loop
MSFlexGrid1.Visible = True

Set rcs = Nothing
Set db = Nothing

Screen.MousePointer = vbDefault

MsgBox "Execution time: " & Timer - lTimer & " sec." & vbCr & "of " & MSFlexGrid1.Rows - 1 & " record"

End Sub

Private Sub Command3_Click()

    MSFlexGrid1.Rows = 0

End Sub

Private Sub Form_Load()

    Command3_Click
    
End Sub
