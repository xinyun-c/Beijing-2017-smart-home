VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   12855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command9 
      Caption         =   "关闭"
      Height          =   255
      Left            =   10320
      TabIndex        =   18
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton easycmd 
      Caption         =   "全部关闭"
      Height          =   615
      Left            =   960
      TabIndex        =   17
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton finalcmd 
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "关闭"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "关闭"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   7680
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "关闭"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "关闭"
      Height          =   255
      Left            =   12720
      TabIndex        =   7
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关闭"
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   255
      Left            =   7080
      TabIndex        =   0
      Top             =   7200
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   3120
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10920
      Top             =   8880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=远程控制系统"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "远程控制系统"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   "cxy14011401"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2880
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape7 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5160
      Top             =   7680
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   12480
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10920
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape9 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10080
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6840
      Top             =   7200
      Width           =   255
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5760
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   6600
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "irrigation"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "air purifier"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "shower"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   " baker"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "lamp2"
      Height          =   255
      Left            =   12720
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "lamp1"
      Height          =   255
      Left            =   11160
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "AC"
      Height          =   255
      Left            =   10440
      TabIndex        =   4
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "  TV"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   6840
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As String
Dim comnum As Integer
Dim piece(0) As Byte
Dim piece1(0) As Byte
Dim piece2(0) As Byte
Dim piece3(0) As Byte
Dim piece4(0) As Byte
Dim piece5(0) As Byte
Dim piece6(0) As Byte
Dim piece7(0) As Byte
Dim piece8(0) As Byte

Private Sub Command1_Click()
State = "select * from furniture where id = 1"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command1.Caption = "关闭" Then
    Command1.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece1(0) = &H1
    Shape1.FillColor = &HFF&
Else
    Command1.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece1(0) = &H0
    Shape1.FillColor = &HFF00&
End If
  
Call finalcmd_Click
  
End Sub

Private Sub Command2_Click()
State = "select * from furniture where id = 2"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command2.Caption = "关闭" Then
    Command2.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece2(0) = &H2
    Shape2.FillColor = &HFF&
Else
    Command2.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece2(0) = &H0
    Shape2.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command9_Click()
State = "select * from furniture where id = 3"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command9.Caption = "关闭" Then
    Command9.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece3(0) = &H4
    Shape9.FillColor = &HFF&
Else
    Command9.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece1(0) = &H0
    Shape9.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command4_Click()
State = "select * from furniture where id = 4"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command4.Caption = "关闭" Then
    Command4.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece4(0) = &H8
    Shape4.FillColor = &HFF&
Else
    Command4.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece4(0) = &H0
    Shape4.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command5_Click()
State = "select * from furniture where id = 5"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command5.Caption = "关闭" Then
    Command5.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece5(0) = &H10
    Shape5.FillColor = &HFF&
Else
    Command5.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece5(0) = &H0
    Shape5.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command6_Click()
State = "select * from furniture where id = 6"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command6.Caption = "关闭" Then
    Command6.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece6(0) = &H20
    Shape6.FillColor = &HFF&
Else
    Command6.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece6(0) = &H0
    Shape6.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command7_Click()
State = "select * from furniture where id = 7"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command7.Caption = "关闭" Then
    Command7.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece7(0) = &H40
    Shape6.FillColor = &HFF&
Else
    Command7.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece7(0) = &H0
    Shape7.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub


Private Sub Command8_Click()
State = "select * from furniture where id = 8"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command8.Caption = "关闭" Then
    Command8.Caption = "打开"
    Adodc1.Recordset.Fields("status") = "关"
    Adodc1.Recordset.Update
    piece8(0) = &H80
    Shape8.FillColor = &HFF&
Else
    Command8.Caption = "关闭"
    Adodc1.Recordset.Fields("status") = "开"
    Adodc1.Recordset.Update
    piece8(0) = &H0
    Shape8.FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub form_load()
comnum = 3
MSComm1.CommPort = comnum

If MSComm1.CommPort = True Then
    MSComm1.CommPort = False
End If

With MSComm1
    Settings = "9600,N,8,1"
End With

If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
End If
  
MSComm1.InBufferCount = 0
MSComm1.OutBufferCount = 0
End Sub

Private Sub Timer1_Timer()
Select Case i
    Case (1), x = 1
    Case (2), x = 2
    Case (3), x = 4
    Case (4), x = 8
    Case (5), x = 10
    Case (6), x = 20
    Case (7), x = 40
    Case (8), x = 80
End Select

i = 1
For i = 1 To 8
    State = "select * from furniture where id = i"
    Adodc1.RecordSource = State
    Adodc1.Refresh
  
    If Trim(Adodc1.Recordset.Fields("status")) = "开" Then
        Shapei.FillColor = &HC000&
        Commandi.Caption = "关闭"
        piecei(0) = &H0
    Else
        Shapei.FillColor = &HFF&
        Commandi.Caption = "打开"
        piecei(0) = "&H" & x
    End If
Next i
Call finalcmd_Click
End Sub

Private Sub finalcmd_Click()
piece(0) = piece1(0) Or piece2(0) Or piece3(0) Or piece4(0) Or piece5(0) Or piece6(0) Or piece7(0) Or piece8(0)
MSComm1.Output = piece
End Sub

Private Sub easycmd_Click()
    Dim tmpclose As String
    Dim tmpopen As String
    Dim tmpcount As Integer
    Dim i As Integer
  
    
If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
End If
  
MSComm1.CommPort = comnum
MSComm1.Settings = "9600,n,8,1"
MSComm1.Handshaking = comNone

If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
End If
  
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0

If easycmd.Caption = "全部打开" Then
    easycmd.Caption = "全部关闭"

    piece(0) = &H0
    MSComm1.Output = piece

    tmpopen = "select * from furniture where status = '关'"
    Adodc1.RecordSource = tmpopen
    Adodc1.Refresh
 
    If Adodc1.Recordset.BOF = False Then
        tmpcount = Adodc1.Recordset.RecordCount
    End If
 
    For i = 1 To tmpcount
        Adodc1.Recordset.Fields("status") = "开"
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
    Next i
    Command1.Caption = "关闭"
    Command2.Caption = "关闭"
    Command9.Caption = "关闭"
    Command4.Caption = "关闭"
    Command5.Caption = "关闭"
    Command6.Caption = "关闭"
    Command7.Caption = "关闭"
    Command8.Caption = "关闭"
    Shape1.FillColor = &HFF00&
    Shape2.FillColor = &HFF00&
    Shape9.FillColor = &HFF00&
    Shape4.FillColor = &HFF00&
    Shape5.FillColor = &HFF00&
    Shape6.FillColor = &HFF00&
    Shape7.FillColor = &HFF00&
    Shape8.FillColor = &HFF00&
Else
    easycmd.Caption = "全部打开"
    piece(0) = &HFF
    MSComm1.Output = piece
 
    tmpclose = "select * from furniture where status = '开'"
    Adodc1.RecordSource = tmpclose
    Adodc1.Refresh
 
    If Adodc1.Recordset.BOF = False Then
        tmpcount = Adodc1.Recordset.RecordCount
    End If
 
    For i = 1 To tmpcount
        Adodc1.Recordset.Fields("status") = "关"
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
    Next i
    Command1.Caption = "打开"
    Command2.Caption = "打开"
    Command9.Caption = "打开"
    Command4.Caption = "打开"
    Command5.Caption = "打开"
    Command6.Caption = "打开"
    Command7.Caption = "打开"
    Command8.Caption = "打开"
    Shape1.FillColor = &HFF&
    Shape2.FillColor = &HFF&
    Shape9.FillColor = &HFF&
    Shape4.FillColor = &HFF&
    Shape5.FillColor = &HFF&
    Shape6.FillColor = &HFF&
    Shape7.FillColor = &HFF&
    Shape8.FillColor = &HFF&
 
 End If
 
End Sub
