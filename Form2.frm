VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10215
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   9945
   StartUpPosition =   3  '����ȱʡ
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11520
      Top             =   8760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
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
      Connect         =   "DSN=Զ�̿���ϵͳ"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Զ�̿���ϵͳ"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   "cxy14011401"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   10320
      TabIndex        =   18
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton easycmd 
      Caption         =   "ȫ���ر�"
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
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   7680
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   12720
      TabIndex        =   7
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ر�"
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
      Caption         =   "�ر�"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ر�"
      Height          =   255
      Left            =   7080
      TabIndex        =   0
      Top             =   7200
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1080
      Top             =   3120
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   2880
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   5160
      Top             =   7680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   12480
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   10920
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   10080
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   6840
      Top             =   7200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   5760
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
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
      Picture         =   "Form2.frx":0000
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
  
If Command1.Caption = "�ر�" Then
    Command1.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece1(0) = &H1
    Shape1(0).FillColor = &HFF&
Else
    Command1.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece1(0) = &H0
    Shape1(0).FillColor = &HFF00&
End If
  
Call finalcmd_Click
  
End Sub

Private Sub Command2_Click()
State = "select * from furniture where id = 2"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command2.Caption = "�ر�" Then
    Command2.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece2(0) = &H2
    Shape1(1).FillColor = &HFF&
Else
    Command2.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece2(0) = &H0
    Shape1(1).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command9_Click()
State = "select * from furniture where id = 3"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command9.Caption = "�ر�" Then
    Command9.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece3(0) = &H4
    Shape1(2).FillColor = &HFF&
Else
    Command9.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece1(0) = &H0
    Shape1(2).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command4_Click()
State = "select * from furniture where id = 4"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command4.Caption = "�ر�" Then
    Command4.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece4(0) = &H8
    Shape1(3).FillColor = &HFF&
Else
    Command4.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece4(0) = &H0
    Shape1(3).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command5_Click()
State = "select * from furniture where id = 5"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command5.Caption = "�ر�" Then
    Command5.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece5(0) = &H10
    Shape1(4).FillColor = &HFF&
Else
    Command5.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece5(0) = &H0
    Shape1(4).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command6_Click()
State = "select * from furniture where id = 6"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command6.Caption = "�ر�" Then
    Command6.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece6(0) = &H20
    Shape1(5).FillColor = &HFF&
Else
    Command6.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece6(0) = &H0
    Shape1(5).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub

Private Sub Command7_Click()
State = "select * from furniture where id = 7"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command7.Caption = "�ر�" Then
    Command7.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece7(0) = &H40
    Shape1(6).FillColor = &HFF&
Else
    Command7.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece7(0) = &H0
    Shape1(6).FillColor = &HFF00&
End If
  
Call finalcmd_Click

End Sub


Private Sub Command8_Click()
State = "select * from furniture where id = 8"
Adodc1.RecordSource = State
Adodc1.Refresh
  
If Command8.Caption = "�ر�" Then
    Command8.Caption = "��"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece8(0) = &H80
    Shape1(7).FillColor = &HFF&
Else
    Command8.Caption = "�ر�"
    Adodc1.Recordset.Fields("status") = "��"
    Adodc1.Recordset.Update
    piece8(0) = &H0
    Shape1(7).FillColor = &HFF00&
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
State = "select * from furniture where ID = 1"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(0).FillColor = &HC000&
     Command1.Caption = "�ر�"
     
     piece1(0) = &H0
     
  Else
     Shape1(0).FillColor = &HFF&
     Command1.Caption = "��"
     
     piece1(0) = &H1
         
  End If
  
  State = "select * from furniture where ID = 2"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(1).FillColor = &HC000&
     Command2.Caption = "�ر�"
     
     piece2(0) = &H0
  Else
     Shape1(1).FillColor = &HFF&
     Command2.Caption = "��"
     
     piece2(0) = &H2
  End If
  
  State = "select * from furniture where ID = 3"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(2).FillColor = &HC000&
     Command9.Caption = "�ر�"
     piece3(0) = &H0
  Else
     Shape1(2).FillColor = &HFF&
     Command9.Caption = "��"
     piece3(0) = &H4
  End If
  
  State = "select * from furniture where ID = 4"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(3).FillColor = &HC000&
     Command4.Caption = "�ر�"
     piece4(0) = &H0
  Else
     Shape1(3).FillColor = &HFF&
     Command4.Caption = "��"
     piece4(0) = &H8
  End If
  
  State = "select * from furniture where ID = 5"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(4).FillColor = &HC000&
     Command5.Caption = "�ر�"
     piece5(0) = &H0
  Else
     Shape1(4).FillColor = &HFF&
     Command5.Caption = "��"
     piece5(0) = &H10
  End If
  
  State = "select * from furniture where ID = 6"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(5).FillColor = &HC000&
     Command6.Caption = "�ر�"
     piece6(0) = &H0
  Else
     Shape1(5).FillColor = &HFF&
     Command6.Caption = "��"
     piece6(0) = &H20
  End If
  
  State = "select * from furniture where ID = 7"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(6).FillColor = &HC000&
     Command7.Caption = "�ر�"
     piece7(0) = &H0
  Else
     Shape1(6).FillColor = &HFF&
     Command7.Caption = "��"
     piece7(0) = &H40
  End If
  
  State = "select * from furniture where ID = 8"
  Adodc1.RecordSource = State
  Adodc1.Refresh
  
  If Trim(Adodc1.Recordset.Fields("status")) = "��" Then
     Shape1(7).FillColor = &HC000&
     Command8.Caption = "�ر�"
     piece8(0) = &H0
  Else
     Shape1(7).FillColor = &HFF&
     Command8.Caption = "��"
     piece8(0) = &H80
  End If

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

If easycmd.Caption = "ȫ����" Then
    easycmd.Caption = "ȫ���ر�"

    piece(0) = &H0
    MSComm1.Output = piece

    tmpopen = "select * from furniture where status = '��'"
    Adodc1.RecordSource = tmpopen
    Adodc1.Refresh
 
    If Adodc1.Recordset.BOF = False Then
        tmpcount = Adodc1.Recordset.RecordCount
    End If
 
    For i = 1 To tmpcount
        Adodc1.Recordset.Fields("status") = "��"
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
    Next i
    Command1.Caption = "�ر�"
    Command2.Caption = "�ر�"
    Command9.Caption = "�ر�"
    Command4.Caption = "�ر�"
    Command5.Caption = "�ر�"
    Command6.Caption = "�ر�"
    Command7.Caption = "�ر�"
    Command8.Caption = "�ر�"
    For i = 0 To 7
        Shape1(i).FillColor = &HFF00&
    Next i
Else
    easycmd.Caption = "ȫ����"
    piece(0) = &HFF
    MSComm1.Output = piece
 
    tmpclose = "select * from furniture where status = '��'"
    Adodc1.RecordSource = tmpclose
    Adodc1.Refresh
 
    If Adodc1.Recordset.BOF = False Then
        tmpcount = Adodc1.Recordset.RecordCount
    End If
 
    For i = 1 To tmpcount
        Adodc1.Recordset.Fields("status") = "��"
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
    Next i
    Command1.Caption = "��"
    Command2.Caption = "��"
    Command9.Caption = "��"
    Command4.Caption = "��"
    Command5.Caption = "��"
    Command6.Caption = "��"
    Command7.Caption = "��"
    Command8.Caption = "��"
    For i = 0 To 7
        Shape1(i).FillColor = &HFF&
    Next i
 
 End If
 
End Sub
