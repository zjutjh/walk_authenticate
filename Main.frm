VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "毅行报名认证"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8325
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.StatusBar SbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   5520
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6006
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6006
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameShow 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   8295
      Begin VB.Label LabelTeamCode 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label LabelInfo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   48
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "操作"
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton CmdDbCon 
         Caption         =   "连接数据库"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcon 
         Caption         =   "连接机具"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox CombUsb 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CmdSS 
         Caption         =   "开始读卡"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息"
      Height          =   2655
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox Check_down 
         Caption         =   "终点模式"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TextNameDB 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CommandManId 
         Caption         =   "手工输入身份证号"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox TextTeamNum 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TextName 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.ListBox ListInfo 
         Height          =   600
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label LabelLeader 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label LabelName 
         Caption         =   "姓名"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Image PictureIdcard 
         Height          =   2295
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tagOpen As Boolean
Dim CountTest As Integer
Dim CountStart As Integer
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command


  
Private Sub Cmdcon_Click()
  If CVR_InitComm(CombUsb.ListIndex + 1001) = 1 Then
      SbMain.Panels(1).Text = "连接成功"
      CmdSS.Enabled = True
      Else
       SbMain.Panels(1).Text = "连接失败，请重新连接或者更换USB端口"
      End If
      
   
End Sub
Private Sub DbCon()
Dim dir As String
On Error GoTo errx

 dir = App.Path + "/sqldata.txt"
      Dim SQLServer As String
      Dim SQLSid As String
    
      Dim SQLPort As String
      Dim Username As String
      Dim Password As String
      Dim DBname As String
       Open dir For Input As #1
        ' 循环至文件尾.
                Line Input #1, SQLSid
              Line Input #1, SQLServer ' 读入一行数据并将其赋予某变量.
                
                Line Input #1, SQLPort
                Line Input #1, Username
                Line Input #1, Password
               Line Input #1, DBname
       Close #1
       Set conn = New ADODB.Connection
         
       conn.ConnectionString = "DRIVER={MySQL ODBC " & SQLSid & " Driver};" & _
         "SERVER=" & SQLServer & ";PORT=" & SQLPort & _
         ";DATABASE=" & DBname & ";" & _
         "UID=" & Username & ";PWD=" & Password & ";"
         
     conn.Open
    SbMain.Panels(2).Text = "数据库位置:" & SQLServer & ":" & SQLPort
   
     rs.CursorLocation = adUseClient
     Exit Sub
errx:
 
 SbMain.Panels(1).Text = "数据库连接失败" + Err.Description
 

 


  

End Sub

Private Sub CmdDbCon_Click()

 Call DbCon
 If conn.State = 1 Then
     SbMain.Panels(1).Text = "数据库连接成功"
     Me.CmdDbCon.Enabled = False
     Me.CommandManId.Enabled = True
 End If
 
End Sub

Private Sub CmdSS_Click()
If tagOpen = False Then
    CmdSS.Caption = "停止读卡"
    Me.Timer1.Enabled = True
    tagOpen = True
    Else
    CmdSS.Caption = "开始读卡"
    Me.Timer1.Enabled = False
        tagOpen = False
End If
End Sub


Private Sub CommandInputOk_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CommandManId_Click()
  Dim IdcardString As String
  IdcardString = "-1"
  IdcardString = InputBox("请输入这位毅友的身份证号", "请输入")
  If IdcardString <> "" Then Call CompareDb(0, IdcardString)
End Sub

Private Sub Form_Load()
   Me.CombUsb.Clear
   Dim i As Integer
   
   For i = 1 To 16
       Me.CombUsb.AddItem "Usb端口:" & i
   
   Next i
   CountTest = 0
   tagOpen = False
   Dim dir As String
   Me.CombUsb.ListIndex = 0
   CountStart = 0
  dir = App.Path + "/count.txt"
  Dim CountStartS As String
     Open dir For Input As #1
        ' 循环至文件尾.
              Line Input #1, CountStartS ' 读入一行数据并将其赋予某变量.
         
      
       Close #1
     
   CountStart = Val(CountStartS)
   FormShow.Show
End Sub


Private Sub CompareDb(ByVal Tag As Integer, Optional ByVal IdcardNum As String = "-1")
    On Error GoTo DBerr
   
  If (rs.State And 1) = 1 Then rs.Close
  

    Dim Mode_tag As Integer
    If Check_down.Value = 1 Then
      Mode_tag = 1
    Else
      Mode_tag = 0
    End If
    
    Dim TempName As String
    Dim TempIDCode As String
    Dim Started As Boolean
If Tag = 1 Then
     Dim StringA As String
     Dim dir As String

     Dim sinfo(0 To 7) As String
     Dim i As Integer
     i = 0
     dir = App.Path + "/wz.txt"
     Open dir For Input As #1
        Do While Not EOF(1) ' 循环至文件尾.
              Line Input #1, sinfo(i) ' 读入一行数据并将其赋予某变量.
              i = i + 1
        Loop
      Close #1

    TempName = sinfo(0)
    TempIDCode = MD5(sinfo(5))
  Else
  If Tag = 0 Then TempIDCode = MD5(IdcardNum)
  End If
  
    
    Started = False
     
     Dim Success As Boolean
     Success = False

     
     ' rs.Open sql, conn, adOpenKeyset, adLockPessimistic
     
        
      rs.Open "Select * From user_info where idcard='" & TempIDCode & "'", conn, adOpenDynamic, adLockOptimistic
     ' rs.Open , conn, adOpenDynamic, adLockOptimistic
      
Dim Name As String


   

  
     Dim iscome As Integer
     If rs.RecordCount > 0 Then
        iscome = rs.Fields("iscome").Value
        Name = rs.Fields("name").Value
        Me.TextNameDB.Text = Name
  
   
     End If
     If rs.RecordCount <= 0 Or iscome = Mode_tag + 1 Then Success = False
     If iscome = (Mode_tag + 1) Then Started = True
     
 


 If Tag = 1 Then
     
      If rs.RecordCount > 0 Then
       If iscome = Mode_tag Then Success = True
       End If
     Else
     If rs.RecordCount > 0 Then
     If MsgBox("这位毅友的姓名是【" & Name & "】请人工比对!", vbYesNo, "请比对") = vbYes And iscome = Mode_tag Then
        Success = True
     Else
        If iscome = (Mode_tag + 1) Then Started = True
        
     End If
    End If
    
     End If
     If rs.Fields("gid").Value <= 0 Then Success = False
     
     
     If Success = True Then
     Me.LabelInfo.Caption = "认证成功"
     FormShow.LabelState.Caption = "认证成功"
     'If (Tag <> 0) Then SavePicture Me.PictureIdcard.Picture, App.Path & "\tempic\" & Name & ".bmp"
     Me.LabelInfo.ForeColor = RGB(0, 255, 0)
     FormShow.LabelState.ForeColor = RGB(0, 255, 0)
     Dim Gid As Integer
     Gid = rs.Fields("gid").Value
     Me.LabelTeamCode.Caption = "组号:" & Gid
     FormShow.LabelTeamCode.Caption = "组号:" & Gid
     Dim Leadera As String
      Dim Uida As String
      Leadera = rs.Fields("gleader").Value
      Uida = rs.Fields("uid").Value
      If Uida = Leadera Then
      Me.LabelLeader.Caption = "队长"
      Me.LabelLeader.ForeColor = RGB(0, 122, 204)
      FormShow.LabelLeader.Caption = "队长"
      FormShow.LabelMsg.Caption = "您是这个队伍的队长，快去领取打卡单吧~"
       FormShow.LabelLeader.ForeColor = RGB(0, 122, 204)
      Else
      Me.LabelLeader.Caption = "队员"
      Me.LabelLeader.ForeColor = RGB(0, 0, 0)
      FormShow.LabelLeader.Caption = "队员"
       FormShow.LabelMsg.Caption = "快去找你的队长吧~"
      FormShow.LabelLeader.ForeColor = RGB(0, 0, 0)
      End If
      
     rs.Fields("iscome").Value = CStr(CInt(rs.Fields("iscome").Value) + 1)
     
     rs.Update
     rs.Close
     
     rs.Open "SELECT * FROM user_info WHERE gid='" & Gid & "'", conn, adOpenDynamic, adLockOptimistic
   
     Dim AllPep As Integer
     Dim ComePep As Integer
   AllPep = rs.RecordCount
     
   rs.MoveFirst
   Dim LeaderCame As String
   For i = 0 To AllPep - 1
      iscome = rs.Fields("iscome").Value
         If iscome <> 0 Then
             ComePep = ComePep + 1
             Dim Leader As String
             Dim Uid As String
             
             Leader = rs.Fields("gleader").Value
             Uid = rs.Fields("uid").Value
            
             LeaderCame = "未到"
             If Leader = Uid Then
               LeaderCame = "已到"
             End If
         
         End If
         If rs.EOF <> True Then rs.MoveNext
   Next i
   Me.TextTeamNum = "已认证人数" & ComePep & "/" & AllPep
   FormShow.LabelTeamCome.Caption = "您的队伍总共:" & AllPep & "人" & vbCrLf & "您的队伍已到:" & ComePep & "人" & vbCrLf & "队长:" & LeaderCame
     CountStart = CountStart + 1
     Else
     Me.LabelInfo.Caption = "认证失败"
     Me.LabelInfo.ForeColor = RGB(255, 0, 0)
          FormShow.LabelState.Caption = "认证失败"
     FormShow.LabelState.ForeColor = RGB(255, 0, 0)
      FormShow.LabelMsg.Caption = "唔，您似乎未报名，如有疑问，请联系应急工作人员~"
     If Started = True Then
     Me.LabelTeamCode = "您已认证成功，无需重复认证"
     FormShow.LabelState.Caption = "已认证"
      FormShow.LabelMsg.Caption = "您早已认证成功~别调皮啦"
     FormShow.LabelTeamCode.Caption = "组号:" & rs.Fields("gid").Value
     FormShow.LabelState.ForeColor = RGB(255, 255, 0)
     End If
     
     End If
     SbMain.Panels(3).Text = "上山计数:" & CountStart
      dir = App.Path + "/count.txt"
      Open dir For Output As #1
              Print #1, CountStart ' 读入一行数据并将其赋予某变量.

      Close #1
      rs.Close
      Exit Sub
DBerr:
 If Err.Number = -2147467259 Then
   MsgBox ("数据库连接断开，请重新连接")
   Me.CmdDbCon.Enabled = True
   Me.CommandManId.Enabled = False
 End If
 SbMain.Panels(1).Text = "错误!" + Err.Description
      
End Sub
Private Sub TestIdcard()
   If CVR_Read_Content(4) = 1 Then
     ListInfo.AddItem "读到信息"
         Dim dic As String
         Me.SbMain.Panels(1).Text = "读卡成功"
     dic = App.Path + "\zp.bmp"
      Me.PictureIdcard.Picture = LoadPicture(dic)
        
      FormShow.ImageIdpic.Picture = Me.PictureIdcard.Picture
        
      Dim strTempName As String
       Dim nReturnLen As Integer
        Dim nReturn As Integer
      
         strTempName = Space(255)
         nReturn = GetPeopleName(strTempName, nReturnLen)
        TextName.Text = strTempName
        Dim Dics As String
        Dics = App.Path & "\pic\" + TextName.Text + ".bmp"
        SavePicture Me.PictureIdcard.Picture, Dics
        FormShow.LabelName.Caption = strTempName
        '====比对数据库
        ListInfo.AddItem "比对数据库..."
        Call CompareDb(1)
        
        ListInfo.AddItem "比对完成"
        
     
      Me.Timer1.Enabled = True
      Else
          Me.SbMain.Panels(1).Text = "读卡异常，请重试"
          FormShow.LabelState.Caption = "认证异常"
          FormShow.LabelState.ForeColor = RGB(0, 122, 204)
          FormShow.LabelMsg.Caption = "卡中信息未读出，请重新放置"
      Me.Timer1.Enabled = True
     End If

End Sub
Private Sub allClear()
    Me.TextName = ""
    Me.PictureIdcard.Picture = LoadPicture()
    ListInfo.Clear
    FormShow.LabelName.Caption = ""
        FormShow.LabelTeamCode.Caption = ""
        FormShow.LabelTeamCome.Caption = ""
        FormShow.LabelState.Caption = ""
        FormShow.ImageIdpic.Picture = LoadPicture()
        Me.LabelTeamCode.Caption = ""
        Me.LabelInfo.Caption = ""
        Me.LabelLeader.Caption = ""
        FormShow.LabelLeader.Caption = ""
        FormShow.LabelMsg.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload FormShow
End Sub

Private Sub Timer1_Timer()
    If CVR_Authenticate() = 1 Then
 
         Me.SbMain.Panels(1).Text = "读卡中...."
        CountTest = 0
          Me.Timer1.Enabled = False
          Call allClear
        ListInfo.AddItem "读到身份证"
        Call TestIdcard
          Beep 3000, 400
    Else
        Me.SbMain.Panels(1).Text = "等待读卡..."
        CountTest = CountTest + 1
        Me.SbMain.Panels(2).Text = "检测:" & CountTest
        
    End If
End Sub
