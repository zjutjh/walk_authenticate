VERSION 5.00
Begin VB.Form FormShow 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÏûÏ¢"
   ClientHeight    =   11145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   22305
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   22305
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame FrameShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   16935
      Begin VB.Label LabelMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   720
         TabIndex        =   6
         Top             =   5400
         Width           =   105
      End
      Begin VB.Image ImageIdpic 
         Height          =   2055
         Left            =   480
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label LabelName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Width           =   285
      End
      Begin VB.Label LabelState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   72
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   8400
         TabIndex        =   4
         Top             =   720
         Width           =   420
      End
      Begin VB.Label LabelTeamCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   720
         TabIndex        =   3
         Top             =   3360
         Width           =   285
      End
      Begin VB.Label LabelTeamCome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9720
         TabIndex        =   2
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label LabelLeader 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FormShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Form_Resize()
Me.FrameShow.Left = Me.ScaleWidth / 2 - Me.FrameShow.Width / 2
Me.FrameShow.Top = Me.ScaleHeight / 2 - Me.FrameShow.Height / 2
End Sub

