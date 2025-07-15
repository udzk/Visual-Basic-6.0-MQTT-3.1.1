VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MQTT Client"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   20880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   20880
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDCDCD&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   14040
      TabIndex        =   35
      Top             =   0
      Width           =   6855
      Begin VB.TextBox log 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   600
         Width           =   6735
      End
      Begin VB.TextBox txtState 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0059911C&
         Height          =   285
         Left            =   6360
         TabIndex        =   36
         Text            =   "0"
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7215
         Left            =   6600
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Height          =   7215
         Left            =   0
         TabIndex        =   39
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Labellog 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Log"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C1710F&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame title 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   12600
      TabIndex        =   41
      Top             =   0
      Width           =   1455
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   960
         X2              =   1200
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   960
         X2              =   1200
         Y1              =   480
         Y2              =   240
      End
      Begin VB.Label exit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005A49F1&
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   42
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6375
      Left            =   7080
      TabIndex        =   26
      Top             =   1200
      Width           =   6855
      Begin VB.TextBox txtTopicGET 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C1710F&
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Text            =   "AAA"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtRx 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label cmdGET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Subscribe"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   1440
         TabIndex        =   29
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label border_cmdGET 
         BackColor       =   &H00C1710F&
         Height          =   495
         Left            =   5040
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   6855
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtTopic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C1710F&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Text            =   "AAA"
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   1440
         TabIndex        =   25
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label cmdPUT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label border_cmdPUT 
         BackColor       =   &H00C1710F&
         Height          =   495
         Left            =   5520
         TabIndex        =   24
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0059911C&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Text            =   "60"
         Top             =   2040
         Width           =   5055
      End
      Begin VB.TextBox txtIdentifier 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0059911C&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Text            =   "mqttid0000"
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox Port 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0059911C&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Text            =   "1883"
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox cmbBroker 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0059911C&
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label cmdDisconnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label cmdConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   5295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Keep Alive"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Host"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label border_cmdConnect 
         BackColor       =   &H0059911C&
         Height          =   495
         Left            =   5280
         TabIndex        =   14
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label border_cmdDisconnect 
         BackColor       =   &H00808080&
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDCDCD&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   11520
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   50000
      Left            =   12000
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   1883
   End
   Begin VB.Label logbtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<Log>"
      BeginProperty Font 
         Name            =   "System"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1710F&
      Height          =   375
      Left            =   13200
      TabIndex        =   55
      Top             =   720
      Width           =   735
   End
   Begin VB.Label cmdPing 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<Ping>"
      BeginProperty Font 
         Name            =   "System"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0059911C&
      Height          =   375
      Left            =   12360
      TabIndex        =   54
      Top             =   720
      Width           =   855
   End
   Begin VB.Label titletxtmqtt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MQTT 3.1.1 Client"
      BeginProperty Font 
         Name            =   "System"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1200
      TabIndex        =   53
      Top             =   720
      Width           =   2040
   End
   Begin VB.Label txtMQTTState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "System"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   3360
      TabIndex        =   52
      Top             =   720
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmMain.frx":0452
      Top             =   120
      Width           =   960
   End
   Begin VB.Label titletxt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0059911C&
      Height          =   375
      Left            =   1200
      TabIndex        =   43
      Top             =   240
      Width           =   2340
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' udk
' Linux_index@outlook.com

' 顶部声明
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
Private Const MAX_PACKET_SIZE As Long = 1048576 ' 1MB 最大分片大小

' 声明常量
Private Const MQTT_CONNECT As Integer = &H10
Private Const MQTT_CONNACK As Integer = &H20
Private Const MQTT_PUBLISH As Integer = &H30
Private Const MQTT_SUBSCRIBE As Integer = &H82
Private Const MQTT_SUBACK As Integer = &H90
Private Const MQTT_UNSUBSCRIBE As Integer = &HA2
Private Const MQTT_PINGREQ As Integer = &HC0
Private Const MQTT_PINGRESP As Integer = &HD0
Private Const MQTT_DISCONNECT As Integer = &HE0
Private Const SOCKET_CONNECTED As Integer = 7

' 全局变量
Dim m_Connected As Boolean
Dim partialBuffer() As Byte   ' 用于分片消息
Dim partialSize As Long       ' 当前分片大小
Dim partialStart As Long      ' 分片起始位置

' 剩余长度编码函数
Private Function EncodeRemainingLength(ByVal length As Long) As Byte()
    Dim result() As Byte
    ReDim result(0 To 3) As Byte ' 最大4字节
    Dim index As Long: index = 0
    Dim temp As Long: temp = length
    
    Do
        Dim digit As Byte
        digit = temp Mod 128
        temp = temp \ 128
        If temp > 0 Then
            digit = digit Or &H80 ' 设置最高位表示还有后续字节
        End If
        result(index) = digit
        index = index + 1
    Loop While temp > 0 And index < 4
    
    ReDim Preserve result(0 To index - 1)
    EncodeRemainingLength = result
End Function

' 使用Windows API的UTF-8转换函数
Private Function UTF8ToStr(b() As Byte) As String
    ' 检查空数组
    If UBound(b) < LBound(b) Then
        UTF8ToStr = ""
        Exit Function
    End If
    
    Dim length As Long
    Dim result As String
    
    ' 获取所需缓冲区长度
    length = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), UBound(b) - LBound(b) + 1, 0, 0)
    
    If length > 0 Then
        result = String$(length, 0)
        Dim ret As Long
        ret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), UBound(b) - LBound(b) + 1, StrPtr(result), length)
        
        If ret > 0 Then
            UTF8ToStr = result
        Else
            ' 尝试使用ADODB作为后备方案
            On Error Resume Next
            With CreateObject("ADODB.Stream")
                .Type = 1 ' adTypeBinary
                .Open
                .Write b
                .Position = 0
                .Type = 2 ' adTypeText
                .Charset = "utf-8"
                UTF8ToStr = .ReadText
                .Close
            End With
            If Err.Number <> 0 Then
                UTF8ToStr = "[Invalid UTF-8]"
            End If
            On Error GoTo 0
        End If
    Else
        UTF8ToStr = ""
    End If
End Function

Private Function StringToUTF8(ByVal str As String) As Byte()
    With CreateObject("ADODB.Stream")
        .Type = 2  ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText str
        .Position = 0
        .Type = 1  ' adTypeBinary
        .Position = 3  ' 跳过BOM
        StringToUTF8 = .Read
        .Close
    End With
End Function

' 字节数组十六进制显示函数
Private Function HexDumpBytes(data() As Byte) As String
    Dim i As Long
    Dim hexStr As String
    Dim asciiStr As String
    Dim result As String
    
    For i = LBound(data) To UBound(data)
        hexStr = hexStr & Right("0" & Hex(data(i)), 2) & " "
        
        If data(i) >= 32 And data(i) <= 126 Then
            asciiStr = asciiStr & Chr(data(i))
        Else
            asciiStr = asciiStr & "."
        End If
        
        If (i + 1) Mod 16 = 0 Then
            result = result & hexStr & "  " & asciiStr & vbCrLf
            hexStr = ""
            asciiStr = ""
        End If
    Next i
    
    If hexStr <> "" Then
        While Len(hexStr) < 48
            hexStr = hexStr & "   "
        Wend
        result = result & hexStr & "  " & asciiStr
    End If
    
    HexDumpBytes = result
End Function

Private Sub Form_Load()
    logbtn_Click '开启日志
    txtIdentifier = "mqttid" & Format(Time, "mm") & Trim(Format(Time, "ss"))
    m_Connected = False
    partialSize = 0
    partialStart = 0
    log.Text = "Application started" & vbCrLf
    DoEvents
End Sub

' 关闭连接
Private Sub cmdClose()
    On Error Resume Next
    log.Text = log.Text & "Closing connection..." & vbCrLf
    Winsock1.Close
    m_Connected = False
    While Winsock1.State <> 0
        DoEvents
    Wend
    log.Text = log.Text & "Connection closed" & vbCrLf
End Sub

' 打开连接
Private Sub cmdOpen()
    On Error Resume Next
    log.Text = log.Text & "Opening TCP connection to " & cmbBroker.Text & vbCrLf
    cmdClose
    Winsock1.LocalPort = 0
    Winsock1.RemotePort = Port.Text
    Winsock1.RemoteHost = cmbBroker.Text
    Winsock1.Connect
End Sub

' 断开MQTT连接
Private Sub cmdDisconnect_Click()
    If Winsock1.State = SOCKET_CONNECTED And m_Connected Then
        ' 构建DISCONNECT报文
        Dim discPacket() As Byte
        Dim lenBytes() As Byte
        
        ' 剩余长度 = 0
        lenBytes = EncodeRemainingLength(0)
        Dim lenBytesCount As Long: lenBytesCount = UBound(lenBytes) - LBound(lenBytes) + 1
        
        ReDim discPacket(0 To lenBytesCount) ' 类型(1字节) + 剩余长度字节
        discPacket(0) = MQTT_DISCONNECT
        
        Dim i As Long
        For i = 0 To lenBytesCount - 1
            discPacket(i + 1) = lenBytes(i)
        Next i
        
        log.Text = log.Text & "Sending MQTT DISCONNECT" & vbCrLf
        Winsock1.SendData discPacket
    End If
    
    Winsock1.Close
    m_Connected = False
End Sub

' 建立MQTT连接
Private Sub cmdConnect_Click()
    Dim Identifier As String
    Dim i As Long
    
    Identifier = txtIdentifier
    
    ' 将客户端ID转换为UTF-8字节数组
    Dim idBytes() As Byte
    idBytes = StringToUTF8(Identifier)
    Dim idByteLen As Long: idByteLen = UBound(idBytes) - LBound(idBytes) + 1
    
    ' 计算剩余长度（不包括固定头）
    Dim MsgLen As Long
    MsgLen = 10 + 2 + idByteLen  ' 协议名(6) + 协议级别(1) + 标志(1) + KeepAlive(2) + 客户端ID长度(2) + ID
    
    ' 保存原始MsgLen值
    Dim originalMsgLen As Long: originalMsgLen = MsgLen
    
    ' 编码剩余长度
    Dim lenBytes() As Byte
    lenBytes = EncodeRemainingLength(MsgLen)
    Dim lenBytesCount As Long: lenBytesCount = UBound(lenBytes) - LBound(lenBytes) + 1
    
    ' 计算报文总大小
    Dim packetSize As Long: packetSize = 1 + lenBytesCount + originalMsgLen
    Dim connectPacket() As Byte
    ReDim connectPacket(0 To packetSize - 1)
    
    ' 固定头
    connectPacket(0) = MQTT_CONNECT
    
    ' 复制剩余长度字节
    For i = 0 To lenBytesCount - 1
        connectPacket(1 + i) = lenBytes(i)
    Next i
    
    ' 可变头部和负载
    Dim pos As Long: pos = 1 + lenBytesCount
    
    ' 协议名
    connectPacket(pos) = 0: pos = pos + 1
    connectPacket(pos) = 4: pos = pos + 1
    connectPacket(pos) = Asc("M"): pos = pos + 1
    connectPacket(pos) = Asc("Q"): pos = pos + 1
    connectPacket(pos) = Asc("T"): pos = pos + 1
    connectPacket(pos) = Asc("T"): pos = pos + 1
    
    ' 协议级别
    connectPacket(pos) = 4: pos = pos + 1
    
    ' 连接标志
    connectPacket(pos) = &H2: pos = pos + 1
    
    ' 保持时间
    connectPacket(pos) = &H0: pos = pos + 1
    connectPacket(pos) = &H3C: pos = pos + 1
    
    ' 客户端ID长度 (高字节在前)
    connectPacket(pos) = idByteLen \ 256: pos = pos + 1
    connectPacket(pos) = idByteLen Mod 256: pos = pos + 1
    
    ' 客户端ID (UTF-8字节)
    For i = 0 To idByteLen - 1
        connectPacket(pos) = idBytes(i): pos = pos + 1
    Next i
    
    ' 连接并发送
    cmdOpen
    While Winsock1.State <> SOCKET_CONNECTED
        DoEvents
    Wend
    
    Winsock1.SendData connectPacket
    log.Text = log.Text & "CONNECT sent (" & packetSize & " bytes)" & vbCrLf
End Sub

' 发送PING请求
Private Sub cmdPing_Click()
    If Winsock1.State = SOCKET_CONNECTED And m_Connected Then
        Dim pingReq() As Byte
        Dim lenBytes() As Byte
        
        ' 剩余长度 = 0
        lenBytes = EncodeRemainingLength(0)
        Dim lenBytesCount As Long: lenBytesCount = UBound(lenBytes) - LBound(lenBytes) + 1
        
        ReDim pingReq(0 To lenBytesCount) ' 类型(1字节) + 剩余长度字节
        pingReq(0) = MQTT_PINGREQ
        
        Dim i As Long
        For i = 0 To lenBytesCount - 1
            pingReq(i + 1) = lenBytes(i)
        Next i
        
        log.Text = log.Text & "Sending PINGREQ" & vbCrLf
        Winsock1.SendData pingReq
    Else
        MsgBox "Please establish MQTT connection first"
    End If
End Sub

' 发布消息
Private Sub cmdPUT_Click()
    If Winsock1.State <> SOCKET_CONNECTED Or Not m_Connected Then Exit Sub
    If Len(txtMsg) < 1 Then Exit Sub
    
    ' 转换为UTF-8
    Dim topicBytes() As Byte
    Dim msgBytes() As Byte
    topicBytes = StringToUTF8(txtTopic)
    msgBytes = StringToUTF8(txtMsg)
    
    Dim TopicLength As Long: TopicLength = UBound(topicBytes) - LBound(topicBytes) + 1
    Dim MsgLength As Long: MsgLength = UBound(msgBytes) - LBound(msgBytes) + 1
    
    ' 计算剩余长度 = 2(主题长度) + TopicLength + MsgLength
    Dim remainingLength As Long
    remainingLength = 2 + TopicLength + MsgLength
    
    ' 编码剩余长度
    Dim lenBytes() As Byte
    lenBytes = EncodeRemainingLength(remainingLength)
    Dim lenBytesCount As Long: lenBytesCount = UBound(lenBytes) - LBound(lenBytes) + 1
    
    ' 构建报文
    Dim publishPacket() As Byte
    ReDim publishPacket(0 To 1 + lenBytesCount + remainingLength - 1)
    
    ' 固定头
    publishPacket(0) = MQTT_PUBLISH
    
    ' 剩余长度
    Dim i As Long
    For i = 0 To lenBytesCount - 1
        publishPacket(1 + i) = lenBytes(i)
    Next i
    
    Dim pos As Long: pos = 1 + lenBytesCount
    
    ' 主题长度 (高字节在前)
    publishPacket(pos) = TopicLength \ 256: pos = pos + 1
    publishPacket(pos) = TopicLength Mod 256: pos = pos + 1
    
    ' 主题内容
    For i = 0 To TopicLength - 1
        publishPacket(pos) = topicBytes(i): pos = pos + 1
    Next i
    
    ' 消息内容
    For i = 0 To MsgLength - 1
        publishPacket(pos) = msgBytes(i): pos = pos + 1
    Next i
    
    Winsock1.SendData publishPacket
    log.Text = log.Text & "PUBLISH to [" & txtTopic & "]: " & txtMsg & " (" & UBound(publishPacket) + 1 & " bytes)" & vbCrLf
End Sub

' 订阅主题
Private Sub cmdGET_Click()
    Static msgID As Integer
    Dim i As Long
    
    If Not m_Connected Then
        MsgBox "Please establish MQTT connection first"
        Exit Sub
    End If
    
    ' 转换为UTF-8
    Dim topicBytes() As Byte
    topicBytes = StringToUTF8(txtTopicGET)
    Dim TopicLength As Long: TopicLength = UBound(topicBytes) - LBound(topicBytes) + 1
    
    ' 剩余长度 = 2(消息ID) + 2(主题长度) + TopicLength + 1(QoS)
    Dim remainingLength As Long
    remainingLength = 2 + 2 + TopicLength + 1
    
    ' 编码剩余长度
    Dim lenBytes() As Byte
    lenBytes = EncodeRemainingLength(remainingLength)
    Dim lenBytesCount As Long: lenBytesCount = UBound(lenBytes) - LBound(lenBytes) + 1
    
    ' 生成消息ID
    msgID = (msgID + 1) Mod 65536
    If msgID = 0 Then msgID = 1
    
    ' 构建报文
    Dim byteArr() As Byte
    ReDim byteArr(0 To 1 + lenBytesCount + remainingLength - 1)
    
    ' 固定头
    byteArr(0) = MQTT_SUBSCRIBE
    
    ' 剩余长度
    For i = 0 To lenBytesCount - 1
        byteArr(1 + i) = lenBytes(i)
    Next i
    
    Dim pos As Long: pos = 1 + lenBytesCount
    
    ' 消息ID (高字节在前)
    byteArr(pos) = msgID \ 256: pos = pos + 1
    byteArr(pos) = msgID Mod 256: pos = pos + 1
    
    ' 主题长度 (高字节在前)
    byteArr(pos) = TopicLength \ 256: pos = pos + 1
    byteArr(pos) = TopicLength Mod 256: pos = pos + 1
    
    ' 主题内容
    For i = 0 To TopicLength - 1
        byteArr(pos) = topicBytes(i): pos = pos + 1
    Next i
    
    ' QoS 级别
    byteArr(pos) = 0  ' QoS 0
    
    ' 发送
    Winsock1.SendData byteArr
    log.Text = log.Text & "SUBSCRIBE sent for [" & txtTopicGET & "] (ID:" & msgID & ", " & UBound(byteArr) + 1 & " bytes)" & vbCrLf
End Sub

' 接收数据处理 - 修复分片处理中的下标越界问题
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrorHandler
    
    Dim byteData() As Byte
    ReDim byteData(0 To bytesTotal - 1)
    Winsock1.GetData byteData, vbArray + vbByte
    
    log.Text = log.Text & "Received " & bytesTotal & "B: " & HexDumpBytes(byteData) & vbCrLf
    
    ' 如果有分片数据，合并到新数据前面
    If partialSize > 0 Then
        Dim mergedData() As Byte
        Dim mergedSize As Long
        mergedSize = partialSize + bytesTotal
        
        ' 安全边界检查
        If mergedSize > MAX_PACKET_SIZE Then
            log.Text = log.Text & "ERROR: Merged packet exceeds max size (" & MAX_PACKET_SIZE & "B)" & vbCrLf
            partialSize = 0
            Erase partialBuffer
            Exit Sub
        End If
        
        ReDim mergedData(0 To mergedSize - 1)
        
        ' 复制分片数据
        If partialSize > 0 Then
            CopyMemory mergedData(0), partialBuffer(0), partialSize
        End If
        
        ' 复制新数据
        CopyMemory mergedData(partialSize), byteData(0), bytesTotal
        
        ' 替换当前数据
        byteData = mergedData
        bytesTotal = mergedSize
        partialSize = 0
        Erase partialBuffer
    End If
    
    Dim startIndex As Long: startIndex = 0
    Do While startIndex < bytesTotal
        Dim MsgType As Byte
        MsgType = byteData(startIndex) And &HF0
        
        ' 处理PUBLISH
        If MsgType = MQTT_PUBLISH Then
            Dim pos As Long: pos = startIndex + 1
            Dim remainingLength As Long
            Dim multiplier As Long: multiplier = 1
            Dim digit As Byte
            Dim lenBytesCount As Long: lenBytesCount = 0
            
            ' 解析剩余长度
            Do
                If pos > UBound(byteData) Then
                    log.Text = log.Text & "Error: Unexpected end of data while parsing remaining length" & vbCrLf
                    Exit Do
                End If
                
                digit = byteData(pos)
                remainingLength = remainingLength + (digit And 127) * multiplier
                multiplier = multiplier * 128
                pos = pos + 1
                lenBytesCount = lenBytesCount + 1
            Loop While (digit And 128) <> 0
            
            ' 检查是否有足够的数据
            Dim expectedSize As Long
            expectedSize = startIndex + 1 + lenBytesCount + remainingLength - 1 ' 整个消息结束位置
            
            If expectedSize > UBound(byteData) Then
                ' 消息不完整，保存分片
                Dim remainingBytes As Long
                remainingBytes = bytesTotal - startIndex
                
                If remainingBytes > MAX_PACKET_SIZE Then
                    log.Text = log.Text & "ERROR: Partial packet too large (" & remainingBytes & "B)" & vbCrLf
                    Exit Do
                End If
                
                ReDim partialBuffer(0 To remainingBytes - 1)
                CopyMemory partialBuffer(0), byteData(startIndex), remainingBytes
                partialSize = remainingBytes
                log.Text = log.Text & "Partial message saved (" & partialSize & " bytes)" & vbCrLf
                Exit Do ' 退出循环，等待更多数据
            End If
            
            ' 主题长度 (高位在前)
            If pos + 1 > UBound(byteData) Then
                log.Text = log.Text & "Error: Missing topic length data" & vbCrLf
                Exit Do
            End If
            
            Dim TopicLen As Integer
            TopicLen = byteData(pos) * 256 + byteData(pos + 1)
            pos = pos + 2
            
            ' 检查主题数据是否完整
            If pos + TopicLen - 1 > UBound(byteData) Then
                log.Text = log.Text & "Error: Incomplete topic data" & vbCrLf
                Exit Do
            End If
            
            ' 提取主题(UTF-8)
            Dim topicBytes() As Byte
            ReDim topicBytes(0 To TopicLen - 1)
            CopyMemory topicBytes(0), byteData(pos), TopicLen
            pos = pos + TopicLen
            
            ' 计算有效载荷长度
            Dim payloadLen As Long
            payloadLen = remainingLength - 2 - TopicLen
            
            ' 检查有效载荷数据是否完整
            If payloadLen > 0 Then
                If pos + payloadLen - 1 > UBound(byteData) Then
                    log.Text = log.Text & "Error: Incomplete payload data" & vbCrLf
                    Exit Do
                End If
                
                ' 提取有效载荷(UTF-8)
                Dim payloadBytes() As Byte
                ReDim payloadBytes(0 To payloadLen - 1)
                CopyMemory payloadBytes(0), byteData(pos), payloadLen
                pos = pos + payloadLen
            Else
                ReDim payloadBytes(0)
            End If
            
            ' 安全转换UTF-8
            Dim topic As String
            Dim payload As String
            
            On Error Resume Next
            topic = UTF8ToStr(topicBytes)
            If Err.Number <> 0 Then
                log.Text = log.Text & "Error converting topic to UTF-8: " & Err.Description & vbCrLf
                topic = "[Invalid UTF-8]"
                Err.Clear
            End If
            
            If payloadLen > 0 Then
                On Error Resume Next
                payload = UTF8ToStr(payloadBytes)
                If Err.Number <> 0 Then
                    log.Text = log.Text & "Error converting payload to UTF-8: " & Err.Description & vbCrLf
                    payload = "[Invalid UTF-8]"
                    Err.Clear
                End If
            Else
                payload = ""
            End If
            On Error GoTo 0
            
            log.Text = log.Text & "PUBLISH received: [" & topic & "] " & payload & vbCrLf
            txtRx = topic & vbCrLf & vbCrLf & payload
            
            ' 移动到下一个消息
            startIndex = pos
        Else
            ' 处理其他类型的消息...
            ' 处理CONNACK
            If MsgType = MQTT_CONNACK Then
                If UBound(byteData) >= startIndex + 3 Then
                    Dim connCode As Integer
                    connCode = byteData(startIndex + 3)
                    If connCode = 0 Then
                        m_Connected = True
                        log.Text = log.Text & "MQTT connection established" & vbCrLf
                    Else
                        m_Connected = False
                        log.Text = log.Text & "Connection refused: code " & connCode & vbCrLf
                    End If
                End If
                startIndex = startIndex + 4 ' CONNACK固定4字节
            End If
            
            ' 处理SUBACK
            If MsgType = MQTT_SUBACK Then
                log.Text = log.Text & "SUBACK received" & vbCrLf
                If UBound(byteData) >= startIndex + 4 Then
                    Dim returnCode As Integer
                    Dim packetID As Integer
                    packetID = byteData(startIndex + 2) * 256 + byteData(startIndex + 3)
                    returnCode = byteData(startIndex + 4)
                    
                    log.Text = log.Text & "Packet ID: " & packetID & ", Return code: " & Hex(returnCode) & vbCrLf
                    
                    If returnCode < &H80 Then
                        log.Text = log.Text & "Subscribe successful (QoS " & returnCode & ")" & vbCrLf
                    Else
                        log.Text = log.Text & "Subscribe failed: " & Hex(returnCode) & vbCrLf
                    End If
                End If
                startIndex = startIndex + 5 ' SUBACK至少5字节
            End If
            
            ' 处理PING响应
            If MsgType = MQTT_PINGRESP Then
                log.Text = log.Text & "PINGRESP received" & vbCrLf
                startIndex = startIndex + 2 ' PINGRESP固定2字节
            End If
            
            ' 如果未处理，移动到下一个字节
            If startIndex = 0 Then
                startIndex = startIndex + 1
            End If
        End If
    Loop
    
    Exit Sub
    
ErrorHandler:
    log.Text = log.Text & "ERROR in DataArrival: " & Err.Description & vbCrLf
    partialSize = 0
    Erase partialBuffer
End Sub

' 定时更新状态
Private Sub Timer1_Timer()
    txtState = Winsock1.State
    If txtState = "7" Then
        txtMQTTState.Caption = "Connected"
        m_Connected = True
    Else
        txtMQTTState.Caption = "Disconnected"
        m_Connected = False
    End If
End Sub

' 定时发送PING
Private Sub Timer2_Timer()
    If m_Connected Then
        cmdPing_Click
    End If
End Sub

Private Sub log_Change()
log.SelStart = Len(log.Text)
End Sub

Private Sub logbtn_Click()
If frmMain.Width = 14045 Then
frmMain.Width = 20880
logbtn.ForeColor = &HC1710F
Else
frmMain.Width = 14045
logbtn.ForeColor = &H808080
End If
End Sub

Private Sub border_cmdConnect_Click()
cmdConnect_Click
End Sub

Private Sub border_cmdDisconnect_Click()
cmdDisconnect_Click
End Sub

Private Sub border_cmdGET_Click()
cmdGET_Click
End Sub

Private Sub border_cmdPUT_Click()
cmdPUT_Click
End Sub

Private Sub Labellog_DblClick()
log.Text = ""
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static FO(2): If Button <> 0 Then Me.Move Me.Left - FO(0) + x, Me.Top - FO(1) + y Else: FO(0) = x: FO(1) = y
End Sub

