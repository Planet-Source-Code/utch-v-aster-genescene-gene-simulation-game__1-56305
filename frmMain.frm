VERSION 5.00
Object = "{A32A88B3-817C-11D1-A762-00AA0044064C}#1.0#0"; "mscecomdlg.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00F3E9DA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GeneScene"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3D2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4125
      Left            =   2760
      TabIndex        =   4
      Top             =   4470
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command1 
         Caption         =   "Go!"
         Height          =   375
         Left            =   1448
         TabIndex        =   6
         Top             =   3540
         Width           =   1065
      End
      Begin VB.TextBox txtCodeDisp 
         BackColor       =   &H00F3E9DA&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2610
         Width           =   3615
      End
      Begin VB.Label lblBeaten 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have Beaten Level 99!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   427
         TabIndex        =   15
         Top             =   120
         Width           =   3300
      End
      Begin VB.Label lblPrepare 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prepare for Level 99!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3060
         Width           =   3270
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Points:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1230
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Points Earned:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   255
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   255
         X2              =   3705
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=     Total Points:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   11
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label lblPrev 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   10
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label lblEarned 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   9
         Top             =   1590
         Width           =   840
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   8
         Top             =   2070
         Width           =   840
      End
      Begin VB.Label lblCodeLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level 2 Code:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   2430
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3D6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4185
      Left            =   6990
      TabIndex        =   16
      Top             =   4470
      Width           =   7215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1905
         Left            =   270
         TabIndex        =   19
         Top             =   1410
         Width           =   6705
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   3990
            Picture         =   "frmMain.frx":058A
            ScaleHeight     =   915
            ScaleWidth      =   915
            TabIndex        =   21
            Top             =   450
            Width           =   915
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   5130
            Picture         =   "frmMain.frx":0695
            ScaleHeight     =   915
            ScaleWidth      =   915
            TabIndex        =   20
            Top             =   450
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00F3E9DA&
            FillColor       =   &H00F3E9DA&
            FillStyle       =   0  'Solid
            Height          =   90
            Left            =   0
            Top             =   0
            Width           =   46500
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":0782
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   180
            TabIndex        =   24
            Top             =   270
            Width           =   3705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "'Square'"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4095
            TabIndex        =   23
            Top             =   1410
            Width           =   705
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "'Plus'"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5355
            TabIndex        =   22
            Top             =   1410
            Width           =   465
         End
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Begin"
         Default         =   -1  'True
         Height          =   375
         Left            =   2865
         TabIndex        =   18
         Top             =   3510
         Width           =   1515
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   555
         Picture         =   "frmMain.frx":08E8
         ScaleHeight     =   1275
         ScaleWidth      =   6135
         TabIndex        =   17
         Top             =   210
         Width           =   6135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2265
      Left            =   4290
      TabIndex        =   25
      Top             =   1720
      Width           =   2865
      Begin VB.PictureBox picSquare 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F3E9DA&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   180
         Picture         =   "frmMain.frx":15AC
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   36
         Top             =   60
         Width           =   915
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   90
         TabIndex        =   33
         Top             =   1890
         Width           =   2625
      End
      Begin VB.PictureBox picOffColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picOnColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   29
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picPlus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F3E9DA&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   180
         Picture         =   "frmMain.frx":16B7
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   28
         Top             =   60
         Width           =   915
      End
      Begin VB.OptionButton optSquare 
         BackColor       =   &H00F3E9DA&
         Caption         =   "'Square' Mode"
         Height          =   195
         Left            =   1320
         TabIndex        =   27
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optPlus 
         BackColor       =   &H00F3E9DA&
         Caption         =   "'Plus' Mode"
         Height          =   195
         Left            =   1320
         TabIndex        =   26
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level Code:"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restart Current Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1155
         TabIndex        =   34
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00F3E9DA&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   675
         Left            =   0
         Top             =   1590
         Width           =   2835
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'Off' Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   32
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "'On' Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   570
         TabIndex        =   31
         Top             =   1230
         Width           =   690
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00F3E9DA&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   435
         Left            =   0
         Top             =   1110
         Width           =   2805
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00F3E9DA&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1065
         Left            =   0
         Top             =   0
         Width           =   2820
      End
   End
   Begin CEComDlgCtl.CommonDialog CD 
      Left            =   2880
      Top             =   1590
      _cx             =   847
      _cy             =   847
      CancelError     =   0   'False
      Color           =   0
      DefaultExt      =   ""
      DialogTitle     =   ""
      FileName        =   ""
      Filter          =   ""
      FilterIndex     =   0
      Flags           =   0
      HelpCommand     =   0
      HelpContext     =   ""
      HelpFile        =   ""
      InitDir         =   ""
      MaxFileSize     =   256
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   ""
      FontSize        =   10
      FontUnderline   =   0   'False
      Max             =   0
      Min             =   0
      FontStrikethru  =   0   'False
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   2
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   37
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   2
         Left            =   0
         Picture         =   "frmMain.frx":17A4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   3
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   38
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   3
         Left            =   0
         Picture         =   "frmMain.frx":1937
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   4
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   4
         Left            =   0
         Picture         =   "frmMain.frx":1ACA
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   5
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   5
         Left            =   0
         Picture         =   "frmMain.frx":1C5D
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   6
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   41
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   6
         Left            =   0
         Picture         =   "frmMain.frx":1DF0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   7
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   42
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   7
         Left            =   0
         Picture         =   "frmMain.frx":1F7C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   8
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   43
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   8
         Left            =   0
         Picture         =   "frmMain.frx":2108
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   9
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   44
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   9
         Left            =   0
         Picture         =   "frmMain.frx":2294
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   10
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   45
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   10
         Left            =   0
         Picture         =   "frmMain.frx":2420
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   13
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   46
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   13
         Left            =   0
         Picture         =   "frmMain.frx":25AC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   14
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   47
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   14
         Left            =   0
         Picture         =   "frmMain.frx":2738
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   15
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   48
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   15
         Left            =   0
         Picture         =   "frmMain.frx":28C4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   16
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   49
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   16
         Left            =   0
         Picture         =   "frmMain.frx":2A50
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   17
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   50
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   17
         Left            =   0
         Picture         =   "frmMain.frx":2BDC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   18
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   51
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   18
         Left            =   0
         Picture         =   "frmMain.frx":2D68
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   19
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   52
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   19
         Left            =   0
         Picture         =   "frmMain.frx":2EF4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   20
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   53
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   20
         Left            =   0
         Picture         =   "frmMain.frx":3080
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   21
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   54
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   21
         Left            =   0
         Picture         =   "frmMain.frx":320C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   22
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   55
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   22
         Left            =   0
         Picture         =   "frmMain.frx":3398
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   23
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   56
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   23
         Left            =   0
         Picture         =   "frmMain.frx":3524
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   24
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   57
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   24
         Left            =   0
         Picture         =   "frmMain.frx":36B0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   25
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   58
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   25
         Left            =   0
         Picture         =   "frmMain.frx":383C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   26
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   59
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   26
         Left            =   0
         Picture         =   "frmMain.frx":39C8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   27
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   60
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   27
         Left            =   0
         Picture         =   "frmMain.frx":3B54
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   28
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   61
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   28
         Left            =   0
         Picture         =   "frmMain.frx":3CE0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   29
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   62
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   29
         Left            =   0
         Picture         =   "frmMain.frx":3E6C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   30
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   63
      Top             =   870
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   30
         Left            =   0
         Picture         =   "frmMain.frx":3FF8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   31
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   64
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   31
         Left            =   0
         Picture         =   "frmMain.frx":4184
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   32
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   65
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   32
         Left            =   0
         Picture         =   "frmMain.frx":4310
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   33
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   66
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   33
         Left            =   0
         Picture         =   "frmMain.frx":449C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   34
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   67
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   34
         Left            =   0
         Picture         =   "frmMain.frx":4628
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   35
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   68
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   35
         Left            =   0
         Picture         =   "frmMain.frx":47B4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   36
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   69
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   36
         Left            =   0
         Picture         =   "frmMain.frx":4940
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   37
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   70
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   37
         Left            =   0
         Picture         =   "frmMain.frx":4ACC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   38
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   71
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   38
         Left            =   0
         Picture         =   "frmMain.frx":4C58
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   39
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   72
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   39
         Left            =   0
         Picture         =   "frmMain.frx":4DE4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   40
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   73
      Top             =   1260
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   40
         Left            =   0
         Picture         =   "frmMain.frx":4F70
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   41
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   74
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   41
         Left            =   0
         Picture         =   "frmMain.frx":50FC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   42
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   75
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   42
         Left            =   0
         Picture         =   "frmMain.frx":5288
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   43
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   76
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   43
         Left            =   0
         Picture         =   "frmMain.frx":5414
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   44
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   77
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   44
         Left            =   0
         Picture         =   "frmMain.frx":55A0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   45
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   78
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   45
         Left            =   0
         Picture         =   "frmMain.frx":572C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   46
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   79
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   46
         Left            =   0
         Picture         =   "frmMain.frx":58B8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   47
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   80
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   47
         Left            =   0
         Picture         =   "frmMain.frx":5A44
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   48
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   81
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   48
         Left            =   0
         Picture         =   "frmMain.frx":5BD0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   49
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   82
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   49
         Left            =   0
         Picture         =   "frmMain.frx":5D5C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   50
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   83
      Top             =   1650
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   50
         Left            =   0
         Picture         =   "frmMain.frx":5EE8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   51
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   84
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   51
         Left            =   0
         Picture         =   "frmMain.frx":6074
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   52
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   85
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   52
         Left            =   0
         Picture         =   "frmMain.frx":6200
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   53
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   86
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   53
         Left            =   0
         Picture         =   "frmMain.frx":638C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   54
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   87
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   54
         Left            =   0
         Picture         =   "frmMain.frx":6518
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   55
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   88
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   55
         Left            =   0
         Picture         =   "frmMain.frx":66A4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   56
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   89
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   56
         Left            =   0
         Picture         =   "frmMain.frx":6830
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   57
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   90
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   57
         Left            =   0
         Picture         =   "frmMain.frx":69BC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   58
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   91
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   58
         Left            =   0
         Picture         =   "frmMain.frx":6B48
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   59
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   92
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   59
         Left            =   0
         Picture         =   "frmMain.frx":6CD4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   60
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   93
      Top             =   2040
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   60
         Left            =   0
         Picture         =   "frmMain.frx":6E60
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   61
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   94
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   61
         Left            =   0
         Picture         =   "frmMain.frx":6FEC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   62
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   95
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   62
         Left            =   0
         Picture         =   "frmMain.frx":7178
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   63
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   96
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   63
         Left            =   0
         Picture         =   "frmMain.frx":7304
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   64
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   97
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   64
         Left            =   0
         Picture         =   "frmMain.frx":7490
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   65
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   98
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   65
         Left            =   0
         Picture         =   "frmMain.frx":761C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   66
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   99
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   66
         Left            =   0
         Picture         =   "frmMain.frx":77A8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   67
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   100
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   67
         Left            =   0
         Picture         =   "frmMain.frx":7934
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   68
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   101
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   68
         Left            =   0
         Picture         =   "frmMain.frx":7AC0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   69
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   102
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   69
         Left            =   0
         Picture         =   "frmMain.frx":7C4C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   70
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   103
      Top             =   2430
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   70
         Left            =   0
         Picture         =   "frmMain.frx":7DD8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   71
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   104
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   71
         Left            =   0
         Picture         =   "frmMain.frx":7F64
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   72
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   105
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   72
         Left            =   0
         Picture         =   "frmMain.frx":80F0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   73
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   106
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   73
         Left            =   0
         Picture         =   "frmMain.frx":827C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   74
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   107
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   74
         Left            =   0
         Picture         =   "frmMain.frx":8408
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   75
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   108
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   75
         Left            =   0
         Picture         =   "frmMain.frx":8594
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   76
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   109
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   76
         Left            =   0
         Picture         =   "frmMain.frx":8720
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   77
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   110
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   77
         Left            =   0
         Picture         =   "frmMain.frx":88AC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   78
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   111
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   78
         Left            =   0
         Picture         =   "frmMain.frx":8A38
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   79
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   112
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   79
         Left            =   0
         Picture         =   "frmMain.frx":8BC4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   80
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   113
      Top             =   2820
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   80
         Left            =   0
         Picture         =   "frmMain.frx":8D50
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   81
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   114
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   81
         Left            =   0
         Picture         =   "frmMain.frx":8EDC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   82
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   115
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   82
         Left            =   0
         Picture         =   "frmMain.frx":9068
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   83
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   116
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   83
         Left            =   0
         Picture         =   "frmMain.frx":91F4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   84
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   117
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   84
         Left            =   0
         Picture         =   "frmMain.frx":9380
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   85
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   118
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   85
         Left            =   0
         Picture         =   "frmMain.frx":950C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   86
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   119
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   86
         Left            =   0
         Picture         =   "frmMain.frx":9698
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   87
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   120
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   87
         Left            =   0
         Picture         =   "frmMain.frx":9824
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   89
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   121
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   89
         Left            =   0
         Picture         =   "frmMain.frx":99B0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   90
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   122
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   90
         Left            =   0
         Picture         =   "frmMain.frx":9B3C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   91
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   123
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   91
         Left            =   0
         Picture         =   "frmMain.frx":9CC8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   92
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   124
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   92
         Left            =   0
         Picture         =   "frmMain.frx":9E54
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   93
      Left            =   870
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   125
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   93
         Left            =   0
         Picture         =   "frmMain.frx":9FE0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   94
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   126
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   94
         Left            =   0
         Picture         =   "frmMain.frx":A16C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   95
      Left            =   1650
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   127
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   95
         Left            =   0
         Picture         =   "frmMain.frx":A2F8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   96
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   128
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   96
         Left            =   0
         Picture         =   "frmMain.frx":A484
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   97
      Left            =   2430
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   129
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   97
         Left            =   0
         Picture         =   "frmMain.frx":A610
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   98
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   130
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   98
         Left            =   0
         Picture         =   "frmMain.frx":A79C
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   99
      Left            =   3210
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   131
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   99
         Left            =   0
         Picture         =   "frmMain.frx":A928
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   100
      Left            =   3600
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   132
      Top             =   3600
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   100
         Left            =   0
         Picture         =   "frmMain.frx":AAB4
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   88
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   133
      Top             =   3210
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   88
         Left            =   0
         Picture         =   "frmMain.frx":AC40
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   1
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   134
      Top             =   90
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   11
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   135
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   11
         Left            =   0
         Picture         =   "frmMain.frx":ADCC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   12
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   136
      Top             =   480
      Width           =   400
      Begin VB.Image Image1 
         Height          =   375
         Index           =   12
         Left            =   0
         Picture         =   "frmMain.frx":AF58
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2385
      Left            =   4200
      Top             =   1665
      Width           =   3015
   End
   Begin VB.Label lblmoves 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moves: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5310
      TabIndex        =   3
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label lblLights 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lights: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5325
      TabIndex        =   2
      Top             =   795
      Width           =   735
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5625
      TabIndex        =   1
      Top             =   60
      Width           =   135
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5340
      TabIndex        =   0
      Top             =   570
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1410
      Picture         =   "frmMain.frx":B0E4
      Top             =   5130
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   4020
      Left            =   30
      Top             =   30
      Width           =   4035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLevel As Integer
Dim mOnColor As Long
Dim mOffColor As Long
Dim mScore As Long
Dim mPoints As Long
Dim mMoves As Double
Dim InitialLoad As Boolean

Private Sub cmdBegin_Click()
  InitialLoad = True
  Call DoLevel(0)
  Frame3D6.Visible = False
  InitialLoad = False
End Sub

Private Sub Command1_Click()
  DoLevel (mLevel + 1)
  Frame3D2.Visible = False
End Sub

Private Sub Command2_Click()
  Call DoLevel(mLevel + 1)
End Sub

Private Sub Form_Load()
  
  Call InitializeForm
End Sub

Sub InitializeForm()
  
  Visible = True
  Do Until Visible
    DoEvents
  Loop
  
  Dim S As String
  
  Dim X As Integer
  For X = 1 To 50
    Debug.Print "Level " & X & " - " & Encrypt("level" & X)
  Next
  
  Frame3D2.Move 0, 0
  Frame3D2.BackColor = BackColor
  Frame3D6.Move 0, 0
  Frame3D6.BackColor = BackColor

  S = GetSetting("Genes", "Settings", "onColor")
  If S = "" Then
    mOnColor = vbRed
    Call SaveSetting("Genes", "Settings", "onColor", mOnColor)
  Else
    mOnColor = Val(S)
  End If
  picOnColor.BackColor = mOnColor
  
  S = GetSetting("Genes", "Settings", "offColor")
  If S = "" Then
    mOffColor = vbBlack
    Call SaveSetting("Genes", "Settings", "offColor", mOffColor)
  Else
    mOffColor = Val(S)
  End If
  picOffColor.BackColor = mOffColor
  
  If Val(GetSetting("Genes", "Settings", "clickMode")) = 0 Then
    optSquare.Value = True
  Else
    optPlus.Value = True
  End If
  
  For X = 1 To 100
    Image1(X).Picture = Image2.Picture
  Next
  
End Sub

Private Sub Image1_Click(Index As Integer)
  picBox_Click (Index)
End Sub

Private Sub Label10_Click()
  Dim Answer As String
  Answer = MsgBox("Are you sure you want to restart this level?", vbYesNo + vbQuestion, "Restart Level?")
  If Answer = vbNo Then Exit Sub
  
  Call DoLevel(mLevel)
  mMoves = 0
  lblmoves = "Moves: " & mMoves
End Sub

Private Sub Label11_Click()

End Sub

Private Sub optPlus_Click()
  Call SaveSetting("Genes", "Settings", "clickMode", "1")
  picSquare.Visible = False
  picPlus.Visible = True
End Sub

Private Sub optSquare_Click()
  Call SaveSetting("Genes", "Settings", "clickMode", "0")
  picSquare.Visible = True
  picPlus.Visible = False
End Sub

Private Sub picBox_Click(Index As Integer)
  Call ChangeColor(Index)
  If Index Mod 10 <> 0 Then ChangeColor (Index + 1)
  If Index Mod 10 <> 1 Then ChangeColor (Index - 1)
  If Index <= 90 Then ChangeColor (Index + 10)
  If Index >= 11 Then ChangeColor (Index - 10)
  If optSquare.Value Then
    If Index >= 11 Then
      If Index Mod 10 <> 0 Then ChangeColor (Index - 9)
      If Index Mod 10 <> 1 Then ChangeColor (Index - 11)
    End If
    If Index <= 90 Then
      If Index Mod 10 <> 1 Then ChangeColor (Index + 9)
      If Index Mod 10 <> 0 Then ChangeColor (Index + 11)
    End If
  End If
  
  mMoves = mMoves + 1
  lblmoves = "Moves: " & mMoves
  
  Call CheckForWin
End Sub

Sub CheckForWin()
  Dim X As Integer
  Dim LightsOn As Integer
  
  For X = 1 To 100
    If picBox(X).BackColor = mOnColor Then LightsOn = LightsOn + 1
  Next
  
  lblLights = "Lights Remaining: " & LightsOn
  
  If LightsOn = 0 Then
    mScore = mScore + mPoints
    lblScore = "Score: " & mScore
    lblBeaten = "You Have Beaten Level " & mLevel + 1 & "!"
    lblCodeLevel = "Code for Level " & mLevel + 2 & ":"
    txtCodeDisp = Trim$(Encrypt("level" & mLevel + 2))
    lblPrev = mScore - mPoints
    lblEarned = mPoints
    lblTotal = mScore
    
    mMoves = 0
    lblmoves = "Moves: " & mMoves
    
    lblPrepare = "Prepare for Level " & mLevel + 2
    Frame3D2.Visible = True
  End If
End Sub

Sub DoLevel(intLevel As Integer)
  mLevel = intLevel
  lblLevel = "Level " & intLevel + 1
  Dim bData As String
  If intLevel = 0 Then
    bData = "0100010001" & _
            "1110111011" & _
            "0101010101" & _
            "0011101110" & _
            "0101010101" & _
            "1110111011" & _
            "0101010101" & _
            "0011101110" & _
            "0101010101" & _
            "1110111011"
    mPoints = 50
  ElseIf intLevel = 1 Then
    bData = "1000000001" & _
            "0100000010" & _
            "0010000100" & _
            "0001001000" & _
            "0000110000" & _
            "0000110000" & _
            "0001001000" & _
            "0010000100" & _
            "0100000010" & _
            "1000000001"
    mPoints = 100
  ElseIf mLevel = 2 Then
    bData = "1010101010" & _
            "0101010101" & _
            "1010101010" & _
            "0101010101" & _
            "1010101010" & _
            "0101010101" & _
            "1010101010" & _
            "0101010101" & _
            "1010101010" & _
            "0101010101"
    mPoints = 125
  ElseIf mLevel = 3 Then
    bData = "0111111110" & _
            "0111111110" & _
            "0000110000" & _
            "0000110000" & _
            "0000110000" & _
            "0000110000" & _
            "0000110000" & _
            "0000110000" & _
            "0111110000" & _
            "0111110000"
    mPoints = 150
  ElseIf mLevel = 4 Then
    bData = "0100100100" & _
            "1001001001" & _
            "0010010010" & _
            "0100100100" & _
            "1001001001" & _
            "0010010010" & _
            "0100100100" & _
            "1001001001" & _
            "0010010010" & _
            "0100100100"
    mPoints = 200
  ElseIf mLevel = 5 Then
    bData = "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100" & _
            "0100100100"
    mPoints = 250
  ElseIf mLevel = 6 Then
    bData = "1111111111" & _
            "0010000100" & _
            "0010000100" & _
            "1111111111" & _
            "0010000100" & _
            "0010000100" & _
            "1111111111" & _
            "0010000100" & _
            "0010000100" & _
            "1111111111"
    mPoints = 300
  ElseIf mLevel = 7 Then
    bData = "1001001001" & _
            "0110110110" & _
            "1001001001" & _
            "1001011001" & _
            "0110100110" & _
            "1001001001" & _
            "1001001001" & _
            "0110110110" & _
            "1001001001" & _
            "0110110110"
    mPoints = 300
  ElseIf mLevel = 8 Then
    bData = "1010101010" & _
            "0101010101" & _
            "1--0101010" & _
            "0--1010101" & _
            "1010--1010" & _
            "0101--0101" & _
            "1010101--0" & _
            "0101010--1" & _
            "1010101010" & _
            "0101010101"
    mPoints = 400
  ElseIf mLevel = 9 Then
    bData = "1111111111" & _
            "0010--0100" & _
            "0010--0100" & _
            "1111--1111" & _
            "0010--0100" & _
            "0010--0100" & _
            "1111--1111" & _
            "0010--0100" & _
            "0010--0100" & _
            "1111111111"
    mPoints = 400
  ElseIf mLevel = 10 Then
    bData = "0100100100" & _
            "10------01" & _
            "0010010010" & _
            "0100100100" & _
            "1001-01001" & _
            "0010-10010" & _
            "0100-00100" & _
            "1-01001-01" & _
            "0-10010-10" & _
            "0-00-00-00"
    mPoints = 500
  Else
    Exit Sub
  End If
  Call PopulateBoard(bData)
  Call CheckForWin
  DoEvents
End Sub

Sub PopulateBoard(BoardData As String)
  Dim X As Integer
  For X = 1 To 100
    picBox(X).BackColor = IIf(Mid$(BoardData, X, 1) = "1", mOnColor, mOffColor)
    If Mid$(BoardData, X, 1) = "-" Then
      picBox(X).BackColor = mOffColor
      picBox(X).Visible = False
    Else
      picBox(X).Visible = True
    End If
  Next
End Sub

Sub ChangeColor(Index)
  On Error Resume Next
  picBox(Index).BackColor = IIf(picBox(Index).BackColor = mOffColor, mOnColor, mOffColor)
End Sub

Private Sub picOffColor_Click()
  If SelectColor Then
    If CD.Color = picOnColor.BackColor Then
      MsgBox "Color Conflict"
      Exit Sub
    End If
    Dim X As Integer
    picOffColor.BackColor = CD.Color
    For X = 1 To 100
      If picBox(X).BackColor = mOffColor Then picBox(X).BackColor = CD.Color
    Next
    mOffColor = CD.Color
    Call SaveSetting("Genes", "Settings", "offColor", mOffColor)
  End If
End Sub

Private Sub picOnColor_Click()
  If SelectColor Then
    If CD.Color = picOffColor.BackColor Then
      MsgBox "Color Conflict"
      Exit Sub
    End If
    Dim X As Integer
    picOnColor.BackColor = CD.Color
    For X = 1 To 100
      If picBox(X).BackColor = mOnColor Then picBox(X).BackColor = CD.Color
    Next
    mOnColor = CD.Color
    Call SaveSetting("Genes", "Settings", "onColor", mOnColor)
  End If
End Sub

Function SelectColor() As Boolean
  CD.CancelError = True
  On Error GoTo Err
  CD.ShowColor
  SelectColor = True
  Exit Function
  
Err:
  SelectColor = False
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    Dim i As Integer
    If StringToDecrypt = "" Then Exit Function
    If AlphaDecoding Then
        Decrypt = StringToDecrypt
        StringToDecrypt = ""
        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 147)
        Next i
    End If
    Decrypt = ""
    Do Until StringToDecrypt = ""
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
    Loop
    Exit Function
ErrorHandler:
    Decrypt = ""
End Function

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim Char As String
    Encrypt = ""
    If StringToEncrypt = "" Then Exit Function
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    If AlphaEncoding Then
        StringToEncrypt = Encrypt
        Encrypt = ""
        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        Next i
    End If
    Exit Function
ErrorHandler:
    Encrypt = ""
End Function

Private Sub txtCode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Dim S As String
    S = LCase$(Decrypt(txtCode))
    If Left$(S, 5) = "level" Then
      S = Mid(S, 6)
      Call DoLevel(Val(S) - 1)
      txtCode = ""
      KeyAscii = 0
    End If
  End If
End Sub
