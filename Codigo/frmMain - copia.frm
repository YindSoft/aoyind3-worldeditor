VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   13845
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   19110
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   923
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Luces (Se borran con Layer3)"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   9360
      TabIndex        =   170
      Top             =   10680
      Width           =   2535
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "\"
         Height          =   255
         Index           =   5
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "|"
         Height          =   255
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "/"
         Height          =   255
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         Height          =   255
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         Height          =   255
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         Height          =   255
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "/"
         Height          =   255
         Index           =   6
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "|"
         Height          =   255
         Index           =   8
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton bLuz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "\"
         Height          =   255
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optNoche 
         BackColor       =   &H00000000&
         Caption         =   "Noche"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   183
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optDia 
         BackColor       =   &H00000000&
         Caption         =   "Día"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   182
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox tRango 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   171
         Text            =   "5"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rango: (1-15)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   181
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Mapa2"
      Height          =   255
      Left            =   15840
      TabIndex        =   83
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox chkAcoplar 
      BackColor       =   &H00000000&
      Caption         =   "Acoplar"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   18120
      TabIndex        =   169
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Re Cargar Areas"
      Height          =   375
      Left            =   17640
      TabIndex        =   91
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CheckBox chkCostas 
      BackColor       =   &H00000000&
      Caption         =   "Auto-Costas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   167
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox picMapGrh 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   4200
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   166
      Top             =   13680
      Width           =   7680
   End
   Begin VB.TextBox tZX1 
      Height          =   285
      Left            =   15000
      TabIndex        =   160
      Top             =   11160
      Width           =   615
   End
   Begin VB.TextBox tZY1 
      Height          =   285
      Left            =   15000
      TabIndex        =   159
      Top             =   11520
      Width           =   615
   End
   Begin VB.TextBox tZX2 
      Height          =   285
      Left            =   15000
      TabIndex        =   158
      Top             =   11880
      Width           =   615
   End
   Begin VB.TextBox tZY2 
      Height          =   285
      Left            =   15000
      TabIndex        =   157
      Top             =   12240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dibujar Areas"
      Height          =   255
      Left            =   17880
      TabIndex        =   156
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BMP Parcial"
      Height          =   375
      Left            =   5640
      TabIndex        =   155
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton btnMidi 
      Caption         =   "P"
      Height          =   255
      Index           =   4
      Left            =   18240
      TabIndex        =   154
      Top             =   12600
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Caption         =   "P"
      Height          =   255
      Index           =   3
      Left            =   18240
      TabIndex        =   153
      Top             =   12240
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Caption         =   "P"
      Height          =   255
      Index           =   2
      Left            =   18240
      TabIndex        =   152
      Top             =   11880
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Caption         =   "P"
      Height          =   255
      Index           =   1
      Left            =   18240
      TabIndex        =   151
      Top             =   11520
      Width           =   255
   End
   Begin VB.CommandButton btnMidi 
      Caption         =   "P"
      Height          =   255
      Index           =   0
      Left            =   18240
      TabIndex        =   150
      Top             =   11160
      Width           =   255
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   4
      Left            =   16560
      TabIndex        =   148
      Top             =   12600
      Width           =   1695
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   3
      Left            =   16560
      TabIndex        =   146
      Top             =   12240
      Width           =   1695
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   2
      Left            =   16560
      TabIndex        =   144
      Top             =   11880
      Width           =   1695
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   1
      Left            =   16560
      TabIndex        =   142
      Top             =   11520
      Width           =   1695
   End
   Begin VB.TextBox tMusica 
      Height          =   285
      Index           =   0
      Left            =   16560
      TabIndex        =   140
      Top             =   11160
      Width           =   1695
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   1080
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   1320
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   1080
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   1080
      Width           =   240
   End
   Begin VB.CommandButton cmdDM 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox pGuardar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4320
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   134
      Top             =   11280
      Width           =   990
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   131
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "BMP"
      Height          =   375
      Left            =   4560
      TabIndex        =   132
      Top             =   10680
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45000
      Left            =   -28920
      ScaleHeight     =   3000
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2200
      TabIndex        =   130
      Top             =   11040
      Width           =   33000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modifica"
      Height          =   375
      Left            =   18240
      TabIndex        =   129
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox tCant 
      Height          =   285
      Left            =   17160
      TabIndex        =   127
      Text            =   "1"
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "StartCopy"
      Height          =   375
      Left            =   15240
      TabIndex        =   125
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Pintar Rect"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   124
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Ti"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   123
      Top             =   720
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "De"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   122
      Top             =   720
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Nie"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   121
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   600
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   120
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape29 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   120
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   600
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   119
      Top             =   1200
      Width           =   255
      Begin VB.Shape Shape22 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   90
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   118
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape25 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   117
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape28 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   116
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   115
      Top             =   1200
      Width           =   255
      Begin VB.Shape Shape18 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   600
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   114
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape27 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape Shape23 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   120
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   113
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape24 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox tp1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   112
      Tag             =   "102"
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape26 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   111
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape17 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   90
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   110
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape15 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   109
      Top             =   1440
      Width           =   255
      Begin VB.Shape Shape16 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   90
         Width           =   210
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   108
      Top             =   1200
      Width           =   255
      Begin VB.Shape Shape10 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   90
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   107
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   106
      Top             =   1200
      Width           =   255
      Begin VB.Shape Shape7 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   105
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape14 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   90
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   1200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   104
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape12 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox tp2 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   960
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   103
      Top             =   960
      Width           =   255
      Begin VB.Shape Shape13 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   225
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Re Cargar Zonas"
      Height          =   375
      Left            =   16800
      TabIndex        =   102
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Agregar Zona"
      Height          =   375
      Left            =   15240
      TabIndex        =   101
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   15240
      TabIndex        =   100
      Text            =   "Nombre"
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Segura"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   17280
      TabIndex        =   99
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Guardar Zona"
      Height          =   375
      Left            =   15240
      TabIndex        =   98
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Guardar Area"
      Height          =   375
      Left            =   16440
      TabIndex        =   94
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Agregar Area"
      Height          =   375
      Left            =   15240
      TabIndex        =   93
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ListBox lstNpc 
      Height          =   1425
      Left            =   15240
      TabIndex        =   92
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   18240
      TabIndex        =   89
      Top             =   7440
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   15240
      TabIndex        =   88
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Listar"
      Height          =   375
      Left            =   17760
      TabIndex        =   87
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox tNPC 
      Height          =   285
      Left            =   15840
      TabIndex        =   86
      Text            =   "1"
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   17760
      TabIndex        =   85
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Mapa1"
      Height          =   255
      Left            =   15240
      TabIndex        =   84
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Dibujar Zonas"
      Height          =   255
      Left            =   16680
      TabIndex        =   82
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.PictureBox picRadar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   15240
      Picture         =   "frmMain.frx":628A
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   78
      Top             =   1440
      Width           =   3300
      Begin VB.TextBox tpX 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2520
         TabIndex        =   80
         Text            =   "1"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox tpY 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2880
         TabIndex        =   79
         Text            =   "1"
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1485
         Left            =   0
         Top             =   0
         Width           =   45
      End
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   90
         Left            =   360
         Top             =   600
         Width           =   105
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   11835
      TabIndex        =   77
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Particulas"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":3683C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   9120
      Left            =   4560
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   704
      TabIndex        =   76
      Top             =   1440
      Width           =   10560
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   13320
      TabIndex        =   29
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":3748E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   10320
      TabIndex        =   28
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":37A54
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8955
      TabIndex        =   27
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":37F55
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7440
      TabIndex        =   26
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":38309
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5925
      TabIndex        =   25
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":3868A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   5160
      TabIndex        =   24
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":3BCEA
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   435
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4275
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmMain.frx":3F230
      Top             =   6360
      Width           =   4320
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      Picture         =   "frmMain.frx":3F270
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   4
      Top             =   1800
      Width           =   4395
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   66
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   64
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":518D4
         Left            =   1080
         List            =   "frmMain.frx":518E4
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   53
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   52
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":518F4
         Left            =   120
         List            =   "frmMain.frx":518F6
         Sorted          =   -1  'True
         TabIndex        =   51
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   56
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":518F8
         Left            =   3360
         List            =   "frmMain.frx":518FA
         TabIndex        =   47
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":518FC
         Left            =   840
         List            =   "frmMain.frx":518FE
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":51900
         Left            =   120
         List            =   "frmMain.frx":51902
         TabIndex        =   46
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   45
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":51904
         Left            =   840
         List            =   "frmMain.frx":51906
         TabIndex        =   38
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":51908
         Left            =   3360
         List            =   "frmMain.frx":5190A
         TabIndex        =   37
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":5190C
         Left            =   120
         List            =   "frmMain.frx":5190E
         TabIndex        =   35
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":51910
         Left            =   120
         List            =   "frmMain.frx":51912
         TabIndex        =   34
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   8
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   9
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   30
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   49
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   50
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   63
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":51914
         Left            =   840
         List            =   "frmMain.frx":51916
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   58
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":51918
         Left            =   120
         List            =   "frmMain.frx":5191A
         TabIndex        =   59
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":5191C
         Left            =   3360
         List            =   "frmMain.frx":5191E
         TabIndex        =   60
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   74
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   73
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   72
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   21
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   60
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6270
      Width           =   4455
      Begin VB.PictureBox PreviewGrh 
         BackColor       =   &H00004000&
         FillColor       =   &H00C0C0C0&
         Height          =   3780
         Left            =   45
         ScaleHeight     =   248
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   287
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   4365
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   9840
      TabIndex        =   75
      Top             =   240
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":51920
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   18240
      TabIndex        =   90
      Top             =   7080
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PB2 
      Height          =   255
      Left            =   6840
      TabIndex        =   133
      Top             =   10680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label tNumZona 
      BackColor       =   &H00000000&
      Caption         =   "Nº Zona:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   17280
      TabIndex        =   168
      Top             =   10080
      Width           =   1815
   End
   Begin VB.Label POSX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X: ?? - Y:??"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   16680
      TabIndex        =   165
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   14640
      TabIndex        =   164
      Top             =   11160
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   14640
      TabIndex        =   163
      Top             =   11520
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   14640
      TabIndex        =   162
      Top             =   11880
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   14640
      TabIndex        =   161
      Top             =   12240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   15720
      TabIndex        =   149
      Top             =   12600
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   15720
      TabIndex        =   147
      Top             =   12240
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   15720
      TabIndex        =   145
      Top             =   11880
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   15720
      TabIndex        =   143
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Musica 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   15720
      TabIndex        =   141
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Cant:"
      Height          =   255
      Left            =   16560
      TabIndex        =   128
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "NPC:"
      Height          =   255
      Left            =   15240
      TabIndex        =   126
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Npc Areas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15240
      TabIndex        =   97
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblNum 
      BackColor       =   &H00000000&
      Caption         =   "Nº de Area: 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15240
      TabIndex        =   96
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00000000&
      Caption         =   "(0,0)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   17160
      TabIndex        =   95
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   329
      X2              =   329
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   337
      X2              =   337
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   336
      X2              =   336
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   328
      X2              =   328
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mmGuardarCliente 
         Caption         =   "&Guardar Mapa Cliente"
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarcomoBMP 
         Caption         =   "Guardar Render en &BMP"
      End
      Begin VB.Menu mnuGuardarcomoJPG 
         Caption         =   "Guardar Render en &JPG"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuParticle 
         Caption         =   "...&Particle"
      End
      Begin VB.Menu mmTrans 
         Caption         =   "...Transparencias Techo/Layer2"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLineAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit
Public MouseX As Integer
Public MouseY As Integer

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 by GS
'*************************************************
Dim OBJIndex As Long
Dim NPCIndex As Long
Dim X, Y, i
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
Dim Leer As New clsIniReader
i = n


Do While i > 0
    X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
    Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
    Select Case T
        Case 0
            If MapData(X, Y).OBJInfo.OBJIndex = 0 Then
                  i = i - 1
                  If cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                  Else
                    MapData(X, Y).Blocked = 0
                  End If
                  If cNumFunc(2).text > 0 Then
                      OBJIndex = cNumFunc(2).text
                      InitGrh MapData(X, Y).ObjGrh, ObjData(OBJIndex).grhindex
                      MapData(X, Y).OBJInfo.OBJIndex = OBJIndex
                      MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).text)
                      Select Case ObjData(OBJIndex).OBJType ' GS
                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                      End Select
                  End If
            End If
        Case 1
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).text > 0 Then
                        NPCIndex = cNumFunc(T - 1).text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex
                  End If
            End If
        Case 2
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).text >= 0 Then
                        NPCIndex = cNumFunc(T - 1).text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex
                  End If
           End If
        End Select
        DoEvents
Loop
End Sub

Private Sub bLuz_Click(Index As Integer)


If PonerLuz = Index + 1 Or (PonerLuz = 13 And Index = 5) Or (PonerLuz = 10 And Index = 7) Then
    PonerLuz = 0
    QuitarLuz
Else
    QuitarLuz
    Call mnuQuitarFunciones_Click
    If Index = 5 Then
        PonerLuz = 13
    ElseIf Index = 7 Then
        PonerLuz = 10
    Else
        PonerLuz = Index + 1
    End If
    bLuz(Index).BackColor = vbYellow
End If
End Sub
Sub QuitarLuz()
Dim i As Integer
PonerLuz = 0
For i = 0 To 8

    bLuz(i).BackColor = vbWhite
Next i
End Sub

Private Sub btnMidi_Click(Index As Integer)

   Dim ret As Integer

Dim Num As Integer

Num = Val(tMusica(Index).text)

   ' The following will open the sequencer with the C:\WIN31\CANYON.MID
   ' file. Canyon is the device_id.
If IsPlaying Then
   ret = mciSendString("close mus", 0&, 0, 0)
   IsPlaying = False
Else
   ret = mciSendString("open " & """" & App.Path & "\Midi\" & Num & ".mid" & """" & " type sequencer alias mus", 0&, 0, 0)

   ' The wait tells the MCI command to complete before returning control
   ' to the application.

   ret = mciSendString("play mus", 0&, 0, 0)
   IsPlaying = True
End If


   ' Close CANYON.MID file and sequencer device



End Sub

Private Sub cAgregarFuncalAzar_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If IsNumeric(cCantFunc(Index).text) = False Or cCantFunc(Index).text > 200 Then
    MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
    Exit Sub
End If
cAgregarFuncalAzar(Index).Enabled = False
Call PonerAlAzar(CInt(cCantFunc(Index).text), 1 + (IIf(Index = 2, -1, Index)))
cAgregarFuncalAzar(Index).Enabled = True
End Sub

Private Sub cCantFunc_Change(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Val(cCantFunc(Index)) < 1 Then
      cCantFunc(Index).text = 1
    End If
    If Val(cCantFunc(Index)) > 10000 Then
      cCantFunc(Index).text = 10000
    End If
End Sub

Private Sub cCapas_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/05/06
'*************************************************
    If Val(cCapas.text) < 1 Then
      cCapas.text = 1
    End If
    If Val(cCapas.text) > 4 Then
      cCapas.text = 4
    End If
    cCapas.Tag = vbNullString
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If KeyAscii = 13 Then
    Call Filtrar(Index)
End If
End Sub


Private Sub cGrh_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
If KeyAscii = 13 Then
    Call fPreviewGrh(cGrh.text)
    If frmMain.cGrh.ListCount > 5 Then
        frmMain.cGrh.RemoveItem 0
    End If
    frmMain.cGrh.AddItem frmMain.cGrh.text
End If
Exit Sub
Fallo:
    cGrh.text = 1

End Sub

Private Sub cInsertarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarFunc(Index).value = True Then
    cQuitarFunc(Index).Enabled = False
    cAgregarFuncalAzar(Index).Enabled = False
    If Index <> 2 Then cCantFunc(Index).Enabled = False
    Call modPaneles.EstSelectPanel((Index) + 3, True)
Else
    cQuitarFunc(Index).Enabled = True
    cAgregarFuncalAzar(Index).Enabled = True
    If Index <> 2 Then cCantFunc(Index).Enabled = True
    Call modPaneles.EstSelectPanel((Index) + 3, False)
End If
End Sub

Private Sub cInsertarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
If cInsertarTrans.value = True Then
    cQuitarTrans.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    cQuitarTrans.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub

Private Sub cInsertarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarTrigger.value = True Then
    cQuitarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    cQuitarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub


Private Sub cmdInformacionDelMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub cmdDM_Click(Index As Integer)
frmConfigSup.DespMosaic.value = vbChecked
Select Case Index
    Case 0 'A

frmConfigSup.DMLargo.text = Val(frmConfigSup.DMLargo.text) + 1
    Case 1 '<
    frmConfigSup.DMAncho.text = Val(frmConfigSup.DMAncho.text) + 1
    Case 2 '>
    frmConfigSup.DMAncho.text = Val(frmConfigSup.DMAncho.text) - 1
    Case 3 'V
    frmConfigSup.DMLargo.text = Val(frmConfigSup.DMLargo.text) - 1
    Case 4 '0
frmConfigSup.DMAncho.text = 0
frmConfigSup.DMLargo.text = 0
End Select

End Sub

Private Sub cmdQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call mnuQuitarFunciones_Click
End Sub



Private Sub Command1_Click()
Areas(SelArea).NPC(lstNpc.ListIndex + 1).NPCIndex = Val(tNPC.text)
Areas(SelArea).NPC(lstNpc.ListIndex + 1).Cantidad = Val(tCant.text)

lstNpc.List(lstNpc.ListIndex) = "(" & tCant.text & ") " & NombreNPC(Val(tNPC.text))

'Dim rgb_list(0 To 3) As Long
'rgb_list(0) = RGB(0, 255, 255)
'rgb_list(1) = RGB(0, 255, 255)
'rgb_list(2) = RGB(0, 255, 255)
'rgb_list(3) = RGB(0, 255, 255)
 
'Dim grh_list(1) As Long
'grh_list(1) = 15
'engine.Particle_Group_Create 50, 50, grh_list(), rgb_list(), 50, 4, 1, -1, 0, , 0, 0, 0, -20, 20, -10, -50, 10, 50, 8, 1, 1, 2, -15, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0
 
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim X As Integer
Dim Y As Integer
ReDim MapCopy(CopyRect.Right - CopyRect.Left, CopyRect.Bottom - CopyRect.Top)
'ReDim MapCopyD(CopyRect.Right - CopyRect.Left, CopyRect.Bottom - CopyRect.Top)
For X = CopyRect.Left To CopyRect.Right
    For Y = CopyRect.Top To CopyRect.Bottom
        MapCopy(X - CopyRect.Left, Y - CopyRect.Top) = MapData(X, Y)
    Next Y
Next X
'Command9.BackColor = vbRed
End Sub

Private Sub Command11_Click()
'CopyState = 0
'Command9.BackColor = Command11.BackColor
End Sub

Private Sub Command13_Click()
DibujarZonas
End Sub
Sub PintarMapa()

DibujarZonas
'DibujarAreas
End Sub
Sub DibujarAreas()
Dim i As Integer
picRadar.Cls
picRadar.ForeColor = vbYellow
For i = 1 To NumAreas
    If Areas(i).Mapa = UserMap Then
        picRadar.Line (Areas(i).X1 / 5, Areas(i).Y1 / 5)-(Areas(i).X2 / 5, Areas(i).Y1 / 5)
        picRadar.Line (Areas(i).X1 / 5, Areas(i).Y2 / 5)-(Areas(i).X2 / 5, Areas(i).Y2 / 5)
        picRadar.Line (Areas(i).X1 / 5, Areas(i).Y1 / 5)-(Areas(i).X1 / 5, Areas(i).Y2 / 5)
        picRadar.Line (Areas(i).X2 / 5, Areas(i).Y1 / 5)-(Areas(i).X2 / 5, Areas(i).Y2 / 5)
    End If
Next i
End Sub
Sub DibujarZonas()
Dim i As Integer
picRadar.Cls
picRadar.ForeColor = vbWhite
For i = 1 To NumZonas
    If Zonas(i).Mapa = UserMap Then
        picRadar.Line (Zonas(i).X1 / 5, Zonas(i).Y1 / 5)-(Zonas(i).X2 / 5, Zonas(i).Y1 / 5)
        picRadar.Line (Zonas(i).X1 / 5, Zonas(i).Y2 / 5)-(Zonas(i).X2 / 5, Zonas(i).Y2 / 5)
        picRadar.Line (Zonas(i).X1 / 5, Zonas(i).Y1 / 5)-(Zonas(i).X1 / 5, Zonas(i).Y2 / 5)
        picRadar.Line (Zonas(i).X2 / 5, Zonas(i).Y1 / 5)-(Zonas(i).X2 / 5, Zonas(i).Y2 / 5)
    End If
Next i
End Sub
Private Sub Command2_Click()

frmMain.picture1.Picture = LoadPicture(App.Path & "\mapa" & UserMap & ".bmp")
PB2.max = 1500
PB2.value = 1
UserPos.X = 12 + EditLimit.Left
UserPos.Y = 10 + EditLimit.Top
picture1.Cls
pGuardar.Cls

GrabarParcial = True
Grabar = True
End Sub

Private Sub Command22_Click()
'Dim X As Integer
'Dim Y As Integer
'            For X = CopyRect.Left To CopyRect.Right
'                For Y = CopyRect.Top To CopyRect.Bottom
'                    MapData(mCopyX + X - CopyRect.Left - AddMX, mCopyY + Y - CopyRect.Top - AddMY) = MapCopyD(X - CopyRect.Left, Y - CopyRect.Top)
'                Next Y
'            Next X
End Sub

Private Sub Command23_Click()

Grabar = False
Call SavePicture(frmMain.picture1.Image, App.Path & "\mapa" & UserMap & ".bmp")
End Sub

Private Sub Command24_Click()
AgregarZona = 0
SelZona = 0
Dim i As Integer
For i = 0 To 4
tMusica(i).text = "0"
Next i
Text6.text = ""
Check4.value = vbUnchecked
CargarZonas

End Sub

Private Sub Command25_Click()
Call mnuQuitarFunciones_Click
SelZona = 0
AgregarZona = 1
ZonaR.Left = 0
ZonaR.Top = 0
ZonaR.Right = 0
ZonaR.Bottom = 0

End Sub
Public Sub GuardaArea()
If AgregarArea = 3 Then
    NumAreas = NumAreas + 1
    Call WriteVar(DirDats & "areas.dat", "Config", "Cantidad", CStr(NumAreas))
    ReDim Preserve Areas(1 To NumAreas)
    SelArea = NumAreas
    With Areas(SelArea)
        .Mapa = UserMap
        .Npcs = 0
        .X1 = AreaR.Left
        .Y1 = AreaR.Top
        .X2 = AreaR.Right
        .Y2 = AreaR.Bottom
    End With
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Mapa", CStr(Areas(SelArea).Mapa))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "X1", CStr(Areas(SelArea).X1))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Y1", CStr(Areas(SelArea).Y1))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "X2", CStr(Areas(SelArea).X2))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Y2", CStr(Areas(SelArea).Y2))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Npcs", 0)
    SelectArea
    AgregarArea = 5
End If

End Sub
Private Sub Command26_Click()
Dim i As Integer
If AgregarZona = 3 Then
    NumZonas = NumZonas + 1
    Call WriteVar(DirDats & "zonas.dat", "Config", "Cantidad", CStr(NumZonas))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Nombre", Text6.text)
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Mapa", CStr(UserMap))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "X1", CStr(ZonaR.Left))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Y1", CStr(ZonaR.Top))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "X2", CStr(ZonaR.Right))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Y2", CStr(ZonaR.Bottom))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Segura", IIf(Check4.value = vbChecked, 1, 0))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Acoplar", IIf(chkAcoplar.value = vbChecked, 1, 0))
    For i = 0 To 4
        If Val(tMusica(i).text) > 0 Then Call WriteVar(DirDats & "zonas.dat", "Zona" & NumZonas, "Musica" & (i + 1), tMusica(i).text)
    Next i
    tNumZona.Caption = "N° Zona: " & NumZonas
    CargarZonas
ElseIf SelZona > 0 Then
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Nombre", Text6.text)
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Segura", IIf(Check4.value = vbChecked, 1, 0))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Acoplar", IIf(chkAcoplar.value = vbChecked, 1, 0))
    For i = 0 To 4
        If Val(tMusica(i).text) > 0 Then Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Musica" & (i + 1), tMusica(i).text)
    Next i
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "X1", CStr(tZX1.text))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Y1", CStr(tZY1.text))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "X2", CStr(tZX2.text))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & SelZona, "Y2", CStr(tZY2.text))
    CargarZonas
End If
AgregarZona = 0

End Sub

Private Sub Command27_Click()
Dim i As Integer
Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Mapa", CStr(Areas(SelArea).Mapa))
Areas(SelArea).X1 = Val(tZX1.text)
Areas(SelArea).Y1 = Val(tZY1.text)
Areas(SelArea).X2 = Val(tZX2.text)
Areas(SelArea).Y2 = Val(tZY2.text)
Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "X1", CStr(tZX1.text))
Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Y1", CStr(tZY1.text))
Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "X2", CStr(tZX2.text))
Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Y2", CStr(tZY2.text))

Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Npcs", CStr(Areas(SelArea).Npcs))
For i = 1 To Areas(SelArea).Npcs
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Npc" & i, CStr(Areas(SelArea).NPC(i).NPCIndex))
    Call WriteVar(DirDats & "areas.dat", "Area" & SelArea, "Cant" & i, CStr(Areas(SelArea).NPC(i).Cantidad))
Next i

AgregarArea = 0
End Sub

Private Sub Command28_Click()
Call mnuQuitarFunciones_Click
AgregarArea = 1
AreaR.Left = 0
AreaR.Top = 0
AreaR.Right = 0
AreaR.Bottom = 0
End Sub

Private Sub Command29_Click()
AgregarArea = 0
CargarAreas

End Sub

Private Sub Command3_Click()
DibujarAreas
End Sub

Private Sub Command30_Click()
Dim i As Integer
Areas(SelArea).Npcs = Areas(SelArea).Npcs - 1
For i = lstNpc.ListIndex + 1 To Areas(SelArea).Npcs
    Areas(SelArea).NPC(i) = Areas(SelArea).NPC(i + 1)
Next i

lstNpc.RemoveItem (lstNpc.ListIndex)

End Sub

Private Sub Command31_Click()
Areas(SelArea).Npcs = Areas(SelArea).Npcs + 1
ReDim Preserve Areas(SelArea).NPC(1 To Areas(SelArea).Npcs)
Areas(SelArea).NPC(Areas(SelArea).Npcs).NPCIndex = Val(tNPC.text)
Areas(SelArea).NPC(Areas(SelArea).Npcs).Cantidad = Val(tCant.text)
lstNpc.AddItem "(" & Val(tCant.text) & ") " & NombreNPC(Val(tNPC.text))
End Sub

Private Sub Command32_Click()
    Dim TempInt As Integer
    Dim Y As Integer
    Dim X As Integer
    Dim tmpByte As Byte
    Dim FreeFileInf As Long
    DoEvents
    
    FreeFileInf = FreeFile
    Open App.Path & "\Mapas\Mapa1.inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    frmMain.List1.Clear
    
    'Load arrays
    For Y = 1 To 100
        For X = 1 To 100
    
            
            '.inf file
            Get FreeFileInf, , tmpByte
            
            If tmpByte And 1 Then
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
            Dim i As Integer
            If tmpByte And 2 Then
                'Get and make NPC
                Get FreeFileInf, , TempInt
                
                For i = 0 To frmMain.List1.ListCount - 1
                    If ReadField(1, frmMain.List1.List(i), Asc("#")) = TempInt Then
                        frmMain.List1.List(i) = ReadField(1, frmMain.List1.List(i), Asc("#")) & "#" & ReadField(2, frmMain.List1.List(i), Asc("#")) + 1 & "#" & ReadField(3, frmMain.List1.List(i), Asc("#"))
                        i = -1
                        Exit For
                    End If
                Next i
                If i <> -1 Then
                    frmMain.List1.AddItem TempInt & "#1#" & NombreNPC(TempInt)
                End If
            End If
    
            If tmpByte And 4 Then
                'Get and make Object
                Get FreeFileInf, , TempInt
                Get FreeFileInf, , TempInt
            End If
    
        Next X
    Next Y

    Close FreeFileInf

End Sub
Public Function NombreNPC(NPCIndex As Integer) As String
Dim e As Integer
NombreNPC = NpcData(NPCIndex).name
Exit Function
            For e = 0 To frmMain.lListado(2).ListCount - 1
                If ReadField(2, frmMain.lListado(2).List(e), Asc("#")) = NPCIndex Then
                    NombreNPC = frmMain.lListado(2).List(e)
                    Exit Function
                End If
            Next e
            For e = 0 To frmMain.lListado(1).ListCount - 1
                If ReadField(2, frmMain.lListado(1).List(e), Asc("#")) = NPCIndex Then
                    NombreNPC = frmMain.lListado(1).List(e)
                    Exit Function
                End If
            Next e
End Function
Private Sub Command33_Click()
Dim Texto As String
If List1.ListIndex = -1 Then Exit Sub
Texto = List1.List(List1.ListIndex)

Areas(SelArea).Npcs = Areas(SelArea).Npcs + 1
ReDim Preserve Areas(SelArea).NPC(1 To Areas(SelArea).Npcs)

Areas(SelArea).NPC(Areas(SelArea).Npcs).NPCIndex = ReadField(1, Texto, Asc("#"))
Areas(SelArea).NPC(Areas(SelArea).Npcs).Cantidad = ReadField(2, Texto, Asc("#"))
lstNpc.AddItem "(" & Areas(SelArea).NPC(Areas(SelArea).Npcs).Cantidad & ") " & ReadField(3, Texto, Asc("#")) & "#" & Areas(SelArea).NPC(Areas(SelArea).Npcs).NPCIndex

End Sub

Private Sub Command34_Click()
UserMap = 1
picRadar.Picture = LoadPicture(App.Path & "\mapamini.bmp")
End Sub

Private Sub Command35_Click()
UserMap = 2
picRadar.Picture = LoadPicture(App.Path & "\mapamini2.bmp")
End Sub

Private Sub Command36_Click()
Call SavePicture(picRadar.Image, App.Path & "\mapamini" & IIf(UserMap = 1, "", "2") & ".bmp")

End Sub


Private Sub Command4_Click()
Dim X As Integer
Dim Y As Integer
Dim X2 As Integer, X1 As Integer
Dim i As Integer

For i = 1 To NumAreas
    With Areas(i)
    If .Mapa = 2 Then
        If .X2 <= 720 + 150 And .Y2 <= 1230 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        ElseIf .X2 <= 495 + 150 And .Y2 <= 1355 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        ElseIf .X1 >= 190 + 150 And .X2 <= 495 + 150 And .Y2 <= 1980 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        End If
    Call WriteVar(DirDats & "areas.dat", "Area" & i, "X1", CStr(.X1))
    Call WriteVar(DirDats & "areas.dat", "Area" & i, "Y1", CStr(.Y1))
    Call WriteVar(DirDats & "areas.dat", "Area" & i, "X2", CStr(.X2))
    Call WriteVar(DirDats & "areas.dat", "Area" & i, "Y2", CStr(.Y2))
    End If
    End With
Next i

For i = 1 To NumZonas
    With Zonas(i)
    If .Mapa = 2 Then
        If .X2 <= 720 + 150 And .Y2 <= 1230 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        ElseIf .X2 <= 495 + 150 And .Y2 <= 1355 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        ElseIf .X1 >= 190 + 150 And .X2 <= 495 + 150 And .Y2 <= 1980 Then
            .X1 = .X1 - 150
            .X2 = .X2 - 150
        End If
    Call WriteVar(DirDats & "zonas.dat", "Zona" & i, "X1", CStr(.X1))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & i, "Y1", CStr(.Y1))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & i, "X2", CStr(.X2))
    Call WriteVar(DirDats & "zonas.dat", "Zona" & i, "Y2", CStr(.Y2))
   End If
    End With
Next i
End Sub



Private Sub Command7_Click()
PB2.max = 1500
PB2.value = 1
UserPos.X = 12
UserPos.Y = 10
picture1.Cls
pGuardar.Cls

GrabarParcial = False
Grabar = True
End Sub

Private Sub Command8_Click()
Call mnuQuitarFunciones_Click
CopyState = 1
CopyRect.Left = 0
CopyRect.Top = 0
CopyRect.Right = 0
CopyRect.Bottom = 0
AddMX = 0
AddMY = 0
End Sub

Private Sub Command9_Click()
CopyState = 4
End Sub

Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarTrans.value = (cUnionManual.value = True)
Call cInsertarTrans_Click
End Sub

Private Sub cverBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerBloqueos.Checked = cVerBloqueos.value
End Sub

Private Sub cverTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub cNumFunc_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

If KeyAscii = 13 Then
    Dim Cont As String
    Cont = frmMain.cNumFunc(Index).text
    Call cNumFunc_LostFocus(Index)
    If Cont <> frmMain.cNumFunc(Index).text Then Exit Sub
    If frmMain.cNumFunc(Index).ListCount > 5 Then
        frmMain.cNumFunc(Index).RemoveItem 0
    End If
    frmMain.cNumFunc(Index).AddItem frmMain.cNumFunc(Index).text
    Exit Sub
ElseIf KeyAscii = 8 Then
    
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub cNumFunc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If cNumFunc(Index).text = vbNullString Then
    frmMain.cNumFunc(Index).text = IIf(Index = 1, 500, 1)
End If
End Sub

Private Sub cNumFunc_LostFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Index = 0 Then
        If frmMain.cNumFunc(Index).text > 499 Or frmMain.cNumFunc(Index).text < 1 Then
            frmMain.cNumFunc(Index).text = 1
        End If
    ElseIf Index = 1 Then
        If frmMain.cNumFunc(Index).text < 500 Or frmMain.cNumFunc(Index).text > 32000 Then
            frmMain.cNumFunc(Index).text = 500
        End If
    ElseIf Index = 2 Then
        If frmMain.cNumFunc(Index).text < 1 Or frmMain.cNumFunc(Index).text > 32000 Then
            frmMain.cNumFunc(Index).text = 1
        End If
    End If
End Sub

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cInsertarBloqueo.value = True Then
    cQuitarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cQuitarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cQuitarBloqueo.value = True Then
    cInsertarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cInsertarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
QuitarLuz
If cQuitarEnEstaCapa.value = True Then
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnTodasLasCapas.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnTodasLasCapas.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
QuitarLuz
If cQuitarEnTodasLasCapas.value = True Then
    cCapas.Enabled = False
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cCapas.Enabled = True
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub


Private Sub cQuitarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarFunc(Index).value = True Then
    cInsertarFunc(Index).Enabled = False
    cAgregarFuncalAzar(Index).Enabled = False
    cCantFunc(Index).Enabled = False
    cNumFunc(Index).Enabled = False
    cFiltro((Index) + 1).Enabled = False
    lListado((Index) + 1).Enabled = False
    Call modPaneles.EstSelectPanel((Index) + 3, True)
Else
    cInsertarFunc(Index).Enabled = True
    cAgregarFuncalAzar(Index).Enabled = True
    cCantFunc(Index).Enabled = True
    cNumFunc(Index).Enabled = True
    cFiltro((Index) + 1).Enabled = True
    lListado((Index) + 1).Enabled = True
    Call modPaneles.EstSelectPanel((Index) + 3, False)
End If
End Sub

Private Sub cQuitarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrans.value = True Then
    cInsertarTransOBJ.Enabled = False
    cInsertarTrans.Enabled = False
    cUnionManual.Enabled = False
    cUnionAuto.Enabled = False
    tTMapa.Enabled = False
    tTX.Enabled = False
    tTY.Enabled = False
    mnuInsertarTransladosAdyasentes.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    tTMapa.Enabled = True
    tTX.Enabled = True
    tTY.Enabled = True
    cUnionAuto.Enabled = True
    cUnionManual.Enabled = True
    cInsertarTrans.Enabled = True
    cInsertarTransOBJ.Enabled = True
    mnuInsertarTransladosAdyasentes.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub

Private Sub cQuitarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrigger.value = True Then
    lListado(4).Enabled = False
    cInsertarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    lListado(4).Enabled = True
    cInsertarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cSeleccionarSuperficie.value = True Then
    cQuitarEnTodasLasCapas.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cQuitarEnTodasLasCapas.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cUnionAuto_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub Form_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.SetFocus

End Sub

Private Sub Form_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
' HotKeys
If HotKeysAllow = False Then Exit Sub

Select Case UCase(Chr(KeyAscii))
    Case "S" ' Activa/Desactiva Insertar Superficie
        cSeleccionarSuperficie.value = (cSeleccionarSuperficie.value = False)
        Call cSeleccionarSuperficie_Click
    Case "T" ' Activa/Desactiva Insertar Translados
        cInsertarTrans.value = (cInsertarTrans.value = False)
        Call cInsertarTrans_Click
    Case "B" ' Activa/Desactiva Insertar Bloqueos
        cInsertarBloqueo.value = (cInsertarBloqueo.value = False)
        Call cInsertarBloqueo_Click
    Case "N" ' Activa/Desactiva Insertar NPCs
        cInsertarFunc(0).value = (cInsertarFunc(0).value = False)
        Call cInsertarFunc_Click(0)
   ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
   '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
   '     Call cInsertarFunc_Click(1)
    Case "O" ' Activa/Desactiva Insertar Objetos
        cInsertarFunc(2).value = (cInsertarFunc(2).value = False)
        Call cInsertarFunc_Click(2)
    Case "G" ' Activa/Desactiva Insertar Triggers
        cInsertarTrigger.value = (cInsertarTrigger.value = False)
        Call cInsertarTrigger_Click
    Case "Q" ' Quitar Funciones
        Call mnuQuitarFunciones_Click
End Select
End Sub




Private Sub Form_Load()
EditLimit.Left = 1500
EditLimit.Top = 1100
EditLimit.Right = 0
EditLimit.Bottom = 0
optDia_Click
Command34_Click
Option1_Click
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Seleccionando Then CopiarSeleccion
End Sub


Private Sub lListado_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
On Error Resume Next
PuedeMover = False
If HotKeysAllow = False Then
    lListado(Index).Tag = lListado(Index).ListIndex
    Select Case Index
        Case 0
            cGrh.text = DameGrhIndex(ReadField(2, lListado(Index).text, Asc("#")))
            If SupData(ReadField(2, lListado(Index).text, Asc("#"))).Capa <> 0 Then
                If LenB(ReadField(2, lListado(Index).text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.text
                cCapas.text = SupData(ReadField(2, lListado(Index).text, Asc("#"))).Capa
            Else
                If LenB(cCapas.Tag) <> 0 Then
                    cCapas.text = cCapas.Tag
                    cCapas.Tag = vbNullString
                End If
            End If
            If SupData(ReadField(2, lListado(Index).text, Asc("#"))).Block = True Then
                If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                cInsertarBloqueo.value = True
                Call cInsertarBloqueo_Click
            Else
                If LenB(cInsertarBloqueo.Tag) <> 0 Then
                    cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                    cInsertarBloqueo.Tag = vbNullString
                    Call cInsertarBloqueo_Click
                End If
            End If
            Call fPreviewGrh(cGrh.text)
            Call modPaneles.VistaPreviaDeSup
        Case 1
            cNumFunc(0).text = ReadField(2, lListado(Index).text, Asc("#"))
            tNPC.text = cNumFunc(0).text
        Case 2
            cNumFunc(1).text = ReadField(2, lListado(Index).text, Asc("#"))
            
        Case 3
            frmConfigSup.mLargo.text = 1
            frmConfigSup.mAncho.text = 1
            cNumFunc(2).text = ReadField(2, lListado(Index).text, Asc("#"))
            fPreviewGrh (ObjData(Val(ReadField(2, lListado(Index).text, Asc("#")))).grhindex)
            Call modPaneles.VistaPreviaDeSup
    End Select
Else
    lListado(Index).ListIndex = lListado(Index).Tag
End If

End Sub

Private Sub lListado_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If Index = 3 And Button = 2 Then
    If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc
End If
End Sub

Private Sub lListado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
On Error Resume Next
HotKeysAllow = False
End Sub



Private Sub lstNpc_Click()
tNPC.text = Areas(SelArea).NPC(lstNpc.ListIndex + 1).NPCIndex
tCant.text = Areas(SelArea).NPC(lstNpc.ListIndex + 1).Cantidad
End Sub

Private Sub mmGuardarCliente_Click()
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim ByFlags As Byte
PB.max = 1500
PB.min = 0
If Not MapaCargado Then Exit Sub
Open (App.Path & "\Mapa" & UserMap & ".AO") For Binary As #5
    For Y = 1 To 1500
        For X = 1 To 1100
            If MapData(X, Y).Blocked = 1 Then
                Put #5, , CByte(50 Xor ((X Mod 200) + 55))
            Else
                Put #5, , CByte(MapData(X, Y).Trigger Xor ((X Mod 200) + 55))
            End If
            Put #5, , (MapData(X, Y).Graphic(1).grhindex + X) Xor (Y + 301) Xor (X + 721)
            Put #5, , (MapData(X, Y).Graphic(2).grhindex + X) Xor (Y + 501) Xor (X + 529)
            Put #5, , (MapData(X, Y).Graphic(3).grhindex + X) Xor (X + 239) Xor (Y + 319)
            Put #5, , (MapData(X, Y).Graphic(4).grhindex + X) Xor (X + 671) Xor (Y + 129)
            Put #5, , CByte(MapData(X, Y).LuzRango)
        Next X
        PB.value = Y
    Next Y
Close #5
End Sub

Private Sub mmTrans_Click()
VerTrans = (VerTrans = False)
mmTrans.Checked = VerTrans
End Sub

Private Sub mnuActualizarCabezas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.CargarCabezas
End Sub

Private Sub mnuActualizarCuerpos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.CargarCuerpos
End Sub

Private Sub mnuActualizarGraficos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modDX8FIFO.LoadGrhData
End Sub

Private Sub mnuActualizarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesSuperficie
End Sub

Private Sub mnuAbrirMapa_Click()
    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
    'Call modMapIO.NuevoMapa
    Dim Body As Integer
    Dim Head As Integer
    Dim loopc As Integer
    Dim ByFlags As Integer
    Dim Rango As Byte
    Dim X As Long
    Dim Y As Long
    Dim XX As Integer
    Dim YY As Integer
    Dim i As Integer
    Dim Heading As Byte
    Dim Templng As Long
    Dim tmpInt As Integer
    Open (App.Path & "\SMapa" & UserMap & ".map") For Binary As #155
    'Templng = 10 * (X - 1 + CamX) + (Y - 1 + CamY) * 840 + 1
    PB.max = 1500
    PB.min = 0
    For Y = 1 To 1500
        For X = 1 To 1100
            'Templng = (X - 1) * 10 + (Y - 1) * 11000 + 1
            Get #155, , ByFlags
            MapData(X, Y).Blocked = (ByFlags And 1)
            'Templng = Templng + 1
            For loopc = 1 To 4
                If loopc = 1 Or (ByFlags And 2 ^ (loopc - 1)) Then
                    Get #155, , tmpInt
                    MapData(X, Y).Graphic(loopc).grhindex = tmpInt
                    'Set up GRH
                    If MapData(X, Y).Graphic(loopc).grhindex > 0 Then
                        InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).grhindex
                    ElseIf MapData(X, Y).Graphic(loopc).grhindex < 0 And loopc <> 3 Then
                        MapData(X, Y).Graphic(loopc).grhindex = 0
                    End If
                Else
                    MapData(X, Y).Graphic(loopc).grhindex = 0
                End If
                 'Templng = Templng + 2
            Next loopc
            If ByFlags And 16 Then
                Get #155, , Heading
                 MapData(X, Y).Trigger = Heading
                 'Templng = Templng + 1
                 If MapData(X, Y).Trigger = 50 Then MapData(X, Y).Trigger = 0
            Else
                MapData(X, Y).Trigger = 0
            End If
            MapData(X, Y).NpcArea = 0
            If ByFlags And 32 Then
                Get #155, , MapData(X, Y).NPCIndex
                If MapData(X, Y).NPCIndex <= 0 Then
                    MapData(X, Y).NPCIndex = 0
                    If MapData(X, Y).CharIndex > 0 Then
                        EraseChar (MapData(X, Y).CharIndex)
                        MapData(X, Y).CharIndex = 0
                    End If
                Else
                    Body = NpcData(MapData(X, Y).NPCIndex).Body
                    Head = NpcData(MapData(X, Y).NPCIndex).Head
                    Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                End If
            Else
                MapData(X, Y).NPCIndex = 0
                    If MapData(X, Y).CharIndex > 0 Then
                        EraseChar (MapData(X, Y).CharIndex)
                        MapData(X, Y).CharIndex = 0
                    End If
            End If
            If ByFlags And 64 Then
                Get #155, , MapData(X, Y).OBJInfo.OBJIndex
                Get #155, , MapData(X, Y).OBJInfo.Amount
                If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
                
                    If EsObjetoFijo(ObjData(MapData(X, Y).OBJInfo.OBJIndex).OBJType) Then
                        If ObjData(MapData(X, Y).OBJInfo.OBJIndex).OBJType = eOBJType.otFlores Then
                            InitGrh MapData(X, Y).Graphic(2), ObjData(MapData(X, Y).OBJInfo.OBJIndex).grhindex
                        Else
                        
                            InitGrh MapData(X, Y).Graphic(3), ObjData(MapData(X, Y).OBJInfo.OBJIndex).grhindex
                        End If
                    End If
                
                    InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.OBJIndex).grhindex
                End If
            Else
                MapData(X, Y).OBJInfo.OBJIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
                MapData(X, Y).ObjGrh.grhindex = 0
            End If
            If ByFlags And 128 Then
                Get #155, , MapData(X, Y).TileExit.map
                Get #155, , MapData(X, Y).TileExit.X
                Get #155, , MapData(X, Y).TileExit.Y
            Else
                MapData(X, Y).TileExit.map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
            End If
            MapData(X, Y).Hora = 2
            If MapData(X, Y).Graphic(3).grhindex < 0 Then
                Get #155, , Rango
                MapData(X, Y).LuzRango = Rango
                MapData(X, Y).LightIndex = Light_Create(X, Y, 255, 255, 255, Rango, -MapData(X, Y).Graphic(3).grhindex - 1)
            End If
            
        Next X
        PB.value = Y
    Next Y
    Close #155
    'modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    EngineRun = True
    MapaCargado = True
    CargarAreas
    
Exit Sub
ErrHandler:
End Sub
Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otArbolElfico Or _
               OBJType = eOBJType.otFlores Or _
               OBJType = eOBJType.otYacimiento

End Function
Private Sub mnuacercade_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAbout.Show
End Sub



Private Sub mnuActualizarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesNPC
End Sub

Private Sub mnuActualizarObjs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesOBJ
End Sub

Private Sub mnuActualizarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesTriggers
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAutoGuardarMapa.Show
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 2 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next

modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuBloquearS_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call BlockearSeleccion
End Sub

Private Sub mnuConfigAvanzada_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmConfigSup.Show
End Sub

Private Sub mnuConfigObjTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Cfg_TrOBJ = cNumFunc(2).text
End Sub

Private Sub mnuCopiar_Click()
On Error Resume Next
Dim X As Integer
Dim Y As Integer
ReDim MapCopy(CopyRect.Right - CopyRect.Left, CopyRect.Bottom - CopyRect.Top)
'ReDim MapCopyD(CopyRect.Right - CopyRect.Left, CopyRect.Bottom - CopyRect.Top)
For X = CopyRect.Left To CopyRect.Right
    For Y = CopyRect.Top To CopyRect.Bottom
        MapCopy(X - CopyRect.Left, Y - CopyRect.Top) = MapData(X, Y)
    Next Y
Next X


'Call CopiarSeleccion
End Sub

Private Sub mnuCortar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call CortarSeleccion
End Sub


Private Sub mnuDeshacer_Click()
Call Deshacer
End Sub

Private Sub mnuDeshacerPegado_Click()

End Sub

Private Sub mnuGRHaBMP_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
frmGRHaBMP.Show
End Sub

Private Sub mnuGuardarcomoBMP_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
        RenderToPicture Ratio, True
    End If
End Sub

Private Sub mnuGuardarcomoJPG_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
        engine.RenderToPicture
    End If
End Sub

Private Sub mnuGuardarMapa_Click()
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim ByFlags As Integer
PB.max = 1500
PB.min = 0
If Not MapaCargado Then Exit Sub

Name App.Path & "\SMapa" & UserMap & ".map" As App.Path & "\Backups Mapas\SMapa" & UserMap & " - " & format(Now, "dd-mm-YYYY hh.MM.ss") & ".map"

Open (App.Path & "\SMapa" & UserMap & ".map") For Binary As #5
    For Y = 1 To 1500
        For X = 1 To 1100
            ByFlags = 0
            If MapData(X, Y).Trigger > 50 Then MapData(X, Y).Trigger = 0
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).grhindex > 0 Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).grhindex <> 0 Then ByFlags = ByFlags Or 4 'Si es negativo es una luz
            If MapData(X, Y).Graphic(4).grhindex > 0 Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Trigger > 0 Then ByFlags = ByFlags Or 16
            If MapData(X, Y).NPCIndex > 0 Then ByFlags = ByFlags Or 32
            If MapData(X, Y).OBJInfo.OBJIndex > 0 Then ByFlags = ByFlags Or 64
            If MapData(X, Y).TileExit.map > 0 Then ByFlags = ByFlags Or 128
            If MapData(X, Y).particle_group_index > 0 Then ByFlags = ByFlags Or 256
            'If MapData(X, Y).OBJInfo.objindex > 0 Then
                'If ObjData(MapData(X, Y).OBJInfo.objindex).ObjType = 4 Then ByFlags = ByFlags Or 128
                'If ObjData(MapData(X, Y).OBJInfo.objindex).ObjType = 22 Then ByFlags = ByFlags Or 256
            'End If
            Put #5, , ByFlags
            Put #5, , CInt(MapData(X, Y).Graphic(1).grhindex)
            If MapData(X, Y).Graphic(2).grhindex Then Put #5, , CInt(MapData(X, Y).Graphic(2).grhindex)
            If MapData(X, Y).Graphic(3).grhindex <> 0 Then Put #5, , CInt(MapData(X, Y).Graphic(3).grhindex)
            If MapData(X, Y).Graphic(4).grhindex Then Put #5, , CInt(MapData(X, Y).Graphic(4).grhindex)
            If MapData(X, Y).Trigger Then Put #5, , CByte(MapData(X, Y).Trigger)
            If MapData(X, Y).NPCIndex > 0 Then Put #5, , MapData(X, Y).NPCIndex
            If MapData(X, Y).OBJInfo.OBJIndex > 0 Then Put #5, , MapData(X, Y).OBJInfo.OBJIndex
            If MapData(X, Y).OBJInfo.OBJIndex > 0 Then Put #5, , MapData(X, Y).OBJInfo.Amount
            If MapData(X, Y).TileExit.map > 0 Then
                Put #5, , MapData(X, Y).TileExit.map
                Put #5, , MapData(X, Y).TileExit.X
                Put #5, , MapData(X, Y).TileExit.Y
            End If
            If MapData(X, Y).particle_group_index > 0 Then
                Put #5, , CByte(0)
            End If
            If MapData(X, Y).Graphic(3).grhindex < 0 Then
                Put #5, , MapData(X, Y).LuzRango
                
            End If
        Next X
        PB.value = Y
    Next Y
Close #5
Exit Sub
Open (App.Path & "\SuperMapv3Cliente.map") For Binary As #5
    For Y = 1 To 1500
        For X = 1 To 1100
            If MapData(X, Y).Blocked = 1 Then MapData(X, Y).Trigger = 50
            Put #5, , CByte(MapData(X, Y).Trigger)
            Put #5, , MapData(X, Y).Graphic(1).grhindex
            Put #5, , MapData(X, Y).Graphic(2).grhindex
            Put #5, , MapData(X, Y).Graphic(3).grhindex
            Put #5, , MapData(X, Y).Graphic(4).grhindex
        Next X
        PB.value = Y
    Next Y
Close #5

End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa
End Sub

Private Sub mnuGuardarUltimaConfig_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 23/05/06
'*************************************************
mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub mnuInformes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmInformes.Show
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub mnuManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) <> 0 Then
    Call Shell("explorer " & App.Path & "\manual\index.html")
    DoEvents
End If
End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 3 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 3, True
End Sub



'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

DeseaGuardarMapa Dialog.FileName


frmMain.Dialog.FileName = Empty

If WalkMode = True Then
    Call modGeneral.ToggleWalkMode
End If

Call modMapIO.NuevoMapa

Call cmdInformacionDelMapa_Click

End Sub

Private Sub mnuObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 5 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 5, True
End Sub


Private Sub mnuOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
frmOptimizar.Show
End Sub

Private Sub mnuParticle_Click()
    mnuParticle.Checked = Not mnuParticle.Checked
End Sub

Private Sub mnuPegar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
CopyState = 4
'Call PegarSeleccion
End Sub

Private Sub mnuQBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 3, False
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 5, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 6, False
End Sub


Private Sub mnuQuitarBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(0)
End Sub

Private Sub mnuQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Superficies
cSeleccionarSuperficie.value = False
Call cSeleccionarSuperficie_Click
cQuitarEnEstaCapa.value = False
Call cQuitarEnEstaCapa_Click
cQuitarEnTodasLasCapas.value = False
Call cQuitarEnTodasLasCapas_Click
' Translados
cQuitarTrans.value = False
Call cQuitarTrans_Click
cInsertarTrans.value = False
Call cInsertarTrans_Click
' Bloqueos
cQuitarBloqueo.value = False
Call cQuitarBloqueo_Click
cInsertarBloqueo.value = False
Call cInsertarBloqueo_Click
' Otras funciones
cInsertarFunc(0).value = False
Call cInsertarFunc_Click(0)
cInsertarFunc(1).value = False
Call cInsertarFunc_Click(1)
cInsertarFunc(2).value = False
Call cInsertarFunc_Click(2)
cQuitarFunc(0).value = False
Call cQuitarFunc_Click(0)
cQuitarFunc(1).value = False
Call cQuitarFunc_Click(1)
cQuitarFunc(2).value = False
Call cQuitarFunc_Click(2)
' Triggers
cInsertarTrigger.value = False
Call cInsertarTrigger_Click
cQuitarTrigger.value = False
Call cQuitarTrigger_Click

QuitarLuz
End Sub

Private Sub mnuQuitarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_NPCs(False)
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Objetos
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Capa(cCapas.text)
End Sub

Private Sub mnuQuitarTODO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Borrar_Mapa
End Sub

Private Sub mnuQuitarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Call modEdicion.Quitar_Translados
End Sub

Private Sub mnuQuitarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Triggers
End Sub

Private Sub mnuReAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error GoTo ErrHandler
    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName
        End If
    End If
    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapa Dialog.FileName
    DoEvents
    'mnuReAbrirMapa.Enabled = True
    EngineRun = True
Exit Sub
ErrHandler:
End Sub

Private Sub mnuRealizarOperacion_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call AccionSeleccion
End Sub

Private Sub mnuSalir_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub mnuSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 0 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 1 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 6 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 6, True
End Sub


Private Sub mnuVerAutomatico_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
End Sub

Private Sub mnuVerBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cVerBloqueos.value = (cVerBloqueos.value = False)
mnuVerBloqueos.Checked = cVerBloqueos.value

End Sub

Private Sub mnuVerCapa1_Click()
mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
End Sub

Private Sub mnuVerCapa2_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
End Sub

Private Sub mnuVerCapa3_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
End Sub

Private Sub mnuVerCapa4_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
End Sub


Private Sub mnuVerGrilla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
VerGrilla = (VerGrilla = False)
mnuVerGrilla.Checked = VerGrilla
End Sub

Private Sub mnuVerNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

End Sub

Private Sub mnuVerTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cVerTriggers.value = (cVerTriggers.value = False)
mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub optDia_Click()
IluRGB.r = 240
IluRGB.g = 240
IluRGB.b = 240
IluRGB.A = 255
ColorMapa = D3DColorRGBA(240, 240, 240, 255)
Hora = 0
End Sub

Private Sub Option1_Click()
tp1(0).Tag = "357"
tp1(1).Tag = "366"
tp1(2).Tag = "359"
tp1(3).Tag = "362"
tp1(4).Tag = "137"
tp1(5).Tag = "365"
tp1(6).Tag = "356"
tp1(7).Tag = "367"
tp1(8).Tag = "358"

tp2(0).Tag = "364"
tp2(1).Tag = "367"
tp2(2).Tag = "361"
tp2(3).Tag = "365"
tp2(4).Tag = "0"
tp2(5).Tag = "362"
tp2(6).Tag = "363"
tp2(7).Tag = "366"
tp2(8).Tag = "360"
End Sub

Private Sub Option2_Click()
tp1(0).Tag = "369"
tp1(1).Tag = "374"
tp1(2).Tag = "371"
tp1(3).Tag = "375"
tp1(4).Tag = "137"
tp1(5).Tag = "378"
tp1(6).Tag = "368"
tp1(7).Tag = "379"
tp1(8).Tag = "370"

tp2(0).Tag = "377"
tp2(1).Tag = "379"
tp2(2).Tag = "373"
tp2(3).Tag = "378"
tp2(4).Tag = "130"
tp2(5).Tag = "375"
tp2(6).Tag = "376"
tp2(7).Tag = "374"
tp2(8).Tag = "372"
End Sub

Private Sub Option3_Click()
tp1(0).Tag = "658"
tp1(1).Tag = "667"
tp1(2).Tag = "660"
tp1(3).Tag = "663"
tp1(4).Tag = "137"
tp1(5).Tag = "666"
tp1(6).Tag = "657"
tp1(7).Tag = "668"
tp1(8).Tag = "659"

tp2(0).Tag = "665"
tp2(1).Tag = "668"
tp2(2).Tag = "662"
tp2(3).Tag = "666"
tp2(4).Tag = "621"
tp2(5).Tag = "663"
tp2(6).Tag = "664"
tp2(7).Tag = "667"
tp2(8).Tag = "661"
End Sub

Private Sub optNoche_Click()
IluRGB.r = 150
IluRGB.g = 150
IluRGB.b = 150
IluRGB.A = 255
ColorMapa = D3DColorRGBA(150, 150, 150, 255)
Hora = 1
End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
UserPos.X = X * 5
UserPos.Y = Y * 5
tpX.text = UserPos.X
tpY.text = UserPos.Y
bRefreshRadar = True
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
MiRadarX = X
MiRadarY = Y
If Button = 1 Then
    UserPos.X = X * 5
    UserPos.Y = Y * 5
    tpX.text = UserPos.X
    tpY.text = UserPos.Y
    ApuntadorRadar.Move UserPos.X / 5, UserPos.Y / 5
End If
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

End Sub

Public Sub SelectArea()
    lblNum.Caption = "Nº de Area: " & SelArea
    lstNpc.Clear
    lblPos.Caption = ""
    If SelArea > 0 Then
        lblPos.Caption = "(" & Areas(SelArea).X1 & "," & Areas(SelArea).Y1 & "," & Areas(SelArea).X2 & "," & Areas(SelArea).Y2
        Dim i As Integer
        Dim e As Integer
        Dim Nombre As String
        For i = 1 To Areas(SelArea).Npcs

            Nombre = NombreNPC(Areas(SelArea).NPC(i).NPCIndex)
            lstNpc.AddItem "(" & Areas(SelArea).NPC(i).Cantidad & ") " & Nombre
        Next i
        tZX1.text = Areas(SelArea).X1
        tZY1.text = Areas(SelArea).Y1
        tZX2.text = Areas(SelArea).X2
        tZY2.text = Areas(SelArea).Y2
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************

' Guardar configuración
WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")
If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    WriteVar IniPath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)
End If

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub



Private Sub PreviewGrh_Click()
PuedeMover = True
End Sub

Private Sub Renderer_Click()
    PuedeMover = True
    Me.SetFocus
End Sub

Private Sub Renderer_DblClick()
Dim tx As Integer
Dim ty As Integer

If Not MapaCargado Then Exit Sub

If SobreX > 0 And SobreY > 0 Then
    DobleClick Val(SobreX), Val(SobreY)
End If
End Sub

Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tx As Integer
Dim ty As Integer

If Not MapaCargado Then Exit Sub

If X < 0 Or Y < 0 Then Exit Sub

ConvertCPtoTP X, Y, tx, ty

If CopyState > 0 And Button = 2 Then
    CopyState = 0
End If

If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionIX = tx '+ UserPos.X
    SeleccionIY = ty '+ UserPos.Y
Else
    If chkCostas.value = vbChecked Then
        Call HacerCostas(Button, tx, ty)
    Else
        ClickEdit Button, tx, ty
    End If
End If

Dim i As Integer
Dim e As Integer
If AgregarZona = 0 Then
For i = 1 To NumZonas
    If Zonas(i).Mapa = UserMap And tx >= Zonas(i).X1 And tx <= Zonas(i).X2 And ty >= Zonas(i).Y1 And ty <= Zonas(i).Y2 Then
        SelZona = i
        Text6.text = Zonas(i).Nombre
        Check4.value = IIf(Zonas(i).Segura, vbChecked, vbUnchecked)
        chkAcoplar.value = IIf(Zonas(i).Acoplar, vbChecked, vbUnchecked)
        For e = 1 To 5
            tMusica(e - 1).text = Zonas(i).Musica(e)
        Next e
        tZX1.text = Zonas(i).X1
        tZY1.text = Zonas(i).Y1
        tZX2.text = Zonas(i).X2
        tZY2.text = Zonas(i).Y2
        tNumZona.Caption = "N° Zona: " & i
        If Zonas(i).Acoplar = 0 Then Exit For
    End If
Next i
End If
If AgregarArea = 5 Then
    AgregarArea = 0
Else
If AgregarArea = 0 Then
For i = 1 To NumAreas
    If Areas(i).Mapa = UserMap And tx >= Areas(i).X1 And tx <= Areas(i).X2 And ty >= Areas(i).Y1 And ty <= Areas(i).Y2 Then
        SelArea = i
        SelectArea
        Exit Sub
    End If
Next i
End If




SelArea = 0
SelectArea
End If
End Sub

Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tx As Integer
Dim ty As Integer
Dim XX As Integer
Dim YY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub

If X < 0 Or Y < 0 Then Exit Sub
'Make sure click is in view window
'If X <= Renderer.Left Or X >= Renderer.Left + MainViewWidth Or Y <= Renderer.Top Or Y >= Renderer.Top + MainViewHeight Then
'    Exit Sub
'End If

ConvertCPtoTP X, Y, tx, ty

POSX.Caption = "X: " & tx & " - Y: " & ty
If tx < 10 Or ty < 10 Or tx > 90 Or ty > 90 Then
    POSX.ForeColor = vbRed
Else
    POSX.ForeColor = vbWhite
End If
PRect.Right = tx
PRect.Bottom = ty
If CopyState = 2 Then
            CopyRect.Right = tx
            CopyRect.Bottom = ty
End If
If AgregarZona = 2 Then
            ZonaR.Right = tx
            ZonaR.Bottom = ty
End If
If AgregarArea = 2 Then
            AreaR.Right = tx
            AreaR.Bottom = ty
End If
 If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionFX = tx '+ TileX
    SeleccionFY = ty '+ TileY
Else

If chkCostas.value = vbChecked Then
    Call MostrarCostas(tx, ty)
    Call HacerCostas(Button, tx, ty)
Else
    ClickEdit Button, tx, ty
End If



   
End If
    MouseX = X
    MouseY = Y
    MX = tx
MY = ty







If PRect.Left = 0 And (tx <> LastX Or ty <> LastY) Then
    LastX = 0
    LastY = 0
End If
End Sub
Sub HacerCostas(ByVal Button As Integer, ByVal tx As Integer, ByVal ty As Integer)
    Dim gh(0 To 8) As Integer
    Dim i As Integer
    Dim YY As Integer
    Dim XX As Integer
    
If Button = 1 Then
    MapInfo.Changed = 1
    For YY = ty - 1 To ty + 1
        For XX = tx - 1 To tx + 1
            gh(i) = MapData(XX, YY).Graphic(2).grhindex
            i = i + 1
        Next XX
    Next YY
    If gh(4) > 0 Then
        Exit Sub
    ElseIf gh(1) = DameGrhIndex(tp1(1).Tag) + 1 Then 'Click abajo
        tp1_Click (2)
        ClickEdit 1, tx, ty - 1
        ClickEdit 1, tx - 1, ty - 1
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(1).Tag) + 0 Then
        tp1_Click (0)
        ClickEdit 1, tx, ty - 1
        ClickEdit 1, tx + 1, ty - 1
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(2).Tag) + 3 Then
        tp1_Click (5)
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(0).Tag) + 2 Then
        tp1_Click (3)
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(3).Tag) Or gh(1) = DameGrhIndex(tp1(3).Tag) + 2 Then
        tp1_Click (3)
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(5).Tag) + 1 Or gh(1) = DameGrhIndex(tp1(5).Tag) + 3 Then
        tp1_Click (5)
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(7).Tag) + 3 Then
        tp2_Click (0)
        ClickEdit 1, tx, ty - 1
        tp2_Click (3)
        ClickEdit 1, tx, ty
    ElseIf gh(1) = DameGrhIndex(tp1(7).Tag) + 2 Then
        tp2_Click (2)
        ClickEdit 1, tx, ty - 1
        tp2_Click (5)
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(1).Tag) + 0 Then 'Click arriba
        tp2_Click (5)
        ClickEdit 1, tx, ty
        tp2_Click (8)
        ClickEdit 1, tx, ty + 1
    ElseIf gh(7) = DameGrhIndex(tp1(1).Tag) + 1 Then
        tp2_Click (3)
        ClickEdit 1, tx, ty
        tp2_Click (6)
        ClickEdit 1, tx, ty + 1
    ElseIf gh(7) = DameGrhIndex(tp1(6).Tag) + 0 Then
        tp1_Click (3)
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(8).Tag) + 1 Then
        tp1_Click (5)
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(3).Tag) Or gh(7) = DameGrhIndex(tp1(3).Tag) + 2 Then
        tp1_Click (3)
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(5).Tag) + 1 Or gh(7) = DameGrhIndex(tp1(5).Tag) + 3 Then
        tp1_Click (5)
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(7).Tag) + 3 Then
        tp1_Click (8)
        ClickEdit 1, tx, ty + 1
        ClickEdit 1, tx - 1, ty + 1
        ClickEdit 1, tx, ty
    ElseIf gh(7) = DameGrhIndex(tp1(7).Tag) + 2 Then
        tp1_Click (6)
        ClickEdit 1, tx, ty + 1
        ClickEdit 1, tx + 1, ty + 1
        ClickEdit 1, tx, ty
    ElseIf gh(5) = DameGrhIndex(tp1(3).Tag) + 2 Then 'Click izquierda
        tp2_Click (1)
        ClickEdit 1, tx, ty
        tp2_Click (2)
        ClickEdit 1, tx + 1, ty
    ElseIf gh(5) = DameGrhIndex(tp1(3).Tag) + 0 Then
        tp2_Click (7)
        ClickEdit 1, tx, ty
        tp2_Click (8)
        ClickEdit 1, tx + 1, ty
    ElseIf gh(5) = DameGrhIndex(tp1(5).Tag) + 1 Then
        tp1_Click (2)
        ClickEdit 1, tx, ty
        ClickEdit 1, tx + 1, ty
        ClickEdit 1, tx + 1, ty + 1
    ElseIf gh(5) = DameGrhIndex(tp1(5).Tag) + 3 Then
        tp1_Click (8)
        ClickEdit 1, tx, ty
        ClickEdit 1, tx + 1, ty
        ClickEdit 1, tx + 1, ty - 1
    ElseIf gh(5) = DameGrhIndex(tp1(1).Tag) Or gh(5) = DameGrhIndex(tp1(1).Tag) + 1 Then
        tp1_Click (1)
        ClickEdit 1, tx, ty
    ElseIf gh(5) = DameGrhIndex(tp1(7).Tag) + 2 Or gh(5) = DameGrhIndex(tp1(7).Tag) + 3 Then
        tp1_Click (7)
        ClickEdit 1, tx, ty
    ElseIf gh(5) = DameGrhIndex(tp1(2).Tag) + 0 Then
        tp1_Click (1)
        ClickEdit 1, tx, ty
    ElseIf gh(5) = DameGrhIndex(tp1(8).Tag) + 2 Then
        tp1_Click (7)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(3).Tag) + 0 Then 'Click derecha
        tp1_Click (0)
        ClickEdit 1, tx, ty
        ClickEdit 1, tx - 1, ty
        ClickEdit 1, tx - 1, ty + 1
    ElseIf gh(3) = DameGrhIndex(tp1(3).Tag) + 2 Then
        tp1_Click (6)
        ClickEdit 1, tx, ty
        ClickEdit 1, tx - 1, ty
        ClickEdit 1, tx - 1, ty - 1
    ElseIf gh(3) = DameGrhIndex(tp1(5).Tag) + 3 Then
        tp2_Click (0)
        ClickEdit 1, tx - 1, ty
        tp2_Click (1)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(5).Tag) + 1 Then
        tp2_Click (6)
        ClickEdit 1, tx - 1, ty
        tp2_Click (7)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(1).Tag) Or gh(3) = DameGrhIndex(tp1(1).Tag) + 1 Then
        tp1_Click (1)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(7).Tag) + 2 Or gh(3) = DameGrhIndex(tp1(7).Tag) + 3 Then
        tp1_Click (7)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(0).Tag) + 1 Then
        tp1_Click (1)
        ClickEdit 1, tx, ty
    ElseIf gh(3) = DameGrhIndex(tp1(6).Tag) + 3 Then
        tp1_Click (7)
        ClickEdit 1, tx, ty
    End If
ElseIf Button = 2 Then
    MapInfo.Changed = 1
    MapData(tx, ty).Graphic(2).grhindex = 0
End If
End Sub

Sub MostrarCostas(tx As Integer, ty As Integer)
    Dim gh(0 To 8) As Integer
    Dim i As Integer
    Dim YY As Integer
    Dim XX As Integer
    For YY = ty - 1 To ty + 1
        For XX = tx - 1 To tx + 1
            If InMapBounds(XX, YY) Then
            gh(i) = MapData(XX, YY).Graphic(2).grhindex
            End If
            i = i + 1
        Next XX
    Next YY
    Call MCosta(0, 0, 0, 0)
    If gh(4) > 0 Then
        Exit Sub
    ElseIf gh(1) = DameGrhIndex(tp1(1).Tag) + 1 Then 'Click abajo
        Call MCosta(1, tp1(2).Tag, tx, ty - 1)
        Call MCosta(2, tp1(2).Tag, tx - 1, ty - 1)
        Call MCosta(3, tp1(2).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(1).Tag) + 0 Then
        Call MCosta(1, tp1(0).Tag, tx, ty - 1)
        Call MCosta(2, tp1(0).Tag, tx + 1, ty - 1)
        Call MCosta(3, tp1(0).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(2).Tag) + 3 Then
        Call MCosta(1, tp1(5).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(0).Tag) + 2 Then
        Call MCosta(1, tp1(3).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(3).Tag) Or gh(1) = DameGrhIndex(tp1(3).Tag) + 2 Then
        Call MCosta(1, tp1(3).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(5).Tag) + 1 Or gh(1) = DameGrhIndex(tp1(5).Tag) + 3 Then
        Call MCosta(1, tp1(5).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(7).Tag) + 3 Then
        Call MCosta(1, tp2(0).Tag, tx, ty - 1)
        Call MCosta(2, tp2(3).Tag, tx, ty)
    ElseIf gh(1) = DameGrhIndex(tp1(7).Tag) + 2 Then
        Call MCosta(1, tp2(2).Tag, tx, ty - 1)
        Call MCosta(2, tp2(5).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(1).Tag) + 0 Then 'Click arriba
        Call MCosta(1, tp2(5).Tag, tx, ty)
        Call MCosta(2, tp2(8).Tag, tx, ty + 1)
    ElseIf gh(7) = DameGrhIndex(tp1(1).Tag) + 1 Then
        Call MCosta(1, tp2(3).Tag, tx, ty)
        Call MCosta(2, tp2(6).Tag, tx, ty + 1)
    ElseIf gh(7) = DameGrhIndex(tp1(6).Tag) + 0 Then
        Call MCosta(1, tp1(3).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(8).Tag) + 1 Then
        Call MCosta(1, tp1(5).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(3).Tag) Or gh(7) = DameGrhIndex(tp1(3).Tag) + 2 Then
        Call MCosta(1, tp1(3).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(5).Tag) + 1 Or gh(7) = DameGrhIndex(tp1(5).Tag) + 3 Then
        Call MCosta(1, tp1(5).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(7).Tag) + 3 Then
        Call MCosta(1, tp1(8).Tag, tx, ty + 1)
        Call MCosta(2, tp1(8).Tag, tx - 1, ty + 1)
        Call MCosta(3, tp1(8).Tag, tx, ty)
    ElseIf gh(7) = DameGrhIndex(tp1(7).Tag) + 2 Then
        Call MCosta(1, tp1(6).Tag, tx, ty + 1)
        Call MCosta(2, tp1(6).Tag, tx + 1, ty + 1)
        Call MCosta(3, tp1(6).Tag, tx, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(3).Tag) + 2 Then 'Click izquierda
        Call MCosta(1, tp2(1).Tag, tx, ty)
        Call MCosta(2, tp2(2).Tag, tx + 1, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(3).Tag) + 0 Then
        Call MCosta(1, tp2(7).Tag, tx, ty)
        Call MCosta(2, tp2(8).Tag, tx + 1, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(5).Tag) + 1 Then
        Call MCosta(1, tp1(2).Tag, tx, ty)
        Call MCosta(2, tp1(2).Tag, tx + 1, ty)
        Call MCosta(3, tp1(2).Tag, tx + 1, ty + 1)
    ElseIf gh(5) = DameGrhIndex(tp1(5).Tag) + 3 Then
        Call MCosta(1, tp1(8).Tag, tx, ty)
        Call MCosta(2, tp1(8).Tag, tx + 1, ty)
        Call MCosta(3, tp1(8).Tag, tx + 1, ty - 1)
    ElseIf gh(5) = DameGrhIndex(tp1(1).Tag) Or gh(5) = DameGrhIndex(tp1(1).Tag) + 1 Then
        Call MCosta(1, tp1(1).Tag, tx, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(7).Tag) + 2 Or gh(5) = DameGrhIndex(tp1(7).Tag) + 3 Then
        Call MCosta(1, tp1(7).Tag, tx, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(2).Tag) + 0 Then
        Call MCosta(1, tp1(1).Tag, tx, ty)
    ElseIf gh(5) = DameGrhIndex(tp1(8).Tag) + 2 Then
        Call MCosta(1, tp1(7).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(3).Tag) + 0 Then 'Click derecha
        Call MCosta(1, tp1(0).Tag, tx, ty)
        Call MCosta(2, tp1(0).Tag, tx - 1, ty)
        Call MCosta(3, tp1(0).Tag, tx - 1, ty + 1)
    ElseIf gh(3) = DameGrhIndex(tp1(3).Tag) + 2 Then
        Call MCosta(1, tp1(6).Tag, tx, ty)
        Call MCosta(2, tp1(6).Tag, tx - 1, ty)
        Call MCosta(3, tp1(6).Tag, tx - 1, ty - 1)
    ElseIf gh(3) = DameGrhIndex(tp1(5).Tag) + 3 Then
        Call MCosta(1, tp2(0).Tag, tx - 1, ty)
        Call MCosta(2, tp2(1).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(5).Tag) + 1 Then
        Call MCosta(1, tp2(6).Tag, tx - 1, ty)
        Call MCosta(2, tp2(7).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(1).Tag) Or gh(3) = DameGrhIndex(tp1(1).Tag) + 1 Then
        Call MCosta(1, tp1(1).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(7).Tag) + 2 Or gh(3) = DameGrhIndex(tp1(7).Tag) + 3 Then
        Call MCosta(1, tp1(7).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(0).Tag) + 1 Then
        Call MCosta(1, tp1(1).Tag, tx, ty)
    ElseIf gh(3) = DameGrhIndex(tp1(6).Tag) + 3 Then
        Call MCosta(1, tp1(7).Tag, tx, ty)
    End If
End Sub

Sub MCosta(Pos As Byte, grhindex As Integer, X As Integer, Y As Integer)
If Pos = 0 Then
    MCostaV(1).grhindex = 0
    MCostaV(1).X = 0
    MCostaV(1).Y = 0
    MCostaV(2).grhindex = 0
    MCostaV(2).X = 0
    MCostaV(2).Y = 0
    MCostaV(3).grhindex = 0
    MCostaV(3).X = 0
    MCostaV(3).Y = 0
Else
    MCostaV(Pos).grhindex = DameGrhIndex(grhindex)
    MCostaV(Pos).X = X
    MCostaV(Pos).Y = Y
End If
End Sub


Private Sub Renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub SelectPanel_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte

For i = 0 To 7
    If i <> Index Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
Call VerFuncion(Index, SelectPanel(Index).value)
End Sub

Private Sub TimAutoGuardarMapa_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If mnuAutoGuardarMapas.Checked = True Then
    bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1
    If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
        If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
            modMapIO.GuardarMapa Dialog.FileName
        End If
        bAutoGuardarMapaCount = 0
    End If
End If
End Sub


Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

With Dialog
    .Filter = "Mapas de Argentum Online (*.map)|*.map"
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
    Else
        .DialogTitle = "Cargar"
        .FileName = vbNullString
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With
End Sub

Private Sub tp1_Click(Index As Integer)
'Dim i As Integer
'For i = 0 To lListado(0).ListCount - 1
'    If (InStr(1, lListado(0).List(i), "#" & tp1(Index).Tag, vbTextCompare) > 0) Then
'        lListado(0).ListIndex = i
'        Exit For
'    End If
'Next i

cGrh.text = DameGrhIndex(tp1(Index).Tag)
cCapas.text = IIf(Index = 4, 1, 2)
If Index = 4 Then
    frmConfigSup.mLargo.text = 4
    frmConfigSup.mAncho.text = 4
    chkCostas.value = vbUnchecked
End If
Call fPreviewGrh(cGrh.text)
Call modPaneles.VistaPreviaDeSup



End Sub

Private Sub tp2_Click(Index As Integer)
'Dim i As Integer
'For i = 0 To lListado(0).ListCount - 1
'    If (InStr(1, lListado(0).List(i), "#" & tp2(Index).Tag, vbTextCompare) > 0) Then
'        lListado(0).ListIndex = i
'        Exit For
'    End If
'Next i

cGrh.text = DameGrhIndex(tp2(Index).Tag)
cCapas.text = IIf(Index = 4, 1, 2)
If Index = 4 Then
    frmConfigSup.mLargo.text = 4
    frmConfigSup.mAncho.text = 4
    chkCostas.value = vbUnchecked
End If
Call fPreviewGrh(cGrh.text)
Call modPaneles.VistaPreviaDeSup

End Sub

