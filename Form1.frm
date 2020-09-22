VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OpenGL Color and Lighting Model Demo..."
   ClientHeight    =   6660
   ClientLeft      =   360
   ClientTop       =   2040
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Object Transparency..."
      Height          =   735
      Left            =   120
      TabIndex        =   77
      Top             =   5760
      Width           =   5415
      Begin MSComctlLib.Slider sldAlpha 
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   15
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   525
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "full"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   80
         Top             =   525
         Width           =   495
      End
   End
   Begin VB.Frame frameObjects 
      Caption         =   "Object Color Properties..."
      Height          =   3255
      Left            =   120
      TabIndex        =   43
      Top             =   2400
      Width           =   5415
      Begin VB.OptionButton optCone 
         Caption         =   "Cone"
         Height          =   255
         Left            =   3700
         TabIndex        =   76
         Top             =   2760
         Width           =   855
      End
      Begin VB.OptionButton optCylinder 
         Caption         =   "Cylinder"
         Height          =   255
         Left            =   2150
         TabIndex        =   75
         Top             =   2760
         Width           =   975
      End
      Begin VB.OptionButton optSphere 
         Caption         =   "Sphere"
         Height          =   255
         Left            =   600
         TabIndex        =   74
         Top             =   2760
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSComctlLib.Slider sldAR2 
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldAG2 
         Height          =   255
         Left            =   720
         TabIndex        =   45
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldAB2 
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldDR2 
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldDG2 
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldDB2 
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldSR2 
         Height          =   255
         Left            =   4080
         TabIndex        =   50
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
      End
      Begin MSComctlLib.Slider sldSG2 
         Height          =   255
         Left            =   4080
         TabIndex        =   51
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
      End
      Begin MSComctlLib.Slider sldSB2 
         Height          =   255
         Left            =   4080
         TabIndex        =   52
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
      End
      Begin VB.Label Label11 
         Caption         =   "Object type:"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5160
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   73
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         Caption         =   "Ambient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   71
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   70
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   68
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   66
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "Diffuse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3480
         TabIndex        =   60
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "Specular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   58
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3480
         TabIndex        =   57
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   56
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   54
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame frameLighting 
      Caption         =   "Lighting Color Properties..."
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.Slider sldAR 
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider sldAG 
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   10
      End
      Begin MSComctlLib.Slider sldAB 
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   10
      End
      Begin MSComctlLib.Slider sldDR 
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   8
         Value           =   8
      End
      Begin MSComctlLib.Slider sldDG 
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   8
      End
      Begin MSComctlLib.Slider sldDB 
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   8
      End
      Begin MSComctlLib.Slider sldSR 
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldSG 
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldSB 
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   2
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   39
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Specular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Diffuse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Ambient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.0    0.5    1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Light Position..."
      Height          =   1575
      Left            =   5640
      TabIndex        =   4
      Top             =   4080
      Width           =   3735
      Begin MSComctlLib.Slider sldLightPosition 
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   30
         SmallChange     =   10
         Min             =   -150
         Max             =   150
         TickFrequency   =   15
      End
      Begin VB.OptionButton optAxisZ 
         Caption         =   "Z"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton optAxisY 
         Caption         =   "Y"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton optAxisX 
         Caption         =   "X"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "max"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2980
         TabIndex        =   11
         Top             =   1240
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "min"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   245
         TabIndex        =   10
         Top             =   1240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Select axis:"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object Shading..."
      Height          =   735
      Left            =   5640
      TabIndex        =   1
      Top             =   5760
      Width           =   3735
      Begin VB.OptionButton optFlatShading 
         Caption         =   "Flat"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSmoothShading 
         Caption         =   "Smooth"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox picOutput 
      Height          =   3750
      Left            =   5640
      ScaleHeight     =   3690
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   240
      Width           =   3750
   End
   Begin VB.Timer timRotate 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'* OpenGL Color and Lighting Model Demonstration in Visual Basic *
'*      By: Daniel S. Soper                                      *
'*****************************************************************
'*                  If you like it, vote for it!                 *
'*****************************************************************

Option Explicit 'explicitly declare all variables


Private Const Pi = 3.14159265359 'PI value

Dim hGLRC As Long 'handle to the gl rendering context
Dim LampPosition(0 To 3) As Single 'holds the lamp's position coordinates
Dim AngleOfRotation As Integer 'the current angle of rotation from the origin
Dim ObjectAlpha As Single 'the current alpha level of the object

Private Function Initialize() As Boolean
    Dim PFD As PIXELFORMATDESCRIPTOR 'Describes the pixel format of a drawing surface
    Dim RetVal As Long 'Return value
    
    PFD.nSize = Len(PFD)
    PFD.nVersion = 1
    PFD.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    PFD.iPixelType = PFD_TYPE_RGBA 'set the pixel type
    PFD.cColorBits = 24 'set the color bits
    PFD.cDepthBits = 16 'set the depth bits
    PFD.iLayerType = PFD_MAIN_PLANE 'set the layer type
    RetVal = ChoosePixelFormat(picOutput.hDC, PFD) 'Retreive a compatible pixel format for picOutput
    If RetVal = 0 Then 'If a compatible pixel format could not be found
        MsgBox "Could not find a compatible pixel format.", vbCritical + vbOKOnly, "Initialization Error" 'Display error message
        Initialize = False 'Initialization of the OpenGL drawing surface was not successful
        End 'Exit program
    End If
    RetVal = SetPixelFormat(picOutput.hDC, RetVal, PFD) 'Set the pixel format of picOutput
 
    hGLRC = wglCreateContext(picOutput.hDC) 'Return the handle to an OpenGL rendering context that is compatible with picOutput
    wglMakeCurrent picOutput.hDC, hGLRC 'Set the current OpenGL rendering context to picOutput
    glClearColor 0, 0, 0, 1 'Set the clear color
    glEnable glcColorMaterial 'Allow material parameters track the current color
    glColorMaterial faceFront, cmmAmbientAndDiffuse 'Enable color tracking
    glClearDepth 1 'Set the clear value for the depth buffer
    
    glEnable glcBlend 'enable alpha blending
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha 'set alpha blending options
    glEnable glcDepthTest   'Enable depth comparisons
    glEnable glcLighting    'Enable lighting operations
    glEnable glcLight0      'Enable this light
    LampPosition(0) = 100    'X-coordinate of light from origin
    LampPosition(1) = -100    'Y-coordinate of light from origin
    LampPosition(2) = 150    'Z-coordinate of light from origin
    glLightfv ltLight0, lpmPosition, LampPosition(0) 'set the position of the light
    optAxisX_Click 'set the light position slider bar to the correct position
    ObjectAlpha = 1 'set initial alpha level to 1 (no transparency)
    Initialize = True 'initialization succeeded
End Function

Private Sub DrawObjects()
    Dim Obj As Long 'holds the handle to the object
    Dim ObjAmbient(0 To 3) As Single 'holds the object's ambient RGBA values
    Dim ObjDiffuse(0 To 3) As Single 'holds the object's diffuse RGBA values
    Dim ObjSpecular(0 To 3) As Single 'holds the object's specular RGBA values
    
    ObjAmbient(0) = sldAR2.Value / 10 'set ambient red value
    ObjAmbient(1) = sldAG2.Value / 10 'set ambient green value
    ObjAmbient(2) = sldAB2.Value / 10 'set ambient blue value
    ObjAmbient(3) = 1 'set ambient transparency value
    
    ObjDiffuse(0) = sldDR2.Value / 10 'set diffuse red value
    ObjDiffuse(1) = sldDG2.Value / 10 'set diffuse green value
    ObjDiffuse(2) = sldDB2.Value / 10 'set diffuse blue value
    ObjDiffuse(3) = ObjectAlpha 'set diffuse transparency value

    ObjSpecular(0) = sldSR2.Value / 10 'set specular red value
    ObjSpecular(1) = sldSG2.Value / 10 'set specular green value
    ObjSpecular(2) = sldSB2.Value / 10 'set specular blue value
    ObjSpecular(3) = 1 'set specular transparency value

    Obj = gluNewQuadric 'create new quadric object
    glPushMatrix 'push the matrix stack down by one
        glColor4f ObjDiffuse(0), ObjDiffuse(1), ObjDiffuse(2), ObjDiffuse(3) 'set the object's diffuse color
        glMaterialfv faceFront, mprAmbient, ObjAmbient(0) 'set the object's ambient color
        glMaterialfv faceFront, mprSpecular, ObjSpecular(0) 'set the object's specular color
    If optSphere.Value = True Then 'draw a sphere
        gluSphere Obj, 100, 16, 16
    ElseIf optCylinder.Value = True Then 'draw a cylinder
        glTranslatef 0, 0, 50 'multiply the current matrix by these XYZ coordinates
        gluCylinder Obj, 35, 35, 100, 16, 16
    Else 'draw a cone
        gluCylinder Obj, 100, 0, 150, 12, 12
    End If
    glPopMatrix 'pop the matrix stack up by one
End Sub

Private Sub Render()
    Static Busy As Boolean 'holds the current rendering process status

    If Busy = True Then 'if currently rendering
        Exit Sub
    Else
        Busy = True
    End If
    
    wglMakeCurrent picOutput.hDC, hGLRC 'sets this thread's current rendering context
    
    If optSmoothShading.Value = True Then
        glShadeModel smSmooth 'Set the shading mode to smooth
    Else
        glShadeModel smFlat 'Set the shading mode to flat
    End If
    
    glClear clrColorBufferBit Or clrDepthBufferBit 'clear the color and depth buffers
    
    DrawLamp 'call the subroutine that draws the lamp
    
    glRGBA 255, 0, 0, 1       'Red rectangle
    glRectf -150, 150, -50, 0
    glRGBA 0, 255, 0, 1       'Green rectangle
    glRectf -50, 150, 50, 0
    glRGBA 0, 0, 255, 1       'Blue rectangle
    glRectf 50, 150, 150, 0
    glRGBA 255, 255, 255, 1   'White rectangle
    glRectf -150, 0, 0, -150
    glRGBA 0, 0, 0, 1         'Black rectangle
    glRectf 0, 0, 150, -150
        
    glPushMatrix 'push the matrix stack down by one
        glRotated AngleOfRotation, 0.2, 0.2, 0.5 'Rotate object
        DrawObjects 'call the subroutine that draws the objects
    glPopMatrix 'pop the matrix stack up by one
    SwapBuffers picOutput.hDC 'exchange the front and back buffers
    wglMakeCurrent 0, 0
    Busy = False 'rendering process is complete
End Sub

Private Sub DrawLamp()
    Dim objLamp As Long 'holds the handle to the sphere that represents the lamp
    Dim LampAmbient(0 To 3) As Single 'holds the lamp's ambient RGBA values
    Dim LampDiffuse(0 To 3) As Single 'holds the lamp's diffuse RGBA values
    Dim LampSpecular(0 To 3) As Single 'holds the lamp's specular RGBA values
    
    LampAmbient(0) = sldAR.Value / 10 'set ambient red value
    LampAmbient(1) = sldAG.Value / 10 'set ambient green value
    LampAmbient(2) = sldAB.Value / 10 'set ambient blue value
    LampAmbient(3) = 1 'set ambient transparency value
    
    LampDiffuse(0) = sldDR.Value / 10 'set diffuse red value
    LampDiffuse(1) = sldDG.Value / 10 'set diffuse green value
    LampDiffuse(2) = sldDB.Value / 10 'set diffuse blue value
    LampDiffuse(3) = 1 'set diffuse transparency value

    LampSpecular(0) = sldSR.Value / 10 'set specular red value
    LampSpecular(1) = sldSG.Value / 10 'set specular green value
    LampSpecular(2) = sldSB.Value / 10 'set specular blue value
    LampSpecular(3) = 1 'set specular transparency value
    
    objLamp = gluNewQuadric 'create a new quadric object
    glDisable glcLighting 'temporarily disable lighting
        glPushMatrix 'push the matrix stack down by one
            glColor4f LampDiffuse(0), LampDiffuse(1), LampDiffuse(2), 1 'set the color of the sphere that represents the lamp
            glTranslatef LampPosition(0), LampPosition(1), LampPosition(2) 'move the lamp coordinates if necessary
            gluSphere objLamp, 8, 12, 12 'draw the sphere that represents the lamp
            glLightfv ltLight0, lpmAmbient, LampAmbient(0) 'set the lamp's ambient color values
            glLightfv ltLight0, lpmDiffuse, LampDiffuse(0) 'set the lamp's diffuse color values
            glLightfv ltLight0, lpmSpecular, LampSpecular(0) 'set the lamp's specular color values
            glLightfv ltLight0, lpmPosition, LampPosition(0) 'set the position of the light
        glPopMatrix 'pop the matrix stack up by one
    glEnable glcLighting 're-enable lighting
End Sub

Private Sub Form_Load()
    Initialize 'Execute the initialization procedure
End Sub

Private Sub Form_Paint()
    Render 'Execute the rendering procedure
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If hGLRC <> 0 Then 'if the OpenGL rendering context still exists
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC 'remove the OpenGL rendering context from memory
    End If
End Sub

Private Sub Form_Resize()
    Dim viewportW As Long 'gl viewport width
    Dim viewportH As Long 'gl viewport height

    viewportW = picOutput.Width  'set gl viewport width
    viewportH = picOutput.Height  'set gl viewport height
    
    wglMakeCurrent picOutput.hDC, hGLRC 'sets this thread's current rendering context
        glViewport 0, 0, viewportW, viewportH 'set the viewport
        glMatrixMode mmProjection 'set the current matrix mode
        glLoadIdentity 'load identity matrix
        glOrtho -153, 159, -154, 158, -500, 500 'create a parallel projection
        glMatrixMode mmModelView 'set the current matrix mode
        glLoadIdentity 'load the identity matrix
    wglMakeCurrent 0, 0
    Render 'Execute the rendering procedure
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hGLRC <> 0 Then 'if the OpenGL rendering context still exists
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC 'remove the OpenGL rendering context from memory
    End If
End Sub

Private Sub optAxisX_Click()
    sldLightPosition.Value = LampPosition(0) 'set the slider bar to the current light position on the X-axis
End Sub

Private Sub optAxisY_Click()
    sldLightPosition.Value = LampPosition(1) 'set the slider bar to the current light position on the Y-axis
End Sub

Private Sub optAxisZ_Click()
    sldLightPosition.Value = LampPosition(2) - 150 'set the slider bar to the current light position on the Z-axis
End Sub

Private Sub sldAlpha_Change()
    ObjectAlpha = (100 - sldAlpha.Value) / 100 'set object alpha level
End Sub

Private Sub sldAlpha_Click()
    sldAlpha_Change 'call this slider's "change" event
End Sub

Private Sub sldLightPosition_Change()
    If optAxisX.Value = True Then 'move the light along the X-axis
        LampPosition(0) = sldLightPosition.Value
    ElseIf optAxisY.Value = True Then 'move the light along the Y-axis
        LampPosition(1) = sldLightPosition.Value
    ElseIf optAxisZ.Value = True Then 'move the light along the Z-axis
        LampPosition(2) = sldLightPosition.Value + 150
    End If
End Sub

Private Sub sldLightPosition_Click()
    sldLightPosition_Change 'call this slider's "change" event
End Sub

Private Sub timRotate_Timer() 'increments the angle of rotation
    AngleOfRotation = AngleOfRotation + 3 'add 3 degrees to the angle of rotation
    If AngleOfRotation > 357 Then
        AngleOfRotation = 0
    End If
    Render 'Execute the rendering procedure
End Sub

Private Sub glRGBA(RedIn, GreenIn, BlueIn, AlphaIn) 'Converts standard RGB values into the gl model (+ Alpha level)
    glColor4f (RedIn / 255), (GreenIn / 255), (BlueIn / 255), AlphaIn
End Sub
