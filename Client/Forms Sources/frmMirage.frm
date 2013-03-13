VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form FrmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13320
   ClientLeft      =   8040
   ClientTop       =   3210
   ClientWidth     =   23475
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMirage.frx":08EE
   ScaleHeight     =   877.365
   ScaleMode       =   0  'User
   ScaleWidth      =   1565
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   607
      Index           =   1
      Left            =   0
      Picture         =   "frmMirage.frx":15B8
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   256
      Top             =   1560
      Width           =   600
   End
   Begin VB.Frame fra_fenetre 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   8625
      Left            =   15360
      TabIndex        =   223
      Top             =   120
      Width           =   8955
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2505
         Left            =   5040
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   159
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2385
         Begin VB.PictureBox picItems 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2.25000e5
            Left            =   2400
            Picture         =   "frmMirage.frx":28BA
            ScaleHeight     =   2.23636e5
            ScaleMode       =   0  'User
            ScaleWidth      =   477.091
            TabIndex        =   236
            Top             =   2760
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1080
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   234
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox LegsImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   235
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   232
            Top             =   1920
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox BootsImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   233
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1680
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   230
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox Ring2Image 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   231
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   480
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   228
            Top             =   1320
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox Ring1Image 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   229
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   480
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   226
            Top             =   1920
            Visible         =   0   'False
            Width           =   555
            Begin VB.PictureBox GlovesImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   227
               Top             =   15
               Width           =   495
            End
         End
      End
      Begin VB.Frame fraCarte 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   840
         TabIndex        =   224
         Top             =   4320
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Image imgCarte 
            Height          =   2295
            Left            =   0
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   2985
         Left            =   4080
         Picture         =   "frmMirage.frx":1621FC
         Top             =   2880
         Width           =   2595
      End
      Begin VB.Label freduire 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   1920
         TabIndex        =   241
         Top             =   0
         Width           =   375
      End
      Begin VB.Label ffermer 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   2280
         TabIndex        =   240
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblmaskinv 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   480
         MousePointer    =   5  'Size
         TabIndex        =   239
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblmaskinvmin 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   2040
         TabIndex        =   238
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblmaskinvferm 
         BackStyle       =   0  'Transparent
         Caption         =   "                                   "
         Height          =   375
         Left            =   2280
         TabIndex        =   237
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox PicInterface 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7155
      Left            =   1440
      Picture         =   "frmMirage.frx":17B676
      ScaleHeight     =   475
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   190
      Top             =   840
      Visible         =   0   'False
      Width           =   9165
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   720
         ScaleHeight     =   3255
         ScaleWidth      =   7455
         TabIndex        =   247
         Top             =   2760
         Width           =   7455
         Begin VB.PictureBox Picture17 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3480
            Picture         =   "frmMirage.frx":250518
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   251
            Top             =   2880
            Width           =   270
         End
         Begin VB.PictureBox Picture18 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3840
            Picture         =   "frmMirage.frx":2507B0
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   250
            Top             =   2880
            Width           =   270
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   5000
            Left            =   0
            Picture         =   "frmMirage.frx":250A3B
            ScaleHeight     =   333
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   494
            TabIndex        =   248
            Top             =   0
            Width           =   7410
            Begin VB.PictureBox picspell 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   249
               Top             =   120
               Width           =   480
            End
            Begin VB.Shape SDAD 
               BorderColor     =   &H00008000&
               BorderWidth     =   3
               Height          =   510
               Left            =   120
               Top             =   105
               Visible         =   0   'False
               Width           =   510
            End
         End
      End
      Begin VB.PictureBox tmpsquete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   3600
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   244
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label minutes 
            BackStyle       =   0  'Transparent
            Caption         =   "00:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   246
            ToolTipText     =   "Minutes restante avant la fin de la quête en cour"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label seconde 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   245
            ToolTipText     =   "Secondes restante avant la fin de la quête en cour"
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         ItemData        =   "frmMirage.frx":2A8BE5
         Left            =   5880
         List            =   "frmMirage.frx":2A8BE7
         TabIndex        =   243
         Top             =   4440
         Width           =   1860
      End
      Begin VB.PictureBox AmuletImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   242
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox picInv3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3600
         Left            =   780
         Picture         =   "frmMirage.frx":2A8BE9
         ScaleHeight     =   240
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   150
         TabIndex        =   217
         Top             =   630
         Width           =   2250
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   75
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   219
            Top             =   15
            Width           =   480
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   330
            Left            =   2640
            Max             =   100
            TabIndex        =   218
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape IDAD 
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            Height          =   525
            Left            =   1560
            Top             =   2040
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Shape EquipS 
            BorderColor     =   &H00D67A27&
            Height          =   510
            Index           =   1
            Left            =   1080
            Top             =   1320
            Width           =   510
         End
         Begin VB.Shape EquipS 
            BorderColor     =   &H00D67A27&
            Height          =   510
            Index           =   2
            Left            =   600
            Top             =   1560
            Width           =   510
         End
         Begin VB.Shape EquipS 
            BorderColor     =   &H00D67A27&
            Height          =   510
            Index           =   3
            Left            =   0
            Top             =   1680
            Width           =   510
         End
         Begin VB.Shape SelectedItem 
            BorderColor     =   &H000000FF&
            Height          =   510
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Shape EquipS 
            BorderColor     =   &H00D67A27&
            Height          =   510
            Index           =   0
            Left            =   240
            Top             =   840
            Width           =   510
         End
         Begin VB.Shape EquipS 
            BorderColor     =   &H00D67A27&
            Height          =   510
            Index           =   4
            Left            =   240
            Top             =   960
            Width           =   510
         End
      End
      Begin VB.PictureBox itmDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   765
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   319
         TabIndex        =   206
         Top             =   4440
         Width           =   4785
         Begin VB.Label descName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nom"
            ForeColor       =   &H80000005&
            Height          =   240
            Left            =   120
            TabIndex        =   216
            ToolTipText     =   "Nom de l'objet"
            Top             =   0
            Width           =   4560
         End
         Begin VB.Shape Usure2 
            BackColor       =   &H0083031A&
            BackStyle       =   1  'Opaque
            FillColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            Top             =   0
            Width           =   4575
         End
         Begin VB.Label desc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000005&
            Height          =   855
            Left            =   120
            TabIndex        =   215
            ToolTipText     =   "Description de l'objet"
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label descMS 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Magi: XXXXX Speed: XXXX"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   0
            TabIndex        =   214
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label descSD 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Str: XXXX Def: XXXXX"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   0
            TabIndex        =   213
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label descHpMp 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   0
            TabIndex        =   212
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "-Donne-"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   360
            TabIndex        =   211
            ToolTipText     =   "Se que vous apporte l'objet"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label descSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   2760
            TabIndex        =   210
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label descDef 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Endurance"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   2760
            TabIndex        =   209
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label descStr 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Strength"
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   2760
            TabIndex        =   208
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "-Requière-"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   2760
            TabIndex        =   207
            ToolTipText     =   "Force/défense/vitesse requise pour équipper l'objet"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picsprts 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   3720
         ScaleHeight     =   690
         ScaleWidth      =   465
         TabIndex        =   195
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox ArmorImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   194
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox ShieldImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4200
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   193
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox HelmetImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   192
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox WeaponImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   191
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblmana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000/0000"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   180
         Left            =   3120
         TabIndex        =   255
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lblvie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000/0000"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   180
         Left            =   3120
         TabIndex        =   254
         Top             =   840
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Height          =   300
         Left            =   8700
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   8760
         TabIndex        =   252
         Top             =   120
         Width           =   180
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   28
         Left            =   3345
         TabIndex        =   222
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Utiliser"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   221
         Top             =   4200
         Width           =   690
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jeter"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2040
         TabIndex        =   220
         Top             =   4200
         Width           =   555
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   300
         Left            =   4200
         TabIndex        =   205
         Top             =   3480
         Width           =   435
      End
      Begin VB.Label AddDef 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6480
         TabIndex        =   204
         Top             =   1920
         Width           =   195
      End
      Begin VB.Label AddMagi 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6480
         TabIndex        =   203
         Top             =   2640
         Width           =   195
      End
      Begin VB.Label AddSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6480
         TabIndex        =   202
         Top             =   3480
         Width           =   195
      End
      Begin VB.Label AddStr 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6480
         TabIndex        =   201
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblSPEED 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   345
         Left            =   5760
         TabIndex        =   200
         ToolTipText     =   "Points permettant d'augmenter vos chances d'esquive"
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label lblMAGI 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   345
         Left            =   5760
         TabIndex        =   199
         ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label lblDEF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   345
         Left            =   5760
         TabIndex        =   198
         ToolTipText     =   "Points permettant d'augmenter votre résistance et vos chances de bloquer"
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label lblSTR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   225
         Left            =   5760
         TabIndex        =   197
         ToolTipText     =   "Points permettant d'augmenter vos dégâts et vos chances de coup critique"
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label monnom 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TEST PSEUDO"
         BeginProperty Font 
            Name            =   "Dungeon"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005982C4&
         Height          =   180
         Left            =   3360
         TabIndex        =   196
         Top             =   600
         Width           =   1350
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   5520
      Top             =   0
   End
   Begin VB.PictureBox Configurer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   15840
      ScaleHeight     =   1785
      ScaleWidth      =   1425
      TabIndex        =   150
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Emotes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   156
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Guilde"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   155
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Privés"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   154
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Admins"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   153
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Local /Serveur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   152
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Joueurs / Map"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   151
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   960
         TabIndex        =   157
         Top             =   1560
         Width           =   345
      End
   End
   Begin VB.TextBox canaltext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   258
      Left            =   15960
      MaxLength       =   9
      TabIndex        =   149
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Popup 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   15360
      ScaleHeight     =   1065
      ScaleWidth      =   1185
      TabIndex        =   145
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label menu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Effacer"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   185
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Annuler"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   148
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Configurer"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   147
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Renommer"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   146
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   425
      Left            =   10200
      Picture         =   "frmMirage.frx":2C2E9F
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   113
      Top             =   9413
      Width           =   390
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   10095
      Left            =   12000
      Picture         =   "frmMirage.frx":2C37A1
      ScaleHeight     =   673
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   111
      Top             =   0
      Width           =   3315
      Begin VB.TextBox classement 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000000&
         Height          =   855
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   186
         Top             =   9240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin RichTextLib.RichTextBox RTBChat 
         Height          =   7695
         Index           =   0
         Left            =   0
         TabIndex        =   140
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   13573
         _Version        =   393217
         BackColor       =   8421504
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMirage.frx":2C5E99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox classement 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000000&
         Height          =   855
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   9240
         Width           =   3255
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Top 3:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   189
         Top             =   9000
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guildes"
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   188
         Top             =   9000
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Joueurs"
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   187
         Top             =   9000
         Width           =   735
      End
      Begin VB.Label Onglet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canal"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   144
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Onglet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canal"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   143
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Onglet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canal"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   142
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Onglet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAIN"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   141
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      BackColor       =   &H003FCBAD&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   3  'Vertical Line
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5760
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   100
      Top             =   5280
      Visible         =   0   'False
      Width           =   3390
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Emotes"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   110
         Top             =   3000
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Messages de Guilde"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   109
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Messages Privés"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   105
         Top             =   2520
         Width           =   2175
      End
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Messages Interne"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   104
         Top             =   2280
         Width           =   2175
      End
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Messages Admins"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   103
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton OptionColor 
         Appearance      =   0  'Flat
         BackColor       =   &H003FCBAD&
         Caption         =   "Messages de Joueurs"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   102
         Top             =   1800
         Width           =   2175
      End
      Begin VB.PictureBox PicColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   25
         Picture         =   "frmMirage.frx":2C5F10
         ScaleHeight     =   1665
         ScaleWidth      =   3270
         TabIndex        =   101
         Top             =   25
         Width           =   3300
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   0
         Top             =   3000
         Width           =   495
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   0
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   107
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label ColorSave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enregistrer"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   0
         Top             =   2520
         Width           =   495
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   0
         Top             =   2280
         Width           =   495
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   2040
         Width           =   495
      End
      Begin VB.Shape ShapeColor1 
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Timer Chat 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6360
      Top             =   7560
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H003FCBAD&
      ForeColor       =   &H80000008&
      Height          =   6105
      Left            =   9240
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   45
      Top             =   3000
      Visible         =   0   'False
      Width           =   2625
      Begin VB.CommandButton CommandColor 
         Caption         =   "Couleurs du T'chat"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   5520
         Width           =   2415
      End
      Begin VB.TextBox txtTempsBulles 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   98
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   5760
         Width           =   2415
      End
      Begin VB.CommandButton CmdoptTouche 
         Caption         =   "Configurer les touches"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   5280
         Width           =   2415
      End
      Begin VB.CheckBox chknobj 
         BackColor       =   &H003FCBAD&
         Caption         =   "Nom des objets aux sol (quand la souris le survole)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "Petite barre affichée au dessus de vous"
         Top             =   960
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerbar 
         BackColor       =   &H003FCBAD&
         Caption         =   "Mini barre de vie"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H003FCBAD&
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chknpcname 
         BackColor       =   &H003FCBAD&
         Caption         =   "Noms"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkbubblebar 
         BackColor       =   &H003FCBAD&
         Caption         =   "Bulles de dialogue"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chknpcbar 
         BackColor       =   &H003FCBAD&
         Caption         =   "Affichés leur mini barre de vie"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   55
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerdamage 
         BackColor       =   &H003FCBAD&
         Caption         =   "Dégâts affichés au dessus de la tête"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox chknpcdamage 
         BackColor       =   &H003FCBAD&
         Caption         =   "Dégâts affichés au dessus de la tête"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H003FCBAD&
         Caption         =   "Musique"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   52
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H003FCBAD&
         Caption         =   "Effets sonores"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkAutoScroll 
         BackColor       =   &H003FCBAD&
         Caption         =   "Défilement automatique"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   4
         TabIndex        =   49
         Top             =   4125
         Value           =   6
         Width           =   2055
      End
      Begin VB.CheckBox chkLowEffect 
         BackColor       =   &H003FCBAD&
         Caption         =   "Désactiver les effets avancés"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   4680
         Width           =   2325
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "- Options du t'Chat -"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   184
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-Musique / Sons -"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   0
         TabIndex        =   183
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "- Affichage des NPCS -"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   182
         Top             =   1275
         Width           =   2535
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "- Affichage du Joueur -"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   181
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -120
         TabIndex        =   180
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lblBulle 
         BackStyle       =   0  'Transparent
         Caption         =   "Temps d'affichage des bulles:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de ligne écrite sur l'écran: 6"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   61
         Top             =   3960
         Width           =   2220
      End
   End
   Begin VB.PictureBox pictMetier 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   4680
      Picture         =   "frmMirage.frx":2D7E52
      ScaleHeight     =   1965
      ScaleWidth      =   3600
      TabIndex        =   89
      Top             =   4440
      Visible         =   0   'False
      Width           =   3630
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblOublierMetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Oublier le Metier"
         Height          =   255
         Left            =   1440
         TabIndex        =   95
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblendmetier 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3390
         TabIndex        =   94
         Top             =   1935
         Width           =   375
      End
      Begin VB.Label lblmetierEnd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fermer"
         Height          =   255
         Left            =   2760
         TabIndex        =   93
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   92
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.PictureBox PicMenuQuitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   6360
      Picture         =   "frmMirage.frx":2EEF04
      ScaleHeight     =   1965
      ScaleWidth      =   3600
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   3600
      Begin VB.Label lblCdP 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   66
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDeco 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   65
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblQuitter 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   64
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   63
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   7125
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   43
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   42
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   6045
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   41
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   5505
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   4965
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   4425
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   3345
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   2805
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   35
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2265
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   32
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   645
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   31
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   105
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   30
      Top             =   9315
      Width           =   480
   End
   Begin VB.ComboBox Canal 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMirage.frx":2F1AB6
      Left            =   120
      List            =   "frmMirage.frx":2F1AC6
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Carte"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1335
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   21
      Top             =   8760
      Visible         =   0   'False
      Width           =   6405
   End
   Begin VB.Timer quetetimersec 
      Enabled         =   0   'False
      Left            =   9240
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picturesprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9300
      Left            =   0
      ScaleHeight     =   620
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   5
         Left            =   0
         Picture         =   "frmMirage.frx":2F1AE6
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   260
         Top             =   3960
         Width           =   600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   4
         Left            =   0
         Picture         =   "frmMirage.frx":2F2DE8
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   259
         Top             =   3360
         Width           =   600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   3
         Left            =   0
         Picture         =   "frmMirage.frx":2F40EA
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   258
         Top             =   2760
         Width           =   600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   0
         Picture         =   "frmMirage.frx":2F53EC
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   257
         Top             =   2160
         Width           =   600
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   0
         Left            =   0
         Picture         =   "frmMirage.frx":2F66EE
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   253
         Top             =   960
         Width           =   600
      End
      Begin VB.PictureBox pictTouche 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4035
         Left            =   3120
         ScaleHeight     =   267
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   511
         TabIndex        =   114
         Top             =   600
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton cmdOTO 
            Caption         =   "Ok"
            Height          =   255
            Left            =   6120
            TabIndex        =   116
            Top             =   3720
            Width           =   735
         End
         Begin VB.CommandButton cmdOTA 
            Caption         =   "Annuler"
            Height          =   255
            Left            =   6840
            TabIndex        =   115
            Top             =   3720
            Width           =   735
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Appuyez sur une action pour la reconfigurer"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   179
            Top             =   2400
            Width           =   3735
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Action :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   178
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Ramasser :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   177
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Courrir :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   176
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Attaque :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   175
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Droite :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   174
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Gauche :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   173
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Bas :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   172
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Haut :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   171
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 14 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   3960
            TabIndex        =   170
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 13 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   169
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 12 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   3960
            TabIndex        =   168
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 11 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   167
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 10 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   3960
            TabIndex        =   166
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 9 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   3960
            TabIndex        =   165
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 8 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   3960
            TabIndex        =   164
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 7 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   3960
            TabIndex        =   163
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 6 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3960
            TabIndex        =   162
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 5 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   161
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 4 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   160
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 3 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   159
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 2 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   158
            Top             =   480
            Width           =   855
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   5040
            TabIndex        =   139
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   5040
            TabIndex        =   138
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   5040
            TabIndex        =   137
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   5040
            TabIndex        =   136
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   5040
            TabIndex        =   135
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   5040
            TabIndex        =   134
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   5040
            TabIndex        =   133
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   5040
            TabIndex        =   132
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   5040
            TabIndex        =   131
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   5040
            TabIndex        =   130
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   5040
            TabIndex        =   129
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   128
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   127
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label key2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   126
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   125
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   124
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   123
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   122
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   121
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   120
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   119
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label key1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifier la touche"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   118
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Raccourci 1 :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   117
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Timer sync 
         Interval        =   250
         Left            =   6720
         Top             =   0
      End
      Begin VB.Frame picParty 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   2985
         Left            =   120
         TabIndex        =   67
         Top             =   4440
         Visible         =   0   'False
         Width           =   2595
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   0
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   80
            Top             =   600
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   81
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   0
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   78
            Top             =   800
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   79
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox Picture15 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            Picture         =   "frmMirage.frx":2F79F0
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   77
            Top             =   2400
            Width           =   270
         End
         Begin VB.PictureBox Picture16 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Picture         =   "frmMirage.frx":2F7C88
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   76
            Top             =   2400
            Width           =   270
         End
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   1
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   74
            Top             =   1275
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   75
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   1
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   72
            Top             =   1485
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   73
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   2
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   70
            Top             =   1995
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   2
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   2
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   68
            Top             =   2205
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   2
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   69
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.Image Image5 
            Height          =   2985
            Left            =   0
            Picture         =   "frmMirage.frx":2F7F13
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   0
            MousePointer    =   5  'Size
            TabIndex        =   88
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   87
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rejoindre/Quitter le groupe"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   0
            TabIndex        =   86
            Top             =   2760
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   85
            Top             =   400
            Width           =   2175
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2040
            TabIndex        =   84
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   83
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6120
         Top             =   0
      End
      Begin VB.PictureBox picquete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4290
         Left            =   480
         Picture         =   "frmMirage.frx":31138D
         ScaleHeight     =   286
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   3345
         Begin VB.TextBox quetetxt 
            Appearance      =   0  'Flat
            Height          =   3015
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label artquete 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1440
            TabIndex        =   17
            Top             =   3840
            Width           =   1845
         End
         Begin VB.Label qf 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3045
            TabIndex        =   20
            Top             =   0
            Width           =   285
         End
         Begin VB.Label av 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   1440
            TabIndex        =   19
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label qt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   18
            Top             =   3600
            Width           =   1020
         End
      End
      Begin VB.PictureBox xp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lexp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape sexp 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox mana 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lmana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CB884B&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape smana 
            BackColor       =   &H00CB884B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox vie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lvie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape svie 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox ObjNm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Label OName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Timer Timer1 
         Left            =   7320
         Top             =   0
      End
      Begin VB.Timer tmrSnowDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8760
         Top             =   0
      End
      Begin VB.Timer tmrRainDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8280
         Top             =   0
      End
      Begin VB.PictureBox ScreenShot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9240
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox txtQ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   5400
         Picture         =   "frmMirage.frx":34028F
         ScaleHeight     =   1545
         ScaleWidth      =   6510
         TabIndex        =   2
         Top             =   7080
         Visible         =   0   'False
         Width           =   6540
         Begin VB.TextBox TxtQ2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "frmMirage.frx":3700E1
            Top             =   240
            Width           =   6285
         End
         Begin VB.Label OK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6000
            TabIndex        =   4
            Top             =   1320
            Width           =   495
         End
      End
   End
   Begin VB.Label lbltimeQuete 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7800
      TabIndex        =   99
      Top             =   9185
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label menu_quete 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   9240
      TabIndex        =   44
      ToolTipText     =   "Quetes"
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label menu_quit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11640
      TabIndex        =   28
      ToolTipText     =   "Quitter"
      Top             =   9240
      Width           =   345
   End
   Begin VB.Label menu_equ 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8760
      TabIndex        =   27
      ToolTipText     =   "Equipements"
      Top             =   9240
      Width           =   345
   End
   Begin VB.Label menu_guild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9840
      TabIndex        =   26
      ToolTipText     =   "Guilde"
      Top             =   9360
      Width           =   405
   End
   Begin VB.Label menu_opt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   25
      ToolTipText     =   "Options"
      Top             =   9360
      Width           =   420
   End
   Begin VB.Label menu_who 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10680
      TabIndex        =   24
      ToolTipText     =   "Qui est en ligne ?"
      Top             =   9360
      Width           =   420
   End
   Begin VB.Label menu_sort 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8280
      TabIndex        =   23
      ToolTipText     =   "Sorts"
      Top             =   9240
      Width           =   465
   End
   Begin VB.Label menu_inv 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7920
      TabIndex        =   22
      ToolTipText     =   "Inventaire"
      Top             =   9240
      Width           =   315
   End
   Begin VB.Image Interface 
      Height          =   900
      Left            =   0
      Picture         =   "frmMirage.frx":3700E7
      Top             =   9120
      Width           =   12000
   End
   Begin WMPLibCtl.WindowsMediaPlayer Mediaplayer 
      Height          =   720
      Left            =   12360
      TabIndex        =   5
      Top             =   4560
      Width           =   480
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   847
      _cy             =   1270
   End
   Begin VB.Menu Tchat 
      Caption         =   "Tchat"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu renommer 
         Caption         =   "Renommer"
      End
      Begin VB.Menu Modifier 
         Caption         =   "Modifier"
      End
   End
End
Attribute VB_Name = "FrmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'1024X768

Private SpellMemorized As Long
Public DragImg As Long

Private OldPCX As Long
Private OldPCY As Long

Private Lon As Long
Private Hau As Long
Public OptionColorSelect As Byte
Private ChatLog As String
Private ChangeTouche As Byte


Private Sub AddDef_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 1 & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 2 & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 3 & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub artquete_Click()
    Player(MyIndex).QueteEnCour = 0
    Accepter = False
    Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
    FrmMirage.picquete.Visible = False
    If quetetimersec.Enabled Then
        quetetimersec.Enabled = False
        tmpsquete.Visible = False
    End If
End Sub

Private Sub cbtr1_Change()

End Sub



Private Sub Canal_Validate(Cancel As Boolean)
Dim C As Integer
TempCommand = ""
 C = InStr(txtMyTextBox.Text, " ")
 If C = 0 Then Exit Sub
 txtMyTextBox.Text = Mid(txtMyTextBox.Text, C, Len(txtMyTextBox.Text) - C + 1)
End Sub

Private Sub cbth_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtb_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtg_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtd_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbta_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtra_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtc_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtac_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtr_keypress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub canaltext_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Call canaltext_LostFocus
If KeyCode = vbKeyReturn Then
    Onglet(OngletActif).Caption = Trim(canaltext.Text)
    Call WriteINI("ONGLET" & OngletActif, "NOM", Trim(canaltext.Text), App.Path & "\Config\Ecriture.ini")
    Call canaltext_LostFocus
End If
End Sub

Private Sub canaltext_LostFocus()
canaltext.Enabled = False
canaltext.Visible = False

Timer3.Enabled = True
FrmMirage.SetFocus
End Sub

Private Sub Chat_Timer()



'RTBChat.SaveFile (App.Path & "\Logs\" & ChatLog & ".rtf")


End Sub

Private Sub chkLowEffect_Click()
    WriteINI "CONFIG", "LowEffect", chkLowEffect.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknobj_Click()
    WriteINI "CONFIG", "NomObjet", chknobj.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.value, App.Path & "\Config\Account.ini"
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
End Sub




Private Sub CmdoptTouche_Click()
Dim i As Byte
For i = 0 To 7
    Call LoadTouches(i)
Next i
For i = 10 To 23
 Call LoadTouches(i)
Next i
pictTouche.Visible = True
End Sub

Private Sub cmdOTA_Click()
Dim i As Byte
    ChangeTouche = 0
    For i = 0 To 7
        key1(i).DataField = vbNull
    Next i
    For i = 0 To 13
        key2(i).DataField = vbNull
    Next i
    pictTouche.Visible = False
End Sub

Private Sub cmdOTO_Click()
Dim i As Byte
    Call WriteINI("TJEU", "haut", key1(0).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "bas", key1(1).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "gauche", key1(2).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "droite", key1(3).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "attaque", key1(4).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "courir", key1(5).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "ramasser", key1(6).DataField, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "action", key1(7).DataField, App.Path & "\Config\Option.ini")
    For i = 0 To 13
        Call WriteINI("TRAC", "rac" & i + 1, key2(i).DataField, App.Path & "\Config\Option.ini")
    Next i
    pictTouche.Visible = False

    
    ChangeTouche = 0
End Sub

Private Sub ColorSave_Click()
Dim i As Byte
Picture19.Visible = False
picOptions.Visible = True

For i = 0 To MaxColor
    MsgRgb(i).r = MsgRgb2(i).r
    MsgRgb(i).g = MsgRgb2(i).g
    MsgRgb(i).B = MsgRgb2(i).B
        WriteINI "canal" & i, "R", STR(MsgRgb(i).r), App.Path & "\Config\Ecriture.ini"
        WriteINI "canal" & i, "G", STR(MsgRgb(i).g), App.Path & "\Config\Ecriture.ini"
        WriteINI "canal" & i, "B", STR(MsgRgb(i).B), App.Path & "\Config\Ecriture.ini"
Next i
End Sub

Private Sub Command1_Click()
picOptions.Visible = False
Call InitAccountOpt
End Sub

Private Sub Command2_Click()
Call Form_Load
End Sub

Private Sub CommandColor_Click()
Dim i As Byte
If Picture19.Visible = False Then
    Picture19.Visible = True
    picOptions.Visible = False
    
End If
For i = 0 To MaxColor
            MsgRgb2(i).r = MsgRgb(i).r
            MsgRgb2(i).g = MsgRgb(i).g
            MsgRgb2(i).B = MsgRgb(i).B
Next i
End Sub



Private Sub ffermer_Click()
fra_fenetre.Visible = False
End Sub

Private Sub Form_GotFocus()
Picsprts.height = 48
On Error Resume Next
txtMyTextBox.SetFocus

End Sub

Private Sub Form_Load()
Dim i As Long, X As Integer, Y As Byte
Dim Ending As String
Dim Qq As Long
    
    
    
    Call LoadChat
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExiste(Rep_Theme & "\Jeu\Text" & Ending) Then txtQ.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\text" & Ending)
        If FileExiste(Rep_Theme & "\info" & Ending) Then FrmMirage.Picture = LoadPNG(App.Path & Rep_Theme & "\info" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\inventaire" & Ending) Then Image3.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\inventaire" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\Carte" & Ending) Then imgCarte.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Carte" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\quitter" & Ending) Then PicMenuQuitter.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\quitter" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\quete" & Ending) Then picquete.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\quete" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\metier" & Ending) Then pictMetier.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\metier" & Ending)
        
    Next i

    Call netbook_change
    
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX
    svie.FillColor = RGB(208, 11, 0)
    smana.FillColor = RGB(208, 11, 0)
    
    'If frmMainMenu.chk_fullscreen.value = Checked Then
        'If (Screen.Height / Screen.TwipsPerPixelY) >= 758 Then txtMyTextBox.Top = 567
        'frmMirage.Height = Screen.Height / Screen.TwipsPerPixelY
        'frmMirage.Width = Screen.Width / Screen.TwipsPerPixelX
        'picScreen.Height = Screen.Height / Screen.TwipsPerPixelY
        'picScreen.Width = Screen.Width / Screen.TwipsPerPixelX
    'End If

    'monnom.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    'monnom.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    txtMyTextBox.Font = ReadINI("POLICE", "PoliceChat", (App.Path & "\Config\Ecriture.ini"))
    
    'Chargement des parametres couleur du tchat
    MaxColor = FrmMirage.OptionColor.count - 1
    ReDim MsgRgb(MaxColor) As RGB
    ReDim MsgRgb2(MaxColor) As RGB
    
    For i = 0 To MaxColor
        MsgRgb(i).r = Val(ReadINI("canal" & i, "R", App.Path & "\Config\Ecriture.ini"))
        MsgRgb(i).g = Val(ReadINI("canal" & i, "G", App.Path & "\Config\Ecriture.ini"))
        MsgRgb(i).B = Val(ReadINI("canal" & i, "B", App.Path & "\Config\Ecriture.ini"))
    FrmMirage.ShapeColor1(i).FillColor = RGB(MsgRgb(i).r, MsgRgb(i).g, MsgRgb(i).B)
    Next i
    
    For Y = 0 To 23
        Call LoadTouches(Y)
    Next Y
    
    fra_fenetre.Visible = False
    Chat.Enabled = True
    ChatLog = Day(Now) & "." & Month(Now) & "." & Year(Now) & "--" & Hour(Now) & "." & Minute(Now) & "." & Second(Now)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If GettingMap Then
        Cancel = True
    Else
        Call GameDestroy
        End
    End If
End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 3
DragX = X
DragY = Y
End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub Interface_Click()
Popup.Visible = False
End Sub


Private Sub key1_Click(Index As Integer)
If ChangeTouche = 0 Then
    key1(Index).Caption = "Appuyez sur une touche"
    ChangeTouche = Index + 1
End If
End Sub

Private Sub key2_Click(Index As Integer)
If ChangeTouche = 0 Then
    key2(Index).Caption = "Appuyez sur une touche"
    ChangeTouche = Index + 10
End If
End Sub

Private Sub Label13_Click()
Call key1_Click(1)
End Sub

Private Sub Label1_Click(Index As Integer)
Dim i As Long
Popup.Visible = False
    For i = 0 To Label1.count - 1
        Label1(i).BackColor = &H0&
    Next i
    Label1(Index).BackColor = &H404040
    Select Case Index
            Case 0
                classement(0).Visible = True
                classement(1).Visible = False
            Case 1
                classement(1).Visible = True
                classement(0).Visible = False
    End Select
            
End Sub

Private Sub Label19_Click()
    If picParty.height >= 2985 / twippy Then picParty.height = 315 / twippy Else picParty.height = 2985 / twippy
End Sub

Private Sub Label20_Click()
Call key1_Click(3)
End Sub

Private Sub Label21_Click()
Call key1_Click(4)
End Sub

Private Sub Label22_Click()
Call key1_Click(4)
End Sub

Private Sub Label27_Click()
    PicMenuQuitter.Visible = False
End Sub

Private Sub Label3_Click()
    If Player(MyIndex).PartyIndex > 0 Then SendLeaveParty: picParty.Visible = False Else picParty.Visible = False
End Sub



Private Sub Label38_Click()
Call key1_Click(5)
End Sub

Private Sub Label39_Click()
Call key1_Click(6)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 6
DragX = X
DragY = Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragImg = 6 Then picParty.Move picParty.Left + ((X / twippx) - (DragX / twippx)), picParty.Top + ((Y / twippy) - (DragY / twippy))
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub Label40_Click()
Call key1_Click(7)
End Sub

Private Sub Label41_Click()
Dim i As Byte
Picture19.Visible = False
picOptions.Visible = True

    For i = 0 To MaxColor
        FrmMirage.ShapeColor1(i).FillColor = RGB(MsgRgb(i).r, MsgRgb(i).g, MsgRgb(i).B)
            MsgRgb2(i).r = MsgRgb(i).r
            MsgRgb2(i).g = MsgRgb(i).g
            MsgRgb2(i).B = MsgRgb(i).B
    Next i
End Sub



Private Sub Label43_Click()
Dim i As Byte

For i = 0 To Check1.count - 1
    WriteINI "ONGLET" & OngletActif, "Canal" & i, Check1(i).value, App.Path & "\Config\Ecriture.ini"
    RTB(OngletActif).Canal(i) = Check1(i).value
Next i
Configurer.Visible = False
End Sub

Private Sub Label7_Click()
Call key1_Click(0)
End Sub

Private Sub Label8_Click()
    picParty.Visible = False
End Sub

Private Sub lblCdP_Click()
    Call SendData("CHANGECHAR" & END_CHAR)
    FrmMirage.Visible = False
    frmMainMenu.Visible = True
    frmMainMenu.PERSONNAGES.Visible = True
    frmsplash.Visible = False
    PicMenuQuitter.Visible = False
End Sub

Private Sub lblDeco_Click()
Dim i As Integer
    Call SendData("CHANGECHAR" & END_CHAR)
    InGame = False
    deco = True
    Sleep 2000
    PicMenuQuitter.Visible = False
    frmMainMenu.Visible = True
    frmMainMenu.PERSONNAGES.Visible = True
    FrmMirage.Visible = False
    FrmMirage.Socket.Close
    FrmMirage.Socket.Connect
End Sub

Private Sub lblendmetier_Click()
pictMetier.Visible = False
End Sub

Private Sub lblmaskinvferm_Click()
fra_fenetre.Visible = False
End Sub

Private Sub lblmaskinvmin_Click()
    If fra_fenetre.height >= 2985 / twippy Then fra_fenetre.height = 315 / twippy Else fra_fenetre.height = 2985 / twippy
End Sub

Private Sub lblmaskinv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 2
DragX = X
DragY = Y
End Sub

Private Sub lblmaskinv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
Call MoveFrame(fra_fenetre, Button, Shift, X, Y)
End Sub

Private Sub lblmaskinv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub



Private Sub lblmaskmenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 1
DragX = X
DragY = Y
End Sub



Private Sub lblmaskmenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub lblMetierApl_Click()
    Call SendData("playermetier" & END_CHAR)
End Sub

Private Sub lblmetierEnd_Click()
    pictMetier.Visible = False
End Sub

Private Sub lblOublierMetier_Click()
    Call SendData("playermetieroublie" & END_CHAR)
    pictMetier.Visible = False
End Sub

Private Sub lblPoints_Change()
    With FrmMirage
        If GetPlayerPOINTS(MyIndex) > 0 Then
            .AddStr.Visible = True
            .AddDef.Visible = True
            .AddSpeed.Visible = True
            .AddMagi.Visible = True
        Else
            .AddStr.Visible = False
            .AddDef.Visible = False
            .AddSpeed.Visible = False
            .AddMagi.Visible = False
        End If
    End With
End Sub

Private Sub lblQuitter_Click()
    Call GameDestroy
    End
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim$(lstOnline.Text) & END_CHAR)
End Sub

Private Sub menu_equ_Click()
If picInv3.Visible = True Then picInv3.Visible = False: PicInterface.Visible = False: Exit Sub

    PicInterface.Visible = True
    picInv3.Visible = True
    Picture13.Visible = False
    Call UpdateVisInv
    
    PrepareSprite (Player(MyIndex).Sprite)
    Picsprts.height = (48)
    Call AffSurfPic(DD_SpriteSurf(Player(MyIndex).Sprite), Picsprts, 0, 0)
    'Call BitBlt(Picsprts.hDC, 0, 0, PIC_X, PIC_Y * PIC_NPC1, Picturesprite.hDC, 3 * PIC_X, Val(Player(MyIndex).Sprite) * (PIC_Y * PIC_NPC1), SRCCOPY)

End Sub

Private Sub menu_guild_Click()

If picOptions.Visible = True Then picOptions.Visible = False
' Set Their Guild Name and Their Rank

If Player(MyIndex).PartyIndex > 0 Then picParty.Visible = True: fra_fenetre.Visible = True
Label3.Visible = picParty.Visible
If picParty.Visible Then
    Dim i As Integer, C As Byte
    If lblPPName(0).Tag <= lblPPName(2).Tag Or lblPPName(2).Caption <> vbNullString Then
        For i = (Val(lblPPName(2).Tag) + 1) To MAX_PLAYERS
            If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C < 3 And i <> MyIndex Then
                C = C + 1
                lblPPName(C - 1).Tag = i
            End If
        Next
        For i = 0 To 2
            lblPPName(i).Visible = (i < C)
            backPPLife(i).Visible = lblPPName(i).Visible
            backPPMana(i).Visible = lblPPName(i).Visible
            If lblPPName(i).Visible Then
                lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
            End If
        Next
    End If
Exit Sub
End If
If Player(MyIndex).Guildaccess > 1 And Player(MyIndex).Guild <> "" Then frmGuild.picGuildAdmin.Visible = True: frmGuild.Show vbModeless, FrmMirage
Call SendData("guildupdate" & END_CHAR)
Exit Sub
End Sub

Private Sub menu_inv_Click()
If picOptions.Visible = True Then picOptions.Visible = False
If fra_fenetre.Visible = True Then fra_fenetre.Visible = False Else fra_fenetre.Visible = True
Call UpdateVisInv
Call ClearPic
End Sub

Private Sub menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To menu.count - 1
        menu(i).ForeColor = &HE0E0E0
    Next i
        menu(Index).ForeColor = &H808080
End Sub

Private Sub menu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Select Case Index
    Case 0 'renommer
        canaltext.Left = Picture20.Left + Onglet(OngletActif).Left
        canaltext.Top = Onglet(OngletActif).Top
        canaltext.Text = Onglet(OngletActif).Caption
        canaltext.Enabled = True
        Popup.Visible = False
        canaltext.Visible = True
        PopupOK = True
        canaltext.SetFocus
        canaltext.SelStart = Len(canaltext.Text)


        
    Case 1 'Configurer
        Configurer.Left = Picture20.Left + Onglet(OngletActif).Left - 20
        Configurer.Top = Onglet(OngletActif).Top + 20
        Configurer.Enabled = True
        Configurer.Visible = True
        Popup.Visible = False
        For i = 0 To FrmMirage.Check1.count - 1
            RTB(OngletActif).Canal(i) = Val(ReadINI("ONGLET" & OngletActif, "Canal" & i, App.Path & "\Config\Ecriture.ini"))
            Check1(i).value = Val(ReadINI("ONGLET" & OngletActif, "Canal" & i, App.Path & "\Config\Ecriture.ini"))
        Next i
    
    Case 2 'annuler
    Popup.Visible = False
    
    Case 3
        RTBChat(OngletActif).Text = ""

End Select
End Sub

Private Sub menu_opt_Click()
    If picquete.Visible = True Then picquete.Visible = False
    
    
    If picOptions.Visible = False Then
        picOptions.Visible = True
    Else
        picOptions.Visible = False
    End If
    

End Sub

Private Sub menu_quete_Click()

If picOptions.Visible = True Then picOptions.Visible = False


If FrmMirage.picquete.Visible = True Then
    FrmMirage.picquete.Visible = False
Else

    If Player(MyIndex).QueteEnCour > 0 Then
        Call ClearPic
        fra_fenetre.Visible = False
        FrmMirage.picquete.Visible = True
        FrmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description
    Else
        Call ClearPic
        fra_fenetre.Visible = False
        FrmMirage.picquete.Visible = True
        FrmMirage.quetetxt.Text = "Pas de quête en cours..."
    End If
    
End If
End Sub

Private Sub menu_quit_Click()
'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
'Call GameDestroy
If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
End Sub

Private Sub menu_sort_Click()
If Picture13.Visible = True Then Picture13.Visible = False: PicInterface.Visible = False: Exit Sub

PicInterface.Visible = True
Picture13.Visible = True

Call ClearPic
Call SendData("spells" & END_CHAR)
End Sub


Private Sub menu_who_Click()
If lstOnline.Visible = True Then lstOnline.Visible = False: PicInterface.Visible = False: Exit Sub

PicInterface.Visible = True
lstOnline.Visible = True
    Call SendOnlineList
End Sub

Private Sub OK_Click()
Dim i As Long
Dim msgb As String

If Player(MyIndex).QueteEnCour > 0 And Accepter = False Then
    msgb = MsgBox("Voulez-vous faire la quête proposée ?", vbYesNo, "Quete")
        If msgb = vbYes Then
            Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
            Accepter = True
        Else
            Player(MyIndex).QueteEnCour = 0
            Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
            Accepter = False
        End If
End If
txtQ.Visible = False
End Sub

Private Sub Onglet_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long


If Button = 2 Then
If Index = 0 Then Popup.Visible = False: Exit Sub
   Popup.Left = picScreen.Width + Onglet(Index).Left + (X / twippx)
   If (picScreen.Width + Onglet(Index).Left + (X / twippx) + Onglet(Index).Width) > Picture20.Width Then Popup.Left = (picScreen.Width + Onglet(Index).Left + (X / twippx)) - Onglet(Index).Width
   Popup.Top = Onglet(Index).Top + (Y / twippy) + 10
   Popup.Visible = True
   Onglet(Index).ForeColor = vbWhite
   OngletActif = Index
ElseIf Button = 1 Then
    For i = 0 To Onglet.count - 1
        Onglet(i).BackColor = &H0&
    Next i
    Onglet(Index).BackColor = &H400000
    OngletActif = Index
    Onglet(Index).ForeColor = &HE0E0E0
    Popup.Visible = False
    Call ShowChat(Index)
End If
End Sub

Private Sub OptionColor_Click(Index As Integer)
Dim i As Byte
For i = 0 To MaxColor
    If i = Index Then
        OptionColor(i).value = True
        OptionColorSelect = i
    Else
        OptionColor(i).value = False
    End If
Next i
End Sub



Private Sub PicColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo trap:
Dim value As Long


If Button = 1 Then
    With MsgRgb2(OptionColorSelect)

        value = PicColor.point(X, Y)
        .r = value Mod 256
        .g = Int(value / 256) Mod 256
        .B = Int(value / 256 / 256) Mod 256
        FrmMirage.ShapeColor1(OptionColorSelect).FillColor = RGB(.r, .g, .B)
    End With
Exit Sub
End If
trap: Exit Sub
End Sub

Private Sub picInv_DblClick(Index As Integer)
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Or Player(MyIndex).Inv(Inventory).num > MAX_ITEMS Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
    End If
Next d
Call UpdateVisInv
End Sub


Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index > MAX_INV - 1 Then Exit Sub
    Inventory = Index + 1
    FrmMirage.SelectedItem.Top = FrmMirage.picInv(Inventory - 1).Top - 1
    FrmMirage.SelectedItem.Left = FrmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        If Player(MyIndex).Inv(Inventory).num <= 0 Or Player(MyIndex).Inv(Inventory).num > MAX_ITEMS Then
            dragAndDropT = 0
            dragAndDrop = 0
            IDAD.Visible = False
        Else
            If dragAndDrop = Inventory Then
                dragAndDrop = 0
                dragAndDropT = 0
                IDAD.Visible = False
            Else
                dragAndDrop = Inventory
                dragAndDropT = 2
                IDAD.Top = FrmMirage.picInv(Inventory - 1).Top - 1
                IDAD.Left = FrmMirage.picInv(Inventory - 1).Left - 1
                IDAD.Visible = True
            End If
        End If
    ElseIf Button = 3 Then
        Call DropItems
    End If
End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Long
d = Index
If Index > MAX_INV - 1 Then Exit Sub
    If Player(MyIndex).Inv(d + 1).num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable <> 0 Then
            descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipé)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipé)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipé)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipé)"
            ElseIf Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable <> 0 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            Else
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name)
            End If
        End If
        
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Force"
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Défense"
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Vitesse"
        descHpMp.Caption = "PV: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " PM: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " End: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "FOR: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magie: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Vitesse: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1 <= 0 Then
            Usure2.Width = descName.Width
        Else
            Usure2.Width = descName.Width * (100 * (GetPlayerInvItemDur(MyIndex, d + 1) / Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1))
        End If
        
        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        descName.ForeColor = Item(GetPlayerInvItemNum(MyIndex, d + 1)).NCoul
    Else
     descName.Caption = ""
     descStr.Caption = ""
     descDef.Caption = ""
     descSpeed.Caption = ""
     descHpMp.Caption = ""
     descSD.Caption = ""
     descMS.Caption = ""
     desc.Caption = ""
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long
Dim PX As Long
Dim PY As Long
Dim Cod As String
Dim tp As Long
    If ChangeTouche > 0 Then
        If KeyCode = vbKeyEscape Then Call LoadTouches(ChangeTouche): ChangeTouche = 0: Exit Sub
        Call ChangerTouche(KeyCode, ChangeTouche): ChangeTouche = 0: Exit Sub 'merde c'est naze ce que j'ai fait
    End If
        
        
    If ConOff = True Or Paralyse = True Then Exit Sub

    Call CheckInput(0, KeyCode, Shift)
    
    If (FrmMirage.txtMyTextBox.Visible = False) And (KeyCode = Raccourcit(7)) Then
        PX = 0
        PY = 0
        If Player(MyIndex).Y - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_UP Then PX = 0: PY = -1
        End If
                
        If Player(MyIndex).Y + 1 < MAX_MAPY + 1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_DOWN Then PX = 0: PY = 1
        End If
                
        If Player(MyIndex).X - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_LEFT Then PX = -1: PY = 0
        End If
        
        If Player(MyIndex).X + 1 < MAX_MAPX + 1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_RIGHT Then PX = 1: PY = 0
        End If
        
        If PX <> 0 Or PY <> 0 Then
        With Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY)
            If .String1 > vbNullString And TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                Dim Packet As String
                Cod = InputBox("Veuillez entre le mot de passe :", "Code")
                If Cod = .String1 Then
                    TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = YES
                    Packet = "OUVRIRE" & SEP_CHAR & GetPlayerX(MyIndex) + PX & SEP_CHAR & GetPlayerY(MyIndex) + PY & END_CHAR
                    Call SendData(Packet)
                    If .Type = TILE_TYPE_COFFRE Then
                        i = FindOpenInvSlot(Val(.Data3))
                        If i > 0 Then
                            Call SetPlayerInvItemNum(MyIndex, i, Val(.Data3))
                            Call SetPlayerInvItemValue(MyIndex, i, GetPlayerInvItemValue(MyIndex, i) + 1)
                            Call SetPlayerInvItemDur(MyIndex, i, Item(Val(.Data3)).Data1)
                            Call UpdateVisInv
                            Packet = "ACOFFRE" & SEP_CHAR & i & SEP_CHAR & Val(.Data3) & SEP_CHAR & 1 & SEP_CHAR & Item(Val(.Data3)).Data1 & END_CHAR
                            Call SendData(Packet)
                        End If
                    End If
                Else
                    Call MsgBox("Mauvais code.", vbCritical)
                End If
            End If
        End With
        End If
        
        If GetPlayerY(MyIndex) - 1 > 0 And GetPlayerY(MyIndex) - 1 < MAX_MAPY Then
            With Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1)
            If .Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                If Trim$(.String1) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String1))
                If Trim$(.String2) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String2))
                If Trim$(.String3) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String3))
                Exit Sub
            End If
            End With
        End If
    End If
    
    If txtMyTextBox.Visible = False Then
    If KeyCode = Raccourcit(10) Then Call useRac(0)
    If KeyCode = Raccourcit(11) Then Call useRac(1)
    If KeyCode = Raccourcit(12) Then Call useRac(2)
    If KeyCode = Raccourcit(13) Then Call useRac(3)
    If KeyCode = Raccourcit(14) Then Call useRac(4)
    If KeyCode = Raccourcit(15) Then Call useRac(5)
    If KeyCode = Raccourcit(16) Then Call useRac(6)
    If KeyCode = Raccourcit(17) Then Call useRac(7)
    If KeyCode = Raccourcit(18) Then Call useRac(8)
    If KeyCode = Raccourcit(19) Then Call useRac(9)
    If KeyCode = Raccourcit(20) Then Call useRac(10)
    If KeyCode = Raccourcit(21) Then Call useRac(11)
    If KeyCode = Raccourcit(22) Then Call useRac(12)
    If KeyCode = Raccourcit(23) Then Call useRac(13)
    End If
    If KeyCode = vbKeyEscape And PopupOK = False Then
        If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
    End If
    
    If KeyCode = vbKeyF1 And (FrmMirage.txtMyTextBox.Visible = False) Then Call menu_inv_Click: Exit Sub '"i" ouvre l'inventaire
    If KeyCode = vbKeyF2 And (FrmMirage.txtMyTextBox.Visible = False) Then Call menu_sort_Click: Exit Sub '"m" ouvre le pannel magie
    If KeyCode = vbKeyF3 And (FrmMirage.txtMyTextBox.Visible = False) Then Call menu_equ_Click: Exit Sub   '"e" ouvre l'equipement
    If KeyCode = vbKeyF4 And (FrmMirage.txtMyTextBox.Visible = False) Then Call menu_quete_Click: Exit Sub   '"m" ouvre le pannel magie
    
    
    
    ' The Guild Handler
    If KeyCode = vbKeyF5 Then
        If Player(MyIndex).Guildaccess > 1 And Player(MyIndex).Guild <> "" Then frmGuild.picGuildAdmin.Visible = True: frmGuild.Show vbModeless, FrmMirage
        Call SendData("guildupdate" & END_CHAR)
        Exit Sub
    End If
    
    'quete desc
    If KeyCode = vbKeyF7 Then
        If Player(MyIndex).QueteEnCour > 0 Then Call ClearPic: fra_fenetre.Visible = False: FrmMirage.picquete.Visible = True: FrmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description Else Call ClearPic
    End If
    
    If KeyCode = vbKeyF8 Then frmPlayerHelp.Show
    
    If KeyCode = vbKeyF9 Then If Player(MyIndex).Access > 0 Then frmadmin.Show
    
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & SpellMemorized & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Vous ne pouvez lancer un sort en marchant.", BrightRed)
                End If
            End If
        Else
            Call AddText("Aucune magie mémoriser.", BrightRed)
        End If
    Else
        Call CheckInput(0, KeyCode, Shift)
    End If
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(FrmMirage)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(FrmMirage, 8, 6, 634, 479)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyEnd Then
    d = GetPlayerDir(MyIndex)
    
        If Player(MyIndex).Moving = NO Then
            If Player(MyIndex).Dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
                If d <> DIR_LEFT Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
                If d <> DIR_UP Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
                If d <> DIR_RIGHT Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
                If d <> DIR_DOWN Then Call Sendplayerdir
            End If
        End If
    End If
End Sub
Private Sub PicOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragX = X
    DragY = Y
End Sub

Private Sub PicOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(FrmMirage.picOptions, Button, Shift, X, Y)
End Sub



Private Sub picquete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 5
DragX = X
DragY = Y
End Sub

Private Sub picquete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
Call MovePicture(picquete, Button, Shift, X, Y)
'If DragImg = 5 Then DoEvents: If DragImg = 5 Then picquete.Top = picquete.Top + ((y / twippy) - (DragY / twippy)): picquete.Left = picquete.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub picquete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub picRac_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Q As Long
Dim Qq As Long
Dim d As Byte
    If Button = 1 Then
        Call useRac(Index)
    End If
    If Button = 2 Then
        If dragAndDrop > 0 Then
            rac(Index, 0) = dragAndDrop
            rac(Index, 1) = dragAndDropT
        End If
        Call saveRac
    End If
    dragAndDropT = 0
    dragAndDrop = 0
    SDAD.Visible = False
    IDAD.Visible = False
End Sub

Private Sub picScreen_Click()
Popup.Visible = False
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlayerSearch(Button, Shift, (X + NewPlayerPicX), (Y + NewPlayerPicY))
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = ((X + NewPlayerPicX) \ 32)
CurY = ((Y + NewPlayerPicY) \ 32)
PotX = X
PotY = Y

If CurX <> OldPCX Or CurY <> OldPCY Then Call CaseChange(CurX, CurY): OldPCX = CurX: OldPCY = CurY
End Sub

Private Sub picspell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Player(MyIndex).Spell(Index + 1) > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & Index + 1 & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Vous ne pouvez lancer un sort en marchant.", BrightRed)
                End If
            End If
        Else
            Call AddText("Aucuns sort ici.", BrightRed)
        End If
    End If
    If Button = 2 Then
        If Player(MyIndex).Spell(Index + 1) > 0 Then
            If dragAndDrop = Index + 1 Then
                dragAndDrop = 0
                dragAndDropT = 0
                SDAD.Visible = False
            Else
                dragAndDrop = Index + 1
                dragAndDropT = 1
                SDAD.Top = picspell(Index).Top - 1
                SDAD.Left = picspell(Index).Left - 1
                SDAD.Visible = True
            End If
        Else
            dragAndDropT = 0
            dragAndDrop = 0
            SDAD.Visible = False
        End If
    End If
End Sub

Private Sub Picture15_Click()
    If Player(MyIndex).PartyIndex > 0 Then
        Dim i As Integer, C As Byte
        If lblPPName(0).Tag <= lblPPName(2).Tag And lblPPName(2).Caption <> vbNullString Then
            For i = (Val(lblPPName(2).Tag) + 1) To MAX_PLAYERS
                If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C < 3 And i <> MyIndex Then
                    C = C + 1
                    lblPPName(C - 1).Tag = i
                End If
            Next
            For i = 0 To 2
                lblPPName(i).Visible = (i < C)
                backPPLife(i).Visible = lblPPName(i).Visible
                backPPMana(i).Visible = lblPPName(i).Visible
                If lblPPName(i).Visible Then
                    lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                    shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                    shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
                    lblPPLife(i).Caption = "PV : " & Player(Val(lblPPName(i).Tag)).HP & "/" & Player(Val(lblPPName(i).Tag)).MaxHp
                    lblPPMana(i).Caption = "PM : " & Player(Val(lblPPName(i).Tag)).MP & "/" & Player(Val(lblPPName(i).Tag)).MaxMp
                End If
            Next
        End If
    Else: picParty.Visible = False: End If
End Sub

Private Sub Picture16_Click()
    If Player(MyIndex).PartyIndex > 0 Then
        Dim i As Integer, C As Byte
        C = 3
        For i = (Val(lblPPName(0).Tag) - 1) To 1 Step -1
            If i > 0 Then
                If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C > 0 And i <> MyIndex Then
                    C = C - 1
                    lblPPName(C).Tag = i
                End If
            End If
        Next
        For i = 0 To 2
            lblPPName(i).Visible = (i <= Abs(C - 3))
            backPPLife(i).Visible = lblPPName(i).Visible
            backPPMana(i).Visible = lblPPName(i).Visible
            If lblPPName(i).Visible Then
                lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
                lblPPLife(i).Caption = "PV : " & Player(Val(lblPPName(i).Tag)).HP & "/" & Player(Val(lblPPName(i).Tag)).MaxHp
                lblPPMana(i).Caption = "PM : " & Player(Val(lblPPName(i).Tag)).MP & "/" & Player(Val(lblPPName(i).Tag)).MaxMp
            End If
            
            'lblPPName(i).Visible = True: backPPLife(i).Visible = True: backPPMana(i).Visible = True
            'lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).Level
            'shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
            'shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
        Next
    Else: picParty.Visible = False: End If
End Sub

Private Sub Picture17_Click()
    Picture11.Top = Picture11.Top + 88
End Sub

Private Sub Picture18_Click()
    Picture11.Top = Picture11.Top - 88
    If Picture11.Top > 0 Then Picture11.Top = 0
End Sub

Private Sub PicColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim value As Long



With MsgRgb2(OptionColorSelect)

value = PicColor.point(X, Y)
.r = value Mod 256
.g = Int(value / 256) Mod 256
.B = Int(value / 256 / 256) Mod 256
FrmMirage.ShapeColor1(OptionColorSelect).FillColor = RGB(.r, .g, .B)
End With
End Sub

Private Sub Picture19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragX = X
    DragY = Y
End Sub

Private Sub Picture19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MovePicture(FrmMirage.Picture19, Button, Shift, X, Y)
End Sub

Private Sub Picture20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Popup.Visible = False
'DragX = x
'DragY = y
End Sub


Private Sub Picture21_Click()
            ConOff = True
            Call SendData("refresh" & END_CHAR)
End Sub


Private Sub Popup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
    For i = 0 To menu.count - 1
        menu(i).ForeColor = &HE0E0E0
    Next i
End Sub

Private Sub qf_Click()
picquete.Visible = False
End Sub

Private Sub quetetimersec_Timer()
Dim Queten As Long

Queten = Val(Player(MyIndex).QueteEnCour)
If Queten <= 0 Then Exit Sub
If quete(Queten).Temps > 0 And Player(MyIndex).QueteEnCour > 0 Then

Seco = Seco - 1
If Seco <= 0 And Minu > 0 Then
    Seco = 59
    seconde.Caption = Seco
    Minu = Minu - 1
    If Len(STR$(Minu)) > 2 Then minutes.Caption = Minu & ":" Else minutes.Caption = "0" & Minu & ":"
End If
If Seco <= 0 And Minu <= 0 Then
    seconde.Caption = 0
    Call MsgBox("La quête : " & Trim$(quete(Queten).nom) & " est terminer, le temps est écoulé")
    Player(MyIndex).QueteEnCour = 0
    quetetimersec.Enabled = False
    tmpsquete.Visible = False
End If

If Len(STR$(Seco)) > 2 Then seconde.Caption = Seco Else seconde.Caption = "0" & Seco
lbltimeQuete.Visible = True
lbltimeQuete.Caption = "Quête se termine dans :" & Minu & " minute(s) et " & Seco & " seconde."
Else
Player(MyIndex).QueteEnCour = 0
tmpsquete.Visible = False
quetetimersec.Enabled = False
lbltimeQuete.Visible = False
End If

End Sub

Private Sub RTBChat_Click(Index As Integer)
Popup.Visible = False
End Sub

Private Sub RTBChat_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call HandleKeypresses(KeyCode)
End If
End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    MAX_BLT_LINE = scrlBltText.value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "Nbr de ligne écrite sur l'écran: " & scrlBltText.value
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingData(bytesTotal)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then KeyAscii = 0
    
    If (FrmMirage.txtMyTextBox.Visible = False) And (KeyAscii = Raccourcit(7)) Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        If fra_fenetre.Visible = True Then fra_fenetre.Visible = False
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ConOff = True Or Paralyse = True Then Exit Sub
    'If PopupOK = True Then Exit Sub
    
    Call CheckInput(1, KeyCode, Shift)
    On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub sync_Timer()
SendData ("sync" & END_CHAR)
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
If Mediaplayer.URL > vbNullString Then
    If Mediaplayer.Controls.currentPosition = 0 And Mediaplayer.currentMedia.name = Mid$(Map(GetPlayerMap(MyIndex)).Music, 1, Len(Map(GetPlayerMap(MyIndex)).Music) - 4) Then Call FrmMirage.Mediaplayer.Controls.Play
End If
End Sub

Private Sub Timer2_Timer()
    Call affrac
    'Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
PopupOK = False
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then tmrRainDrop.Enabled = False: Exit Sub
    If BLT_RAIN_DROPS > 0 Then If DropRain(BLT_RAIN_DROPS).Randomized = False Then Call RNDRainDrop(BLT_RAIN_DROPS)
    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1
    If tmrRainDrop.Interval > 30 Then tmrRainDrop.Interval = tmrRainDrop.Interval - 10
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then tmrSnowDrop.Enabled = False: Exit Sub
    If BLT_SNOW_DROPS > 0 Then If DropSnow(BLT_SNOW_DROPS).Randomized = False Then Call RNDSnowDrop(BLT_SNOW_DROPS)
    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1
    If tmrSnowDrop.Interval > 30 Then tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
End Sub

Private Sub lblUseItem_Click()
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
Next d

Call UpdateVisInv
End Sub

Private Sub lblDropItem_Click()
    Call DropItems
End Sub

Sub DropItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done

If Inventory <= 0 Then Exit Sub
InvNum = Inventory
   
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Empilable <> 0 Then
            GoldAmount = InputBox("Combien de " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") voulez vous jeter?", "Jeter " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, FrmMirage.Left, FrmMirage.Top)
            If IsNumeric(GoldAmount) Then Call SendDropItem(InvNum, GoldAmount)
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then MsgBox "Trop grande quantiter(erreur du logiciel)"
End Sub

Private Sub txtMyTextBox_Click()
Popup.Visible = False
End Sub

Private Sub txtQ_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub txtQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 4
DragX = X
DragY = Y
End Sub

Private Sub txtQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragImg = 4 Then txtQ.Top = txtQ.Top + ((Y / twippy) - (DragY / twippy)): txtQ.Left = txtQ.Left + ((X / twippx) - (DragX / twippx))
End Sub

Private Sub txtQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub TxtQ2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub txtTempsBulles_Change()
If IsNumeric(txtTempsBulles.Text) Then
WriteINI "CONFIG", "bubbletime", txtTempsBulles, App.Path & "\Config\Client.ini"
End If
End Sub
Public Sub ClearPic()
    picquete.Visible = False
    picEquip.Visible = False
End Sub
