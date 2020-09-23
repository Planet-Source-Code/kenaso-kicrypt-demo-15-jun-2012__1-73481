VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   11250
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11250
   Visible         =   0   'False
   Begin VB.Frame fraMain 
      Height          =   5820
      Index           =   1
      Left            =   90
      TabIndex        =   30
      Top             =   1440
      Width           =   11070
      Begin VB.TextBox txtKeyMix 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7200
         MaxLength       =   1
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5340
         Width           =   330
      End
      Begin VB.CheckBox chkExtraInfo 
         Caption         =   "Extra instructions"
         Height          =   195
         Left            =   9000
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   900
         Width           =   1875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy to clipboard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7095
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   5220
         Width           =   1770
      End
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   225
         MaxLength       =   60
         TabIndex        =   5
         Text            =   "txtPwd"
         Top             =   5340
         Width           =   6840
      End
      Begin VB.Frame fraEncrypt 
         Height          =   4500
         Index           =   1
         Left            =   8955
         TabIndex        =   37
         Top             =   1170
         Width           =   1995
         Begin VB.ComboBox cboBlockLength 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030A
            Left            =   135
            List            =   "frmMain.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2970
            Width           =   1695
         End
         Begin VB.ComboBox cboKeyLength 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030E
            Left            =   135
            List            =   "frmMain.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2250
            Width           =   1695
         End
         Begin VB.ComboBox cboHash 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":0312
            Left            =   135
            List            =   "frmMain.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1350
            Width           =   1695
         End
         Begin VB.ComboBox cboEncrypt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   450
            Width           =   1695
         End
         Begin VB.ComboBox cboRounds 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":0316
            Left            =   1080
            List            =   "frmMain.frx":0318
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3510
            Width           =   750
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Block Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   43
            Top             =   2745
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Press GO button and you will be prompted."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   6
            Left            =   225
            TabIndex        =   42
            Top             =   4005
            Width           =   1575
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Password Key Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   195
            TabIndex        =   41
            Top             =   1785
            Width           =   1530
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Encryption Algorithm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   40
            Top             =   180
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Password Hash Algorithm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   150
            TabIndex        =   39
            Top             =   900
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Rounds of encryption"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   5
            Left            =   165
            TabIndex        =   38
            Top             =   3465
            Width           =   855
         End
      End
      Begin VB.Frame fraEncrypt 
         Height          =   4980
         Index           =   0
         Left            =   105
         TabIndex        =   33
         Top             =   105
         Width           =   8745
         Begin VB.TextBox txtInputString 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "frmMain.frx":031A
            Top             =   360
            Width           =   8520
         End
         Begin VB.TextBox txtInputFile 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "txtInputFile"
            Top             =   360
            Visible         =   0   'False
            Width           =   7890
         End
         Begin VB.PictureBox picProgressBar 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   8460
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   4545
            Width           =   8520
         End
         Begin VB.TextBox txtOutput 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "frmMain.frx":032B
            Top             =   3900
            Width           =   8520
         End
         Begin VB.CommandButton cmdBrowse 
            Height          =   375
            Left            =   8100
            Picture         =   "frmMain.frx":0335
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   36
            Top             =   3660
            Width           =   5820
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   35
            Top             =   135
            Width           =   5820
         End
      End
      Begin VB.PictureBox picDataType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   8985
         ScaleHeight     =   645
         ScaleWidth      =   1920
         TabIndex        =   31
         Top             =   225
         Width           =   1920
         Begin VB.OptionButton optDataType 
            Caption         =   "String data"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   330
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optDataType 
            Caption         =   "File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1275
            TabIndex        =   7
            Top             =   330
            Width           =   645
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   495
            TabIndex        =   32
            Top             =   45
            Width           =   990
         End
      End
      Begin VB.Label lblKeyMix 
         Caption         =   "Number of rounds to mix primary key Min - 1     Max - 5"
         Height          =   555
         Left            =   7620
         TabIndex        =   50
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label lblPwd 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   285
         TabIndex        =   44
         Top             =   5145
         Width           =   7425
      End
   End
   Begin VB.Frame fraMain 
      Height          =   5820
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   1440
      Width           =   10035
      Begin VB.ComboBox cboRandom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   225
         Width           =   2985
      End
      Begin VB.TextBox txtRandom 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0437
         Top             =   630
         Width           =   9825
      End
      Begin VB.Label lblAlgo 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Data Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   5280
         TabIndex        =   29
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   2
      Left            =   9765
      Picture         =   "frmMain.frx":0441
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Display credits"
      Top             =   7335
      Width           =   640
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   4463
      Picture         =   "frmMain.frx":074B
      ScaleHeight     =   495
      ScaleWidth      =   2325
      TabIndex        =   22
      Top             =   15
      Width           =   2325
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   1
      Left            =   9045
      Picture         =   "frmMain.frx":0C21
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7335
      Width           =   640
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   10440
      Picture         =   "frmMain.frx":1063
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   300
      Picture         =   "frmMain.frx":136D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   180
      Width           =   480
   End
   Begin MSComDlg.CommonDialog cdFileOpen 
      Left            =   8460
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   3
      Left            =   10485
      Picture         =   "frmMain.frx":1677
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Terminate this application"
      Top             =   7335
      Width           =   640
   End
   Begin VB.CommandButton cmdChoice 
      Height          =   640
      Index           =   0
      Left            =   9045
      Picture         =   "frmMain.frx":1981
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7335
      Width           =   640
   End
   Begin VB.Frame fraChoice 
      Height          =   690
      Left            =   90
      TabIndex        =   18
      Top             =   735
      Width           =   5715
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   75
         ScaleHeight     =   465
         ScaleWidth      =   5535
         TabIndex        =   24
         Top             =   150
         Width           =   5535
         Begin VB.OptionButton optChoice 
            Caption         =   "Encrypt/Decrypt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   75
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "CRC-32"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   1
            Top             =   75
            Width           =   1035
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Hash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3150
            TabIndex        =   2
            Top             =   75
            Width           =   810
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Random data"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4185
            TabIndex        =   3
            Top             =   75
            Width           =   1260
         End
      End
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6720
      TabIndex        =   25
      Top             =   915
      Width           =   4320
   End
   Begin VB.Label lblEncryptMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "lblEncryptMsg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   225
      TabIndex        =   23
      Top             =   7425
      Width           =   5715
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5010
      TabIndex        =   19
      Top             =   495
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       frmMain
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Used clsAPI_Hash class to replace cMD4, cMD5, cSHA1 and
'                cSHA2 classes.
'              - Removed RipeMD classes because they are considered weak.
' 12-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug with selecting number of encryption rounds.
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
'              See ProgressBar() routine.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made updates as per Joe Sova's suggestions
'                1. Locked and unlocked input/output textboxes during
'                   processing to reduce flicker.
'                2. Hide progressbar during string processing.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated Cypher_Processing() routine to reference new property
'              CreateNewFile(). If you want to overwrite file being encrypted
'              or decrypted.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

' ***************************************************************************
' Module API Declares
' ***************************************************************************
  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

  ' Reduce flicker while loading a control
  ' Lock the control to prevent redrawing
  '     Syntax:  LockWindowUpdate frmMain.hWnd
  ' Unlock the control
  '     Syntax:  LockWindowUpdate 0&
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hwnd As Long) As Long

' ***************************************************************************
' Module Variables
'
' Variable name:     mlngCipher
' Naming standard:   m lng Cipher
'                    - --- ---------
'                    |  |    |______ Variable subname
'                    |  |___________ Data type (Long)
'                    |______________ Module level designator
'
' ***************************************************************************
  Private mlngKeyMix        As Long
  Private mlngRounds        As Long
  Private mlngHashAlgo      As Long
  Private mlngKeyLength     As Long
  Private mlngCipherAlgo    As Long
  Private mlngBlockLength   As Long
  Private mlngPwdLength_Min As Long
  Private mlngPwdLength_Max As Long
  Private mstrFolder        As String
  Private mstrFilename      As String
  Private mstrPassword      As String
  Private mblnHash          As Boolean
  Private mblnPrng          As Boolean
  Private mblnCRC32         As Boolean
  Private mblnCipher        As Boolean
  Private mblnLoading       As Boolean
  Private mblnStringData    As Boolean
  Private mblnHashLowercase As Boolean
  Private mblnCreateNewFile As Boolean
  Private mobjRandom        As kiCrypt.cPrng
  Private mobjKeyEdit       As cKeyEdit

  Private WithEvents mobjHash   As kiCrypt.cHash
Attribute mobjHash.VB_VarHelpID = -1
  Private WithEvents mobjCRC32  As kiCrypt.cCRC32
Attribute mobjCRC32.VB_VarHelpID = -1
  Private WithEvents mobjCipher As kiCrypt.cCipher
Attribute mobjCipher.VB_VarHelpID = -1

Private Sub cboEncrypt_Click()
    
    Dim lngIdx As Long
    
    If mblnLoading Then
        Exit Sub
    End If
                
    mlngCipherAlgo = cboEncrypt.ListIndex
    
    Select Case mlngCipherAlgo
           Case 1   ' Base64
                lblEncryptMsg.Visible = False
                lblPwd.Visible = False
                txtPwd.Visible = False
                    
                ' Prepare combo boxes
                cboHash.Enabled = False
                cboHash.ForeColor = vbGrayText
                cboKeyLength.Enabled = False
                cboKeyLength.ForeColor = vbGrayText
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = False
                cboRounds.ForeColor = vbGrayText
                
                mlngKeyLength = 0
                mlngBlockLength = 0
                mlngRounds = 0
           
           Case 0, 3, 5, 6 ' ArcFour, GOST, Serpent, Skipjack
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 128 To 416 Step 32
                        .AddItem CStr(lngIdx) & " bits"   ' 32 bit increments
                    Next lngIdx
                    
                    For lngIdx = 448 To 1024 Step 64
                        .AddItem CStr(lngIdx) & " bits"   ' 64 bit increments
                    Next lngIdx
                    .ListIndex = 0
                End With
    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboRounds_Click
                        
           Case 2, 7  ' Blowfish, TwoFish
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 32 To 448 Step 32
                        .AddItem CStr(lngIdx) & " bits"   ' 32 bit increments
                    Next lngIdx
                    .ListIndex = 0
                End With
    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboRounds_Click
                
           Case 4  ' Rijndael
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 128 To 256 Step 32
                        .AddItem CStr(lngIdx) & " bits"
                    Next lngIdx
                    .ListIndex = 0
                End With
                    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = True
                cboBlockLength.ForeColor = vbBlack
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboBlockLength_Click
                cboRounds_Click
    End Select
                
    Select Case mlngCipherAlgo
    
           Case 2, 3   ' Blowfish, GOST
                lblKeyMix.Enabled = True
                lblKeyMix.Visible = True
                txtKeyMix.Enabled = True
                txtKeyMix.Visible = True
                txtKeyMix.Text = mlngKeyMix
                
           Case Else  ' Hide textbox and label
                lblKeyMix.Enabled = False
                lblKeyMix.Visible = False
                txtKeyMix.Enabled = False
                txtKeyMix.Visible = False
    End Select
                
End Sub

Private Sub cboHash_Click()

    Dim lngIdx As Long
    
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    mlngHashAlgo = cboHash.ListIndex
    
    ' If performing encryption then leave
    If cboEncrypt.Enabled Then
        Exit Sub
    End If
    
    ' Multiple hash rounds only available
    ' for testing hash output values
    Select Case mlngHashAlgo
           
           ' MD2, MD4, MD5, SHA-1, SHA-256, SHA-384, SHA-512
           ' Whirlpool-224, Whirlpool-256, Whirlpool-384, Whirlpool-512
           Case 0 To 6, 14 To 17
                With cboRounds
                    .Clear
                    For lngIdx = 1 To 10
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 0  ' Default rounds = 1
                End With
                    
           Case 7 To 13    ' Tiger family
                With cboRounds
                    .Clear
                    For lngIdx = 3 To 15
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 3   ' Default rounds = 6
                End With
    End Select
    
    mlngRounds = CLng(Trim$(Left$(cboRounds.Text, 2)))
    
End Sub

Private Sub cboKeyLength_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    ' Select user defined key length
    Select Case mlngCipherAlgo
           Case eCIPHER_BASE64
                mlngKeyLength = 0
                
           Case eCIPHER_BLOWFISH, eCIPHER_TWOFISH
                mlngKeyLength = CLng(Trim$(Left$(cboKeyLength.Text, 3)))
                
           Case Else
                mlngKeyLength = CLng(Trim$(Left$(cboKeyLength.Text, 4)))
    End Select
    
End Sub

Private Sub cboBlockLength_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    ' Select user defined block size (RIJNDAEL only)
    Select Case mlngCipherAlgo
           Case eCIPHER_RIJNDAEL: mlngBlockLength = CLng(Trim$(Left$(cboBlockLength.Text, 3)))
           Case Else:             mlngBlockLength = 0
    End Select
    
End Sub

Private Sub cboRounds_Click()

    If mblnLoading Then
        Exit Sub
    End If
    
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False

    If cboEncrypt.Enabled Then
        
        ' Select number of rounds for encryption
        Select Case mlngCipherAlgo
               Case eCIPHER_BASE64: mlngRounds = 0
               Case Else:           mlngRounds = CLng(Trim$(Left$(cboRounds.Text, 2)))
        End Select
    
    Else
        ' Perform hash testing only
        mlngRounds = CLng(Trim$(Left$(cboRounds.Text, 2)))
    End If
    
End Sub

Private Sub cboRandom_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    txtRandom.Text = vbNullString
    RndData_Processing cboRandom.ListIndex

End Sub

Private Sub chkExtraInfo_Click()
    
    ' Designates if input file is to be overwritten after
    ' performing encryption/decryption.
    '
    ' Cipher processing
    ' Checked   - Create new file to hold encrypted/decrypted data
    ' Unchecked - Overwrite input file after encryption/decryption
    '
    ' Hash processing
    ' Checked   - Return hashed data in lowercase format
    ' Unchecked - Return hashed data in uppercase format
    
    If mblnCipher Then
        mblnCreateNewFile = CBool(chkExtraInfo.Value)
        
    ElseIf mblnHash Then
        mblnHashLowercase = CBool(chkExtraInfo.Value)
    End If
    
End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear                   ' clear the clipboard
    Clipboard.SetText txtOutput.Text  ' load clipboard with textbox data

End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub cmdBrowse_Click()
    
    Dim strFilters As String

    mstrFolder = vbNullString
    mstrFilename = vbNullString
    txtInputFile.Text = vbNullString
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    
    strFilters = "All Files (*.*)|*.*"

    On Error GoTo Cancel_Selected

    ' Get the file location. Display the File Open dialog box
    With cdFileOpen
         .CancelError = True    ' Set CancelError is True
         .DialogTitle = "Select file to encrypt"
         .DefaultExt = "*.*"
         .Filter = strFilters
         .Flags = cdlOFNLongNames Or cdlOFNExplorer
         .FilterIndex = 1       ' Specify default filter
         .FileName = vbNullString
         .ShowOpen              ' Display the Open dialog box
    End With

    ' Save the name of the item selected
    mstrFilename = TrimStr(cdFileOpen.FileName)

    ' separate path from filename
    If Len(mstrFilename) > 0 Then
        txtInputFile.Text = vbNullString
        txtInputFile.Text = ShrinkToFit(mstrFilename, 70) ' Original file name
    End If

CleanUp:
    Exit Sub

Cancel_Selected:
    mstrFilename = vbNullString
    mstrFolder = vbNullString
    txtInputFile.Text = vbNullString
    txtOutput.Text = vbNullString
    On Error GoTo 0
    GoTo CleanUp

End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim lngIdx As Long

    Select Case Index
    
           Case 0  ' OK button
                Screen.MousePointer = vbHourglass
                gblnStopProcessing = False
                SetDllProcessingFlag
                ResetProgressBar
                cmdChoice_GotFocus 1
                
                ' Lock controls
                With frmMain
                    .cmdChoice(2).Enabled = False
                    .cmdChoice(3).Enabled = False
                    .fraChoice.Enabled = False
                    .fraEncrypt(1).Enabled = False
                    .picDataType.Enabled = False
                    .txtPwd.Enabled = False
                End With
                
                ' Temporarily lock txtOutput textbox while processing.
                ' This will speed things up and reduce flicker.
                LockWindowUpdate frmMain.txtOutput.hwnd
    
                If mblnStringData Then
                    LockWindowUpdate frmMain.txtInputString.hwnd
                End If
                
                If mblnCipher Then
                    
                    If Len(Trim$(mstrPassword)) = 0 Then
                        InfoMsg "Password is missing."
                    Else
                        Cipher_Processing
                    End If
                End If
                
                If mblnCRC32 Then
                    CRC32_Processing
                End If
                
                If mblnHash Then
                    Hash_Processing
                End If
                
                If mblnPrng Then
                    lngIdx = cboRandom.ListIndex
                    RndData_Processing lngIdx
                End If

                DoEvents
                UpdateRegistry
                
                ' Unlock controls
                With frmMain
                    .cmdChoice(2).Enabled = True
                    .cmdChoice(3).Enabled = True
                    .fraChoice.Enabled = True
                    .fraEncrypt(1).Enabled = True
                    .picDataType.Enabled = True
                    .txtPwd.Enabled = True
                End With
                
                ResetProgressBar
                cmdChoice_GotFocus 0
                
           Case 1  ' Stop button
                DoEvents
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                SetDllProcessingFlag
                DoEvents
                
                UpdateRegistry
                ResetProgressBar
                
                ' Unlock controls
                With frmMain
                    .cmdChoice(2).Enabled = True
                    .cmdChoice(3).Enabled = True
                    .fraChoice.Enabled = True
                    .fraEncrypt(1).Enabled = True
                    .picDataType.Enabled = True
                    .txtPwd.Enabled = True
                End With
                
                cmdChoice_GotFocus 0
                DoEvents

           Case 2  ' Show About form
                frmMain.Hide
                frmAbout.DisplayAbout

           Case Else  ' EXIT button
                DoEvents
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                SetDllProcessingFlag
                DoEvents
                
                UpdateRegistry
                ResetProgressBar
                TerminateProgram
    End Select

CleanUp:
    DoEvents
    Screen.MousePointer = vbDefault   ' Reset mouse pointer to normal
    LockWindowUpdate 0&               ' unlock txtOutput textbox after processing
    DoEvents
                

End Sub

Private Sub cmdChoice_GotFocus(Index As Integer)

    Select Case Index
           Case 0
                cmdChoice(0).Enabled = True
                cmdChoice(0).Visible = True
                cmdChoice(1).Enabled = False
                cmdChoice(1).Visible = False
           Case 1
                cmdChoice(0).Enabled = False
                cmdChoice(0).Visible = False
                cmdChoice(1).Enabled = True
                cmdChoice(1).Visible = True
    End Select

    Refresh

End Sub

Private Sub SetDllProcessingFlag()

    ' Called by cmdChoice_Click()
    
    ' If the particular object is active then
    ' set the property value
    If Not mobjCipher Is Nothing Then
        mobjCipher.StopProcessing = gblnStopProcessing
    End If
    
    If Not mobjCRC32 Is Nothing Then
        mobjCRC32.StopProcessing = gblnStopProcessing
    End If
                    
    If Not mobjHash Is Nothing Then
        mobjHash.StopProcessing = gblnStopProcessing
    End If
    
    If Not mobjRandom Is Nothing Then
        mobjRandom.StopProcessing = gblnStopProcessing
    End If
                
    DoEvents
    
End Sub

Private Sub Form_Load()

    gblnStopProcessing = False
    
    ' Instantiate class and DLL objects
    Set mobjKeyEdit = New cKeyEdit
    Set mobjCipher = New kiCrypt.cCipher
    Set mobjCRC32 = New kiCrypt.cCRC32
    Set mobjHash = New kiCrypt.cHash
    Set mobjRandom = New kiCrypt.cPrng
    
    GetRegistryData
    DisableX frmMain   ' Disable "X" in upper right corner of form
    LoadComboBox
    
    ' Passwords or phrases are case sensitive and
    ' have a potential length of fifty characters
    '
    '               123456789+123456789+123456789+123456789+123456789+
    mstrPassword = "This is My Password"
    mlngPwdLength_Min = mobjCipher.PasswordLength_Min
    mlngPwdLength_Max = mobjCipher.PasswordLength_Max
    mblnCipher = True
    mblnCRC32 = False
    mblnHash = False
    mblnPrng = False
    mblnStringData = True
    
    With frmMain
        .Caption = gstrVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
        .lblEncryptMsg.Visible = False
        .txtRandom.BackColor = &HE0E0E0  ' Light gray
        .txtOutput.BackColor = &HE0E0E0  ' Light gray
        
        ' set the command buttons
        .cmdChoice(0).Enabled = True
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Enabled = False
        .cmdChoice(1).Visible = False
        
        .cmdCopy.Enabled = False
        .cmdCopy.Visible = False
        
        optChoice_Click 0
        cboEncrypt_Click
        ResetProgressBar
        
        ' Center the form on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless   ' reduce flicker
        .Refresh
    End With

    mobjKeyEdit.CenterCaption frmMain
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    gblnStopProcessing = True

    ' If the object is still active then
    ' send a command to stop it.
    If Not mobjCipher Is Nothing Then
        mobjCipher.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjCRC32 Is Nothing Then
        mobjCRC32.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjHash Is Nothing Then
        mobjHash.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjRandom Is Nothing Then
        mobjRandom.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    Set mobjKeyEdit = Nothing
    Set mobjCipher = Nothing
    Set mobjCRC32 = Nothing
    Set mobjHash = Nothing
    Set mobjRandom = Nothing
    Screen.MousePointer = vbDefault

    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

' 29-Jan-2010 Add events to track cipher progress
Private Sub mobjCipher_CipherProgress(ByVal lngProgress As Long)
    
    If mblnStringData Then
        Exit Sub
    End If
    
    ProgressBar picProgressBar, lngProgress, vbBlue
    DoEvents
    
End Sub

' 29-Jan-2010 Add events to track CRC32 progress
Private Sub mobjCRC32_CRCProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress, vbRed
    DoEvents
    
End Sub

' 29-Jan-2010 Add events to track hash progress
Private Sub mobjHash_HashProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress, vbBlack
    DoEvents
    
End Sub

Private Sub optChoice_Click(Index As Integer)

    Dim intIndex As Integer
    
    mblnCipher = False
    mblnCRC32 = False
    mblnHash = False
    mblnPrng = False
    
    Select Case Index
           Case 0  ' encryption
                mblnCipher = True

                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = True
                    .fraEncrypt(1).Visible = True
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = False
                    
                    ' Prepare combo boxes
                    .cboEncrypt.Enabled = True
                    .cboEncrypt.ForeColor = vbBlack
                    
                    Select Case mlngCipherAlgo
                           Case eCIPHER_BASE64
                                .cboBlockLength.Enabled = False
                                .cboBlockLength.ForeColor = vbGrayText
                                .cboKeyLength.Enabled = False
                                .cboKeyLength.ForeColor = vbGrayText
                                .cboRounds.Enabled = False
                                .cboRounds.ForeColor = vbGrayText
                                .lblKeyMix.Enabled = False
                                .lblKeyMix.Visible = False
                                .txtKeyMix.Enabled = False
                                .txtKeyMix.Visible = False
                           
                           Case eCIPHER_RIJNDAEL
                                .cboBlockLength.Enabled = True
                                .cboBlockLength.ForeColor = vbBlack
                                .cboKeyLength.Enabled = True
                                .cboKeyLength.ForeColor = vbBlack
                                .cboRounds.Enabled = True
                                .cboRounds.ForeColor = vbBlack
                                .lblKeyMix.Enabled = False
                                .lblKeyMix.Visible = False
                                .txtKeyMix.Enabled = False
                                .txtKeyMix.Visible = False
                           
                           Case eCIPHER_BLOWFISH, eCIPHER_GOST
                                .cboBlockLength.Enabled = False
                                .cboBlockLength.ForeColor = vbGrayText
                                .cboKeyLength.Enabled = True
                                .cboKeyLength.ForeColor = vbBlack
                                .cboRounds.Enabled = True
                                .cboRounds.ForeColor = vbBlack
                                .lblKeyMix.Enabled = True
                                .lblKeyMix.Visible = True
                                .txtKeyMix.Enabled = True
                                .txtKeyMix.Visible = True
                           
                           Case Else
                                .cboBlockLength.Enabled = False
                                .cboBlockLength.ForeColor = vbGrayText
                                .cboKeyLength.Enabled = True
                                .cboKeyLength.ForeColor = vbBlack
                                .cboRounds.Enabled = True
                                .cboRounds.ForeColor = vbBlack
                                .lblKeyMix.Enabled = False
                                .lblKeyMix.Visible = False
                                .txtKeyMix.Enabled = False
                                .txtKeyMix.Visible = False
                    End Select
                    
                    .lblAlgo(2).Caption = "Password Hash Algorithm"
                    With .lblPwd
                        .Visible = True
                        .Caption = vbNullString
                        .Caption = "Enter " & CStr(mlngPwdLength_Min) & "-" & _
                                   CStr(mlngPwdLength_Max) & _
                                   " character password or phrase  [ Case sensitive ]"
                    End With
                    .lblEncrypt(0).Caption = "Data to be encrypted or decrypted"
                    .lblEncrypt(1).Caption = "Output file name and location"
                    .lblEncrypt(1).Visible = True
                    .lblEncryptMsg.Caption = "After encryption, data sizes will not match original sizes.  " & _
                                             "This is due to internal padding and information " & _
                                             "required for later decryption."
                    .lblEncryptMsg.Visible = True
                    .txtOutput.Text = vbNullString
                    .txtPwd.Visible = True
                    .txtPwd.Locked = False
                    .txtPwd.Text = mstrPassword
                    .chkExtraInfo.Caption = vbNullString
                    .chkExtraInfo.Visible = True
                    .chkExtraInfo.Caption = "Create new target file"
                End With
                 
                intIndex = IIf(mblnStringData, 0, 1)
                optDataType_Click intIndex
                cboRounds_Click
                
           Case 1 ' CRC-32
                mblnCRC32 = True

                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = False
                    .fraEncrypt(1).Visible = False
                    .lblEncrypt(0).Caption = "Data to be calculated for CRC"
                    .lblEncrypt(1).Caption = "Calculated CRC-32 value in hex"
                    .lblEncrypt(1).Visible = True
                    .lblPwd.Visible = False
                    .lblPwd.Caption = vbNullString
                    .lblEncryptMsg.Visible = False
                    .lblKeyMix.Enabled = False
                    .lblKeyMix.Visible = False
                    .txtKeyMix.Enabled = False
                    .txtKeyMix.Visible = False
                    .txtPwd.Visible = False
                    .txtOutput.Text = vbNullString
                    .txtPwd.Text = vbNullString
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = True
                    .chkExtraInfo.Caption = ""
                    .chkExtraInfo.Visible = False
                End With

           Case 2  ' Hash
                mblnHash = True
                
                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = True
                    .fraEncrypt(1).Visible = True
                    
                    ' Prepare combo boxes
                    .cboEncrypt.Enabled = False
                    .cboEncrypt.ForeColor = vbGrayText
                    .cboKeyLength.Enabled = False
                    .cboKeyLength.ForeColor = vbGrayText
                    .cboBlockLength.Enabled = False
                    .cboBlockLength.ForeColor = vbGrayText
                    .cboRounds.Enabled = True
                    .cboRounds.ForeColor = vbBlack
                    
                    .lblEncrypt(0).Caption = "Data to be hashed"
                    .lblEncrypt(1).Caption = "Hashed results"
                    .lblEncrypt(1).Visible = True
                    .lblAlgo(2).Caption = vbNewLine & "Hash Algorithm"
                    .lblPwd.Caption = vbNullString
                    .lblPwd.Visible = False
                    .lblEncryptMsg.Caption = "Be patient.  The more data to process the longer it will take."
                    .lblEncryptMsg.Visible = True
                    .lblKeyMix.Enabled = False
                    .lblKeyMix.Visible = False
                    .txtKeyMix.Enabled = False
                    .txtKeyMix.Visible = False
                    .txtOutput.Text = vbNullString
                    .txtPwd.Visible = False
                    .txtPwd.Text = vbNullString
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = True
                    .chkExtraInfo.Caption = ""
                    .chkExtraInfo.Visible = True
                    .chkExtraInfo.Caption = "Return as lowercase"
                End With
                
                intIndex = IIf(mblnStringData, 0, 1)
                optDataType_Click intIndex
                cboHash_Click
                
           Case 3  ' Random data
                mblnPrng = True
                With frmMain
                    .fraMain(0).Visible = True
                    .fraMain(0).Enabled = True
                    .fraMain(1).Visible = False
                    .fraMain(1).Enabled = False
                    .lblEncryptMsg.Visible = False
                    .lblKeyMix.Enabled = False
                    .lblKeyMix.Visible = False
                    .txtKeyMix.Enabled = False
                    .txtKeyMix.Visible = False
                    .txtRandom.Text = vbNullString
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = False
                    .cboRandom.ListIndex = 0
                    .chkExtraInfo.Visible = False
                End With
                
                cboRandom_Click
    End Select

End Sub

Private Sub optDataType_Click(Index As Integer)

    With frmMain
        Select Case Index
    
               Case 0   ' string data
                    mblnStringData = True
                    .lblEncrypt(0).Caption = "Data to be encrypted or decrypted"
                    .optDataType(0).Value = True
                    .optDataType(1).Value = False
                    .cmdBrowse.Enabled = False
                    .cmdBrowse.Visible = False
                    .txtInputFile.Text = vbNullString
                    .txtInputFile.Visible = False
                    .txtOutput.Text = vbNullString
                    
                    If mblnCRC32 Or mblnHash Then
                        With .txtInputString
                            .Height = 3255
                            .Text = vbNullString
                            .Visible = True
                        End With
                        .txtOutput.Visible = True
                        
                    ElseIf mblnCipher Then
                        With .txtInputString
                            .Height = 4400
                            .Text = vbNullString
                            .Visible = True
                        End With
                        .txtOutput.Visible = False
                    End If
                                    
               Case 1   ' file data
                    mblnStringData = False
                    .lblEncrypt(0).Caption = "Data to be encrypted or decrypted"
                    .optDataType(0).Value = False
                    .optDataType(1).Value = True
                    .cmdBrowse.Enabled = True
                    .cmdBrowse.Visible = True
                    .lblEncrypt(1).Visible = True
                    .txtInputString.Text = vbNullString
                    .txtInputString.Visible = False
                    .txtInputFile.Text = vbNullString
                    .txtInputFile.Visible = True
                    .txtOutput.Text = vbNullString
                    .txtOutput.Visible = True
        End Select
        
        If mblnStringData Then
            .picProgressBar.Visible = False
        Else
            .picProgressBar.Visible = True
        End If

        If mblnHash Then
            If mblnHashLowercase Then
                .chkExtraInfo.Value = vbChecked
            Else
                .chkExtraInfo.Value = vbUnchecked
            End If
                
            .chkExtraInfo.Visible = True   ' Show checkbox
            
        ElseIf mblnCipher Then
            
            If mblnStringData Then
                .chkExtraInfo.Value = vbUnchecked  ' Set to not checked
                .chkExtraInfo.Visible = False      ' Hide checkbox
            Else
                If mblnCreateNewFile Then
                    .chkExtraInfo.Value = vbChecked
                Else
                    .chkExtraInfo.Value = vbUnchecked
                End If
                
                .chkExtraInfo.Visible = True   ' Show checkbox
            End If
            
        Else
            .chkExtraInfo.Value = vbUnchecked  ' Set to not checked
            .chkExtraInfo.Visible = False      ' Hide checkbox
        End If
    End With
    
    If mblnCRC32 Or mblnHash Then
        cmdCopy.Enabled = False
    End If
    
End Sub

' ***************************************************************************
' Data functions
' ***************************************************************************
Private Sub LoadComboBox()

    Dim lngIdx As Long
    
    mblnLoading = True

    ' Encryption algorithms
    With frmMain
        With .cboEncrypt
            .Clear
            .AddItem "ArcFour"
            .AddItem "Base64"
            .AddItem "Blowfish"
            .AddItem "Gost"
            .AddItem "Rijndael (AES)"
            .AddItem "Serpent"
            .AddItem "Skipjack"
            .AddItem "Twofish"
            .ListIndex = 0
        End With
    
        ' Hash algorithms
        With .cboHash
            .Clear
            .AddItem "MD2"               ' 0
            .AddItem "MD4"               ' 1
            .AddItem "MD5"               ' 2
            .AddItem "SHA-1"             ' 3
            .AddItem "SHA-256"           ' 4
            .AddItem "SHA-384"           ' 5
            .AddItem "SHA-512"           ' 6
            .AddItem "Tiger-128"         ' 7
            .AddItem "Tiger-160"         ' 8
            .AddItem "Tiger-192"         ' 9
            .AddItem "Tiger-224"         ' 10
            .AddItem "Tiger-256"         ' 11
            .AddItem "Tiger-384"         ' 12
            .AddItem "Tiger-512"         ' 13
            .AddItem "Whirlpool-224"     ' 14
            .AddItem "Whirlpool-256"     ' 15
            .AddItem "Whirlpool-384"     ' 16
            .AddItem "Whirlpool-512"     ' 17
            .ListIndex = 4               ' Default = SHA-256
        End With
    
        ' Encryption algorithms
        With .cboRandom
            .Clear
            .AddItem "Keyboard Chars (1400 bytes)"
            .AddItem "Hex String (800 bytes)"
            .AddItem "Hex Array (400 elements)"
            .AddItem "Byte Array (320 elements)"
            .AddItem "Long Array (120 digits)"
            .AddItem "Double Array (80 digits)"
            .ListIndex = 0
        End With
    
        With .cboKeyLength
            .Clear
            For lngIdx = 128 To 256 Step 32
                .AddItem CStr(lngIdx) & " bits"
            Next lngIdx
            .ListIndex = 0
            .Enabled = True
        End With
            
        With .cboBlockLength
            .Clear
            For lngIdx = 128 To 256 Step 32
                .AddItem CStr(lngIdx) & " bits"
            Next lngIdx
            .ListIndex = 0
            .Enabled = False
        End With
            
        With .cboRounds
            .Clear
            For lngIdx = 1 To 10
                .AddItem CStr(lngIdx)
            Next lngIdx
            .ListIndex = 0
        End With
    End With
    
    mblnLoading = False
    
End Sub

' ***************************************************************************
' Routine:       Cipher_Processing
'
' Description:   Encryption demonstration
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated routine to match new screen
' 01-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected error handling and flow of data
' 03-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified structure of code for easier reading
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added reference to new property CreateNewFile().  Change input
'              to a variable if you want to overwrite file to be encrypted or
'              decrypted
' ***************************************************************************
Private Sub Cipher_Processing()

    Dim hFile         As Long
    Dim strMsg        As String
    Dim strData       As String
    Dim strOutputFile As String
    Dim abytData()    As Byte
    Dim astrMsgBox()  As String
    Dim lngEncrypt    As enumCIPHER_ACTION
    
    Erase abytData()    ' Always start with empty arrays
    Erase astrMsgBox()
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(Trim$(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process"
            txtInputString.SetFocus
            GoTo Cipher_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(Trim$(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing"
            GoTo Cipher_Processing_CleanUp
        End If
    End If
    
    ' Evaluate password
    Select Case Len(mstrPassword)
           
           Case 0
                InfoMsg "Password is missing." & vbNewLine & vbNewLine & _
                        "Minimum length:   " & CStr(mlngPwdLength_Min) & " characters" & vbNewLine & _
                        "Maximum length:   " & CStr(mlngPwdLength_Max) & " characters"
                GoTo Cipher_Processing_CleanUp
    
           Case Is < mlngPwdLength_Min
                InfoMsg "Password is too short." & vbNewLine & _
                        "Minimum length:   " & CStr(mlngPwdLength_Min) & " characters"
                GoTo Cipher_Processing_CleanUp
    
           Case Is > mlngPwdLength_Max
                InfoMsg "Password is too long." & vbNewLine & _
                        "Maximum length:   " & CStr(mlngPwdLength_Max) & " characters"
                GoTo Cipher_Processing_CleanUp
    End Select
    
    ' Disable form STOP button until a choice is made
    cmdChoice(1).Enabled = False
    
    '----------------------------------------------------------
    ' Prepare message box display.
    '
    ' These are the button captions,
    ' in order, from left to right.
    ReDim astrMsgBox(3)
    astrMsgBox(0) = "Encrypt"
    astrMsgBox(1) = "Decrypt"
    astrMsgBox(2) = "Cancel"
    
    ' Prompt user with message box
    Select Case MessageBoxH(Me.hwnd, GetDesktopWindow(), _
                            "What do you want to do?  ", _
                            PGM_NAME, astrMsgBox(), eMSG_ICONQUESTION)
           
           ' These are valid responses
           Case IDYES:    lngEncrypt = eCA_ENCRYPT
           Case IDNO:     lngEncrypt = eCA_DECRYPT
           Case IDCANCEL: GoTo Cipher_Processing_CleanUp
    End Select
    '----------------------------------------------------------
    
    ' Enable form STOP button after a choice is made
    cmdChoice(1).Enabled = True
    
    ' *********************************************************
    ' Encrypt/Decrpyt - String
    ' *********************************************************
    If mblnStringData Then
                    
        Screen.MousePointer = vbHourglass
        
        Select Case lngEncrypt
               Case eCA_ENCRYPT    ' Encrypt string
                    With mobjCipher
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mstrPassword
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            txtPwd.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        .HashMethod = mlngHashAlgo        ' Type of hash algorithm selected
                        .CipherMethod = mlngCipherAlgo    ' Type of cipher algorithm selected
                        .KeyLength = mlngKeyLength        ' Ignored by Base64
                        .PrimaryKeyRounds = mlngKeyMix    ' Only by Blowfish, GOST
                        .BlockSize = mlngBlockLength      ' Used by Rijndael only
                        .CipherRounds = mlngRounds        ' Ignored by Base64, Rijndael
                    
                        strData = Trim$(txtInputString.Text)      ' Remove any hidden characters
                        abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                        
                        ' See if using Base64 cipher
                        If mlngCipherAlgo = eCIPHER_BASE64 Then
                            ' Encrypt data string
                            If .EncryptString(abytData()) Then
                                strData = ByteArrayToString(abytData())   ' Convert byte array to string data
                            Else
                                gblnStopProcessing = .StopProcessing      ' See if processing aborted
                            End If
                        Else
                            ' Encrypt data string
                            If .EncryptString(abytData()) Then
                            
                                strData = ByteArrayToHex(abytData())  ' convert single charaters to hex
                            
                                ' Verify that this is hex data
                                If Not IsHexData(strData) Then
                                    InfoMsg "Failed to convert encrypted data to hex."
                                    gblnStopProcessing = True
                                End If
                            Else
                                gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            End If
                        End If
                    End With
            
                    txtInputString.Text = vbNullString                            ' Empty text box
                    txtInputString.Text = TrimStr(strData)  ' Store string data in text box
            
               Case eCA_DECRYPT    ' Decrypt string
                    With mobjCipher
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mstrPassword
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            txtPwd.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        .HashMethod = mlngHashAlgo        ' Type of hash algorithm selected
                        .CipherMethod = mlngCipherAlgo    ' Type of cipher algorithm selected
                        .KeyLength = mlngKeyLength        ' Ignored by Base64
                        .PrimaryKeyRounds = mlngKeyMix    ' Only by Blowfish, GOST
                        .BlockSize = mlngBlockLength      ' Used by Rijndael only
                        .CipherRounds = mlngRounds        ' Ignored by Base64, Rijndael
                    
                        strData = Trim$(txtInputString.Text)  ' Remove leading and trailing blanks
                        
                        ' See if using Base64 cipher
                        If mlngCipherAlgo = eCIPHER_BASE64 Then
                            abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                            .DecryptString abytData()                 ' Decrypt data string
                            gblnStopProcessing = .StopProcessing      ' See if processing aborted
                        Else
                            ' Verify that this is hex data
                            If IsHexData(strData) Then
                                abytData() = HexToByteArray(strData)  ' Convert hex to single char
                                .DecryptString abytData()             ' Decrypt data string
                                gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            Else
                                strMsg = "This text is not hex data or length is not divisible by two."
                                strMsg = strMsg & vbNewLine & vbNewLine & "Cannot decrypt."
                                InfoMsg strMsg
                                GoTo Cipher_Processing_CleanUp
                            End If
                        
                        End If
                    End With
            
                    strData = ByteArrayToString(abytData())             ' Convert byte array to string data
                    txtInputString.Text = vbNullString                            ' Empty text box
                    txtInputString.Text = TrimStr(strData)  ' Store string data in text box
        End Select
            
        DoEvents
        If gblnStopProcessing Then
            GoTo Cipher_Processing_CleanUp
        End If
    
    Else
        ' *********************************************************
        ' Encrypt/Decrpyt - File
        ' *********************************************************
        If IsPathValid(mstrFilename) Then
        
            If mblnCreateNewFile Then
                Select Case lngEncrypt
                       Case eCA_ENCRYPT: strOutputFile = mstrFilename & ENCRYPT_EXT
                       Case eCA_DECRYPT: strOutputFile = mstrFilename & DECRYPT_EXT
                End Select
            End If
            
        Else
            InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename
            txtInputFile.SetFocus
            GoTo Cipher_Processing_CleanUp
        End If
    
        ' If output file exist
        ' then verify it is empty
        DoEvents
        If IsPathValid(strOutputFile) Then
            
            SetFileAttributes strOutputFile, FILE_ATTRIBUTE_NORMAL
            hFile = FreeFile
            Open strOutputFile For Output As #hFile
            Close #hFile
            DoEvents
            
        End If
        
        Select Case lngEncrypt
               Case eCA_ENCRYPT    ' Encrypt file
                    ' Verify overwrite message
                    If Not mblnCreateNewFile Then
                        
                        If ResponseMsg("Are you sure you want to overwrite input file?", _
                                       vbYesNo, "Verify output target") = vbNo Then
                                       
                            txtInputFile.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
                        
                        strOutputFile = mstrFilename
                    End If
                    
                    With mobjCipher
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mstrPassword
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            txtPwd.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        .HashMethod = mlngHashAlgo          ' Type of hash algorithm selected
                        .CipherMethod = mlngCipherAlgo      ' Type of cipher algorithm selected
                        .KeyLength = mlngKeyLength          ' Ignored by Base64
                        .PrimaryKeyRounds = mlngKeyMix      ' Only by Blowfish, GOST
                        .BlockSize = mlngBlockLength        ' Used by Rijndael only
                        .CipherRounds = mlngRounds          ' Ignored by Base64, Rijndael
                        .CreateNewFile = mblnCreateNewFile  ' True - Create new output file
                                                            ' False - Overwrite input file
                            
                        If .EncryptFile(mstrFilename) Then
                        
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            txtOutput.Text = vbNullString
                            Screen.MousePointer = vbDefault
                                                
                            If Not gblnStopProcessing Then
                                txtOutput.Text = strOutputFile
                                strMsg = "Finished encrypting file." & vbNewLine & vbNewLine
                                strMsg = strMsg & strOutputFile & vbNewLine & vbNewLine
                                InfoMsg strMsg
                            End If
                        Else
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    End With
            
               Case eCA_DECRYPT    ' Decrypt file
                    ' Verify overwrite message
                    If Not mblnCreateNewFile Then
                        
                        If ResponseMsg("Are you sure you want to overwrite input file?", _
                                       vbYesNo, "Verify output target") = vbNo Then
                                       
                            txtInputFile.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
                        
                        strOutputFile = mstrFilename
                    End If
                    
                    With mobjCipher
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mstrPassword
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        .HashMethod = mlngHashAlgo          ' Type of hash algorithm selected
                        .CipherMethod = mlngCipherAlgo      ' Type of cipher algorithm selected
                        .KeyLength = mlngKeyLength          ' Ignored by Base64
                        .PrimaryKeyRounds = mlngKeyMix      ' Only by Blowfish, GOST
                        .BlockSize = mlngBlockLength        ' Used by Rijndael only
                        .CipherRounds = mlngRounds          ' Ignored by Base64, Rijndael
                        .CreateNewFile = mblnCreateNewFile  ' True - Create new output file
                                                            ' False - Overwrite input file
                
                        If .DecryptFile(mstrFilename) Then
                            
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            txtOutput.Text = vbNullString
                            Screen.MousePointer = vbDefault
                                                
                            If Not gblnStopProcessing Then
                                txtOutput.Text = strOutputFile
                                strMsg = "Finished decrypting file." & vbNewLine & vbNewLine
                                strMsg = strMsg & strOutputFile & vbNewLine & vbNewLine
                                InfoMsg strMsg
                            End If
                            
                        Else
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                    End With
        End Select
            
    End If
    
Cipher_Processing_CleanUp:
    Erase abytData()     ' Always empty arrays when not needed
    Erase astrMsgBox()
    Screen.MousePointer = vbDefault  ' Set mouse pointer back to normal
    
End Sub

Private Sub CRC32_Processing()
                    
    Dim strHex     As String
    Dim abytData() As Byte
    
    Screen.MousePointer = vbHourglass
    
    Erase abytData()   ' Always start eith an empty array
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(Trim$(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process"
            GoTo CRC32_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(Trim$(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing"
            GoTo CRC32_Processing_CleanUp
        End If
    End If
    
    With mobjCRC32
    
        ' CRC32 string data
        If mblnStringData Then
        
            abytData() = StringToByteArray(txtInputString.Text)   ' Convert string data to byte array
            strHex = .CRC32_String(abytData(), True)              ' Calculate CRC
            gblnStopProcessing = .StopProcessing                  ' See if processing aborted
            
        Else
            ' CRC32 a file
            If IsPathValid(mstrFilename) Then
                abytData() = StringToByteArray(mstrFilename)   ' Convert string data to byte array
                strHex = .CRC32_File(abytData(), True)         ' Calculate CRC
                gblnStopProcessing = .StopProcessing           ' See if processing aborted
            Else
                InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename
                txtInputFile.SetFocus
                GoTo CRC32_Processing_CleanUp
            End If
        End If
    
    End With
    
    DoEvents
    If gblnStopProcessing Then
        GoTo CRC32_Processing_CleanUp
    End If
                
    txtOutput.Text = TrimStr(strHex)
    cmdCopy.Enabled = True
    
CRC32_Processing_CleanUp:
    Erase abytData()   ' Always empty arrays when not needed
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Hash_Processing()
                    
    Dim strOutput  As String
    Dim abytData() As Byte
    Dim abytHash() As Byte
    
    Screen.MousePointer = vbHourglass
    
    Erase abytData()    ' Always start with empty arrays
    Erase abytHash()
    
    strOutput = vbNullString
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(Trim$(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process"
            GoTo Hash_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(Trim$(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing"
            GoTo Hash_Processing_CleanUp
        End If
    End If
    
    With mobjHash
        .StopProcessing = False                ' Reset stop flag
        .HashMethod = mlngHashAlgo             ' Hash algorithm selected
        .HashRounds = mlngRounds               ' Number of passes
        .ReturnLowercase = mblnHashLowercase   ' TRUE = Return as lowercase
                                               ' FALSE = Return as uppercase
        ' Hash string data
        If optDataType(0).Value Then
            
            abytData() = StringToByteArray(txtInputString.Text)    ' Convert to byte array
            abytHash() = .HashString(abytData())                   ' Hash string data
            gblnStopProcessing = .StopProcessing                   ' See if processing aborted
            strOutput = ByteArrayToString(abytHash())              ' Convert byte array to string
            
        Else
            ' Hash a file
            If IsPathValid(mstrFilename) Then
                abytData() = StringToByteArray(mstrFilename)   ' Convert to byte array
                abytHash() = .HashFile(abytData())             ' Hash file
                gblnStopProcessing = .StopProcessing           ' See if processing aborted
                strOutput = ByteArrayToString(abytHash())      ' Convert byte array to string
            Else
                InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename
                GoTo Hash_Processing_CleanUp
            End If
        End If
    
    End With
    
    DoEvents
    If gblnStopProcessing Then
        GoTo Hash_Processing_CleanUp
    End If
    
    txtOutput.Text = TrimStr(strOutput)
    cmdCopy.Enabled = True
    
Hash_Processing_CleanUp:
    Erase abytData()    ' Always empty arrays when not needed
    Erase abytHash()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub RndData_Processing(ByVal lngChoice As Long)

    Dim strOutput As String
    Dim lngIndex  As Long
    Dim lngCount  As Long
    Dim avntData  As Variant

    Screen.MousePointer = vbHourglass
    avntData = Empty
    txtRandom.Text = vbNullString
    strOutput = vbNullString
    lngCount = 0

    Select Case lngChoice

           Case 0   ' ASCII String
                txtRandom.Text = mobjRandom.BuildWithinRange(1400, 32, 126, ePRNG_ASCII)

           Case 1   ' Hex String
                txtRandom.Text = mobjRandom.BuildRndData(800, ePRNG_HEX)

           Case 2   ' Hex Array
                avntData = mobjRandom.BuildRndData(400, ePRNG_HEX_ARRAY)

                DoEvents
                If gblnStopProcessing Then
                    GoTo RndData_Processing_CleanUp
                End If

                For lngIndex = 0 To UBound(avntData) - 1
                    
                    If lngCount = 0 Then
                        strOutput = strOutput & Space$(1) & avntData(lngIndex)
                    Else
                        strOutput = strOutput & Space$(2) & avntData(lngIndex)
                    End If
                    
                    lngCount = lngCount + 1

                    If lngCount = 20 Then
                        lngCount = 0
                        strOutput = strOutput & vbNewLine
                    End If

                    DoEvents
                    If gblnStopProcessing Then
                        strOutput = vbNullString
                        Exit For    ' exit For..Next loop
                    End If

                Next lngIndex

                txtRandom.Text = strOutput

           Case 3   ' Byte Array
                avntData = mobjRandom.BuildRndData(320, ePRNG_BYTE_ARRAY)

                DoEvents
                If gblnStopProcessing Then
                    GoTo RndData_Processing_CleanUp
                End If

                For lngIndex = 0 To UBound(avntData) - 1
                    
                    If lngCount = 0 Then
                        strOutput = strOutput & Space$(1) & Format$(avntData(lngIndex), "@@@")
                    Else
                        strOutput = strOutput & Space$(2) & Format$(avntData(lngIndex), "@@@")
                    End If
                    
                    lngCount = lngCount + 1

                    If lngCount = 16 Then
                        lngCount = 0
                        strOutput = strOutput & vbNewLine
                    End If

                    DoEvents
                    If gblnStopProcessing Then
                        strOutput = vbNullString
                        Exit For    ' exit For..Next loop
                    End If

                Next lngIndex

                txtRandom.Text = strOutput

           Case 4   ' Long Array
                avntData = mobjRandom.BuildRndData(120, ePRNG_LONG_ARRAY)

                DoEvents
                If gblnStopProcessing Then
                    GoTo RndData_Processing_CleanUp
                End If

                For lngIndex = 0 To UBound(avntData) - 1
                    
                    If lngCount = 0 Then
                        strOutput = strOutput & Space$(1) & Format$(avntData(lngIndex), String$(11, "@"))
                    Else
                        strOutput = strOutput & Space$(2) & Format$(avntData(lngIndex), String$(11, "@"))
                    End If
                    
                    lngCount = lngCount + 1

                    If lngCount = 6 Then
                        lngCount = 0
                        strOutput = strOutput & vbNewLine
                    End If

                    DoEvents
                    If gblnStopProcessing Then
                        strOutput = vbNullString
                        Exit For    ' exit For..Next loop
                    End If

                Next lngIndex

                txtRandom.Text = strOutput

           Case 5   ' Double Array (0 to 1)
                avntData = mobjRandom.BuildRndData(80, ePRNG_DBL_ARRAY)

                DoEvents
                If gblnStopProcessing Then
                    GoTo RndData_Processing_CleanUp
                End If

                For lngIndex = 0 To UBound(avntData) - 1
                    
                    If lngCount = 0 Then
                        strOutput = strOutput & Space$(1) & Format$(avntData(lngIndex), String$(17, "@"))
                    Else
                        strOutput = strOutput & Space$(3) & Format$(avntData(lngIndex), String$(17, "@"))
                    End If
                    lngCount = lngCount + 1

                    If lngCount = 4 Then
                        strOutput = strOutput & vbNewLine
                        lngCount = 0
                    End If

                    DoEvents
                    If gblnStopProcessing Then
                        strOutput = vbNullString
                        Exit For    ' exit For..Next loop
                    End If

                Next lngIndex

                txtRandom.Text = strOutput

    End Select

RndData_Processing_CleanUp:
    DoEvents
    avntData = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar, 0, vbWhite
    
End Sub

' ***************************************************************************
' Routine:       ProgessBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If forecolor is white then progressbar
    ' is being reset to a starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Arial"     ' Name of font
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent <= MAX_PERCENT Then

        With objProgBar
        
            ' Error trap in case code attempts to set
            ' scalewidth greater than the max allowable
            If lngPercent > .ScaleWidth Then
                lngPercent = .ScaleWidth
            End If
               
            .Cls                        ' Empty picture box
            .ForeColor = lngForeColor   ' Reset forecolor
         
            ' set picture box ScaleWidth equal to maximum percentage
            .ScaleWidth = MAX_PERCENT
            
            ' format percent into a displayable value (ex: 25%)
            strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
            
            ' Calculate X and Y coordinates within
            ' picture box and and center data
            .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
                
            objProgBar.Print strPercent   ' print percentage string in picture box
            
            ' Print flood bar up to new percent position in picture box
            objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
        
        End With
                
        DoEvents   ' allow flood to complete drawing
    
    End If

End Sub

Private Sub GetRegistryData()

    mstrFolder = GetSetting("kiCrypt", "Settings", "LastPath", App.Path & "\")
    mblnCreateNewFile = GetSetting("kiCrypt", "Settings", "OverwriteFile", True)
    mlngKeyMix = GetSetting("kiCrypt", "Settings", "PrimaryKeyMix", "1")
    
End Sub

Private Sub UpdateRegistry()

    SaveSetting "kiCrypt", "Settings", "LastPath", mstrFolder
    SaveSetting "kiCrypt", "Settings", "OverwriteFile", mblnCreateNewFile
    SaveSetting "kiCrypt", "Settings", "PrimaryKeyMix", mlngKeyMix

End Sub

Private Sub txtKeyMix_GotFocus()
    ' Highlight contents in text box
    mobjKeyEdit.TextBoxFocus txtKeyMix
End Sub

Private Sub txtKeyMix_KeyDown(KeyCode As Integer, Shift As Integer)
    ' key control (Ex:   Ctrl+C, etc.)
    mobjKeyEdit.TextBoxKeyDown txtKeyMix, KeyCode, Shift
End Sub

Private Sub txtKeyMix_KeyPress(KeyAscii As Integer)
    ' edit data input
    mobjKeyEdit.ProcessNumericOnly KeyAscii
End Sub

Private Sub txtKeyMix_LostFocus()

    txtKeyMix.Text = Trim$(txtKeyMix.Text)
    
    If Len(txtKeyMix.Text) = 0 Or _
       Val(txtKeyMix.Text) < 1 Or _
       Val(txtKeyMix.Text) > 5 Then
       
        mlngKeyMix = 1
        txtKeyMix.Text = mlngKeyMix
    Else
        mlngKeyMix = Val(txtKeyMix.Text)
    End If
    
End Sub

Private Sub txtPwd_GotFocus()
    ' Highlight contents in text box
    mobjKeyEdit.TextBoxFocus txtPwd
End Sub

Private Sub txtPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    ' key control (Ex:   Ctrl+C, etc.)
    mobjKeyEdit.TextBoxKeyDown txtPwd, KeyCode, Shift
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    ' edit data input
    mobjKeyEdit.ProcessAlphaNumeric KeyAscii
End Sub

Private Sub txtPwd_LostFocus()
    txtPwd.Text = Trim$(txtPwd.Text)
    mstrPassword = txtPwd.Text
End Sub

Private Sub txtInputString_GotFocus()
    ' Highlight contents in text box
    mobjKeyEdit.TextBoxFocus txtInputString
End Sub

Private Sub txtInputString_KeyDown(KeyCode As Integer, Shift As Integer)
    ' key control (Ex:   Ctrl+C, etc.)
    mobjKeyEdit.TextBoxKeyDown txtInputString, KeyCode, Shift
End Sub

Private Sub txtInputString_KeyPress(KeyAscii As Integer)
    
    ' edit data input
    Select Case KeyAscii
           Case 9
                ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
                
           Case 8, 13, 32 To 126
                ' Backspace, ENTER key and
                ' other valid data keys
                
           Case Else  ' Everything else (invalid)
                KeyAscii = 0
    End Select

End Sub

