VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "IntraDream - INap"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIgnore 
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList imgPlay 
      Left            =   0
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   1320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   600
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrDL 
      Interval        =   30000
      Left            =   4320
      Top             =   4320
   End
   Begin MSComctlLib.ImageList imgSpeed 
      Left            =   4800
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4688
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5316
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   2000
      Left            =   5880
      Top             =   4200
   End
   Begin MSWinsockLib.Winsock Transfer 
      Index           =   0
      Left            =   6360
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6699
   End
   Begin VB.Timer tmrStatus 
      Interval        =   10000
      Left            =   5400
      Top             =   4200
   End
   Begin MSComctlLib.StatusBar S1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4815
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            Text            =   "Servers : 0"
            TextSave        =   "Servers : 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            Text            =   "Users : 0"
            TextSave        =   "Users : 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            Text            =   "Files : 0"
            TextSave        =   "Files : 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3916
            Text            =   "Gigabytes : 0"
            TextSave        =   "Gigabytes : 0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   582
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect   "
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat      "
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search    "
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Download "
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Upload    "
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hotlist    "
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Library    "
            ImageIndex      =   9
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8160
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   6840
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7440
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":822C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F4A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FA3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1425E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":145F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14992
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   3180
      Index           =   1
      Left            =   0
      ScaleHeight     =   3180
      ScaleWidth      =   5580
      TabIndex        =   1
      Top             =   360
      Width           =   5580
      Begin MSComctlLib.ListView lvSL 
         Height          =   2055
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Network"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVC 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Server"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Users"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Files"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Network"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   6
      Left            =   0
      ScaleHeight     =   2700
      ScaleWidth      =   8340
      TabIndex        =   6
      Top             =   360
      Width           =   8340
      Begin MSComctlLib.ListView LVU 
         Height          =   2535
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FileName"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Speed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   7
      Left            =   0
      ScaleHeight     =   3780
      ScaleWidth      =   9180
      TabIndex        =   3
      Top             =   360
      Width           =   9180
      Begin MSComctlLib.ListView LVIC 
         Height          =   1815
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgSpeed"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Online"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVNC 
         Height          =   1815
         Left            =   0
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgSpeed"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Offline"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVB 
         Height          =   1815
         Left            =   1800
         TabIndex        =   21
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filename"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Speed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Bitrate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Size"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   0
      ScaleHeight     =   3780
      ScaleWidth      =   9180
      TabIndex        =   5
      Top             =   360
      Width           =   9180
      Begin VB.TextBox txtArtist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   11
         Text            =   "Artist"
         Top             =   30
         Width           =   2535
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   30
         Width           =   1215
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   1815
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgSpeed"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filename"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Speed"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Bitrate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Size"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LVD 
      Height          =   3855
      Left            =   0
      TabIndex        =   16
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FileName"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Speed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Size"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   3900
      Index           =   8
      Left            =   0
      ScaleHeight     =   3900
      ScaleWidth      =   9180
      TabIndex        =   13
      Top             =   360
      Width           =   9180
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Min             =   -2500
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.ListView LVS 
         Height          =   2415
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgPlay"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Folder"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Length"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bitrate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MediaPlayerCtl.MediaPlayer M1 
         Height          =   675
         Left            =   0
         TabIndex        =   18
         Top             =   3120
         Width           =   4455
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   1
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   0   'False
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   0   'False
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   -1  'True
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   0
      ScaleHeight     =   3780
      ScaleWidth      =   9180
      TabIndex        =   4
      Top             =   360
      Width           =   9180
      Begin Napster.Room Room 
         Height          =   2880
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   15
         Visible         =   0   'False
         Width           =   7095
         _extentx        =   12515
         _extenty        =   5080
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5741
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":14D2C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnudiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug Mode"
      End
   End
   Begin VB.Menu mnuChannels 
      Caption         =   "Channels"
      Begin VB.Menu mnuRooms 
         Caption         =   "Get Rooms"
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrivate 
         Caption         =   "Private"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRoom 
         Caption         =   "Room"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuD 
      Caption         =   "D"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuCDC 
         Caption         =   "Clear Complete"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDPM 
         Caption         =   "Private Message"
      End
      Begin VB.Menu mnuATHL 
         Caption         =   "Add To Hotlist"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browse Files"
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "mnuH"
      Visible         =   0   'False
      Begin VB.Menu mnuBrowseHF 
         Caption         =   "Browse Files"
      End
      Begin VB.Menu mnuPMHL 
         Caption         =   "Private Message"
      End
      Begin VB.Menu mnuRMHL 
         Caption         =   "Remove From Hotlist"
      End
   End
   Begin VB.Menu mnuHD 
      Caption         =   "mnuHD"
      Visible         =   0   'False
      Begin VB.Menu mnuRMHDL 
         Caption         =   "Remove From Hotlist"
      End
   End
   Begin VB.Menu mnuPrefs 
      Caption         =   "View"
      Begin VB.Menu mnuOPT 
         Caption         =   "Options"
      End
      Begin VB.Menu mnudash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDLD 
         Caption         =   "Download Directory"
      End
      Begin VB.Menu mnudash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIL 
         Caption         =   "Ignore List"
      End
   End
   Begin VB.Menu mnuC 
      Caption         =   "mnuc"
      Visible         =   0   'False
      Begin VB.Menu mnuGNL 
         Caption         =   "Get Napigator List"
      End
      Begin VB.Menu mnudash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnects 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDisconnects 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuCSrv 
         Caption         =   "Servers"
         Begin VB.Menu mnuDaServer 
            Caption         =   "Server"
            Index           =   0
         End
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSC 
         Caption         =   "Send /Command"
      End
   End
   Begin VB.Menu mnuMedia 
      Caption         =   "Media"
      Begin VB.Menu mnuPlaya 
         Caption         =   "Play"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStopa 
         Caption         =   "Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu MNUDASH9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNexter 
         Caption         =   "Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPrevi 
         Caption         =   "Previous"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuSite 
         Caption         =   "IntraDream.com"
      End
   End
   Begin VB.Menu mnuU 
      Caption         =   "mnuU"
      Visible         =   0   'False
      Begin VB.Menu mnuUCan 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUPM 
         Caption         =   "Private Message"
      End
      Begin VB.Menu mnuUBrowse 
         Caption         =   "Browse Files"
      End
      Begin VB.Menu mnuUWhois 
         Caption         =   "WhoIs"
      End
   End
   Begin VB.Menu mnuPl 
      Caption         =   "mnuPL"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuRL 
         Caption         =   "Refresh List"
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "mnuIcon"
      Visible         =   0   'False
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
      End
      Begin VB.Menu mnuPrevius 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnudash8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Toggle View"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Sub SendData(msg As String, What As String, Index As Integer)
   ' MsgBox Sock(Index).State
    If Sock(Index).State <> 7 Then Exit Sub
    Dim LENGTH As String
    LENGTH = GetLenChr(Len(msg))
    Sock(Index).SendData LENGTH
    Sock(Index).SendData What
    Sock(Index).SendData msg
End Sub

Function GetLenChr(LENGTH As Long)
    Dim X As Integer
    Do While LENGTH > 255
        X = X + 1
        LENGTH = LENGTH - 255
    Loop
    GetLenChr = Chr(LENGTH) & Chr(X)
End Function


Private Sub cmdSearch_Click()
    '200 (0xc8)  client search request [CLIENT]
    LV1.ListItems.Clear
    For i = 0 To Sock.UBound
    If Sock(i).State = 7 Then
        SendData "FILENAME CONTAINS """ & txtArtist & """ MAX_RESULTS 500", Chr(200) & Chr(0), Int(i)
    End If
    Next
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon Me.Icon
    gSysTray.IconInSysTray
    gSysTray.ToolTip = "INap, Internet Music Community"
    Slider1.Value = M1.Volume
    Open App.path & "\Options\Buddy.txt" For Input As #254
        Do Until EOF(254)
            Dim BString As String
            Line Input #254, BString
            LVNC.ListItems.Add , BString, BString, , 6
        Loop
    Close #254
    Open App.path & "\options\Ignore.txt" For Input As #1
        Do Until EOF(1)
            Dim a As String
            Line Input #1, a
            lstIgnore.AddItem a
        Loop
    Close #1
    'Client = "HyperNap" '"audioGnome" '"XNap 2.4-pre5"
    'frmInfo.Show
    Dim X As String
    TB1.Buttons.Item(1).Value = tbrPressed
    P(1).ZOrder 0
    'Open App.path & "\share.txt" For Input As #1
    'Dim os As Long
    'Do Until EOF(Index + 1) Or os > 500
    '    os = os + 1
    '    Line Input #1, x
    '    LVS.ListItems.Add , , x
    '    DoEvents
    'Loop
    'Close #1
    For i = 1 To 50
        Load Room(i)
        Load mnuRoom(i)
    Next

    'x = Inet1.OpenURL("http://www.napigator.com/servers/")
    'Dim Y() As String
    'Dim z() As String
    'Dim c() As String
    'Dim d() As String
    'Y = Split(x, "<td align=""left"" bgcolor=""#93b85f""><font face=""Tahoma, Verdana"" size=""2"">")
    'For i = 1 To UBound(Y) Step 2
    '    z = Split(Y(i), "</font>")
    '    c = Split(z(1), "<td align=""center"" bgcolor=""#93b85f""><font face=""Tahoma, Verdana"" size=""2"">")
    '    d = Split(c(1), "</font>")
    '    'Text1.Text = Text1.Text & vbCrLf & z(0) & ":" & d(0)
    '    LVC.ListItems.Add , "x" & LVC.ListItems.Count + 1, z(0) & ":" & d(0)
    '    LVC.ListItems("x" & LVC.ListItems.Count).SubItems(1) = "Disconnected..."
    '    LVC.ListItems("x" & LVC.ListItems.Count).SubItems(2) = "0"
    '    LVC.ListItems("x" & LVC.ListItems.Count).SubItems(3) = "0"
    '    LVC.ListItems("x" & LVC.ListItems.Count).SubItems(4) = "0"
    '    Load Sock(LVC.ListItems.Count)
    '    ReDim Preserve Buffer(LVC.ListItems.Count)
    '    ReDim Preserve Busy(LVC.ListItems.Count)
    'Next
End Sub

Public Sub LoadGator()
On Error Resume Next
X = Inet1.OpenURL("http://www.napigator.com/servers.php?version=208&client=napigator&rnd=11485")
    Dim Y() As String
    Dim z() As String
    Y = Split(X, Chr(10))
    For i = 0 To UBound(Y)
        z = Split(Y(i), " ")
        If InStr(z(0), ".") And IsNumeric(z(1)) Then
        lvSL.ListItems.Add , z(0) & ":" & z(1), z(0) & ":" & z(1)
        lvSL.ListItems(z(0) & ":" & z(1)).SubItems(1) = z(2)
        If LCase(z(2)) = "n/a" Then 'Dosnt have a network no need to watch for linked servers
            z(2) = z(0) & ":" & z(1)
        End If
        DoEvents
        Err.Clear
        LVC.ListItems.Add , LCase(z(2)), z(0) & ":" & z(1)
        If Err.Description > "" Then GoTo noadd
        LVC.ListItems(LCase(z(2))).SubItems(1) = "Disconnected"
        LVC.ListItems(LCase(z(2))).SubItems(2) = "0"
        LVC.ListItems(LCase(z(2))).SubItems(3) = "0"
        LVC.ListItems(LCase(z(2))).SubItems(4) = "0"
        LVC.ListItems(LCase(z(2))).SubItems(5) = z(2)
        lvSL.ListItems.Remove lvSL.ListItems(z(0) & ":" & z(1)).Index
noadd:
        End If
On Error Resume Next
    Next
End Sub
Sub AddText(msg As String, Vcol As ColorConstants)
    RTB.SelStart = Len(RTB.Text)
    RTB.SelColor = Vcol
    RTB.SelText = vbCrLf & msg
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '1, 4, 5, 6, 7
    On Error Resume Next
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Exit Sub
    End If
    For i = 1 To 8
        P(i).Width = Me.ScaleWidth
        P(i).Height = Me.ScaleHeight - TB1.Height - S1.Height
    Next
    For i = 0 To 50
        If Room(i).Visible = True Then
            Room(i).Width = P(1).Width
            Room(i).Height = P(1).Height
        End If
    Next
    LVIC.Height = P(1).Height / 2
    LVNC.Height = P(1).Height / 2
    LVNC.TOP = LVIC.Height
    LVB.Width = P(1).Width - LVIC.Width
    LVB.Height = P(1).Height
    RTB.Height = P(1).Height
    RTB.Width = P(1).Width
    LVC.Height = P(1).Height
    LVC.Width = P(1).Width
    LVD.Height = P(1).Height
    LVD.Width = P(1).Width
    LVU.Height = P(1).Height
    LVU.Width = P(1).Width
    LVS.Height = P(1).Height - M1.Height
    M1.TOP = LVS.Height
    M1.Width = P(1).Width
    LVS.Width = P(1).Width
    Slider1.TOP = M1.TOP + 300
    Slider1.Left = P(1).Width - Slider1.Width - 50
    'LVS.ColumnHeaders.Item(1).Width = LVS.Width - 300
    LV1.Height = Me.ScaleHeight - S1.Height - cmdSearch.Height - TB1.Height - 50
    LV1.Width = LVC.Width
    cmdSearch.Left = LVC.Width - cmdSearch.Width
    txtArtist.Width = cmdSearch.Left - 50
    SizeBy LVD, 2
    SizeBy LVC, 2
    SizeBy LVB, 2
    SizeBy LVU, 2
    SizeBy LV1, 2
    SizeBy LVS, 1
    SizeBy LVIC, 1
    SizeBy LVNC, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gSysTray.RemoveFromSysTray
    Open App.path & "\Options\Buddy.txt" For Output As #254
        For i = 1 To LVIC.ListItems.Count
            Print #254, LVIC.ListItems.Item(i)
        Next
        For i = 1 To LVNC.ListItems.Count
            Print #254, LVNC.ListItems.Item(i)
        Next
    Close #254
    Open App.path & "\Options\Ignore.txt" For Output As #254
        For i = 0 To lstIgnore.ListCount - 1
            Print #254, lstIgnore.List(i)
        Next
    Close #254
    End
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub gSysTray_RButtonUp()
    PopupMenu mnuIcon
End Sub

Private Sub LV1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LV1.SortOrder = lvwAscending Then
        LV1.SortOrder = lvwDescending
    Else
        LV1.SortOrder = lvwAscending
    End If
    LV1.SortKey = ColumnHeader.Index - 1
    LV1.Sorted = True
    LV1.Sorted = False
End Sub

Private Sub LV1_DblClick()
    On Error Resume Next
    If Sock(LV1.SelectedItem.Tag).State = 7 Then
        LVD.ListItems.Add , LV1.SelectedItem.ToolTipText & LV1.SelectedItem, LV1.SelectedItem
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).SubItems(1) = LV1.SelectedItem.SubItems(1)
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).SubItems(2) = "Resolving Host"
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).SubItems(3) = "0 B/s"
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).SubItems(4) = LV1.SelectedItem.SubItems(5)
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).Tag = LV1.SelectedItem.Tag
        LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).ToolTipText = LV1.SelectedItem.ToolTipText
        LV1.SelectedItem.SmallIcon = 5
        'SetProgress LVD.ListItems(LV1.SelectedItem.ToolTipText & LV1.SelectedItem).Index, Int(0)
        SendData LV1.SelectedItem & " """ & LV1.SelectedItem.ToolTipText & """", Chr(203) & Chr(0), LV1.SelectedItem.Tag
        tmrDL.Enabled = False
        tmrDL.Enabled = True
    Else
        MsgBox "Disconnected from that server!"
    End If
End Sub

Private Sub LVB_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub LVB_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LVB.SortOrder = lvwAscending Then
        LVB.SortOrder = lvwDescending
    Else
        LVB.SortOrder = lvwAscending
    End If
    LVB.SortKey = ColumnHeader.Index - 1
    LVB.Sorted = True
    LVB.Sorted = False
End Sub

Private Sub LVB_DblClick()
    On Error Resume Next
    If Sock(LVB.SelectedItem.Tag).State = 7 Then
        LVD.ListItems.Add , LVB.SelectedItem.ToolTipText & LVB.SelectedItem, LVB.SelectedItem
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).SubItems(1) = LVB.SelectedItem.SubItems(1)
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).SubItems(2) = "Resolving Host"
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).SubItems(3) = "0 B/s"
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).SubItems(4) = LVB.SelectedItem.SubItems(5)
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).Tag = LVB.SelectedItem.Tag
        LVD.ListItems(LVB.SelectedItem.ToolTipText & LVB.SelectedItem).ToolTipText = LVB.SelectedItem.ToolTipText
        'SetProgress LVD.ListItems(lvb.SelectedItem.ToolTipText & lvb.SelectedItem).Index, Int(0)
        SendData LVB.SelectedItem & " """ & LVB.SelectedItem.ToolTipText & """", Chr(203) & Chr(0), LVB.SelectedItem.Tag
        tmrDL.Enabled = False
        tmrDL.Enabled = True
    Else
        MsgBox "Disconnected from that server!"
    End If
End Sub

Private Sub LVC_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LVC.SortOrder = lvwAscending Then
        LVC.SortOrder = lvwDescending
    Else
        LVC.SortOrder = lvwAscending
    End If
    LVC.SortKey = ColumnHeader.Index - 1
    LVC.Sorted = True
    LVC.Sorted = False
End Sub

Private Sub LVC_DblClick()
    On Error Resume Next
    ConnectNum LVC.SelectedItem.Index
End Sub

Private Sub LVC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim mm As Integer
    mm = mnuDaServer.LBound
    For i = mnuDaServer.LBound To mnuDaServer.UBound
        mnuDaServer(i).Visible = False
        mnuDaServer(mm).Caption = ""
    Next
    For i = 1 To lvSL.ListItems.Count
        If LCase(LVC.SelectedItem.Key) = LCase(lvSL.ListItems.Item(i).SubItems(1)) Then
            Load mnuDaServer(mm)
            mnuDaServer(mm).Caption = lvSL.ListItems.Item(i)
            mnuDaServer(mm).Visible = True
            mm = mm + 1
        End If
    Next
    If mm = 0 Then
        mnuCSrv.Visible = False
    Else
        mnuCSrv.Visible = True
    End If
    If Button = 2 Then PopupMenu mnuC
End Sub

Private Sub LVD_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LVD.SortOrder = lvwAscending Then
        LVD.SortOrder = lvwDescending
    Else
        LVD.SortOrder = lvwAscending
    End If
    LVD.SortKey = ColumnHeader.Index - 1
    LVD.Sorted = True
    LVD.Sorted = False
End Sub

Private Sub LVD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuD
End Sub

Private Sub LVS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPl
End Sub

Private Sub LVu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuU
End Sub

Private Sub LVIC_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then PopupMenu mnuH
End Sub

Private Sub LVNC_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVNC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then PopupMenu mnuHD
End Sub

Private Sub LVS_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub LVS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LVS.SortOrder = lvwAscending Then
        LVS.SortOrder = lvwDescending
    Else
        LVS.SortOrder = lvwAscending
    End If
    LVS.SortKey = ColumnHeader.Index - 1
    LVS.Sorted = True
    LVS.Sorted = False
End Sub

Private Sub LVS_DblClick()
    On Error Resume Next
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub LVU_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub M1_EndOfStream(ByVal Result As Long)
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index + 1).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub mnuATHL_Click()
    'add a user to hotlist
    On Error GoTo Err
    LVNC.ListItems.Add , LVD.SelectedItem, LVD.SelectedItem, , 6
    For i = 1 To Sock.UBound
        If Sock(i).State = 7 Then SendData LVD.SelectedItem, NumToChr(207), Int(i)
    Next
Err:
End Sub

Private Sub mnuBrowse_Click()
    '211 nick
    On Error Resume Next
    frmMain.LVB.ListItems.Clear
    frmMain.SendData LVD.SelectedItem, NumToChr(211), LVD.SelectedItem.Tag
End Sub

Private Sub mnuBrowseHF_Click()
    '211 nick
    On Error Resume Next
    frmMain.LVB.ListItems.Clear
    frmMain.SendData LVIC.SelectedItem, NumToChr(211), LVIC.SelectedItem.Tag
End Sub

Private Sub mnuCancel_Click()
    On Error Resume Next
    If Left(LVD.SelectedItem.Key, 1) = "x" And Len(LVD.SelectedItem.Key) < 4 Then
        Transfer_Close Right(LVD.SelectedItem.Key, Len(LVD.SelectedItem.Key) - 1)
    End If
    LVD.ListItems.Remove LVD.SelectedItem.Index
End Sub

Private Sub mnuCDC_Click()
    On Error Resume Next
    For i = 1 To LVD.ListItems.Count
        If LVD.ListItems.Item(i).SubItems(2) = "Completed" Then
            LVD.ListItems.Remove i
        End If
    Next
End Sub

Private Sub mnuConnect_Click()
    On Error Resume Next
    For i = LVC.ListItems.Count To 1 Step -1
        DoEvents
        ConnectNum Int(i)
    Next
End Sub

Sub ConnectNum(Index As Integer)
    On Error Resume Next
    Dim coo As Integer
    coo = 0
    coo = UBound(Network)
    If coo > 0 Then
        For i = 0 To UBound(Network)
            If Network(i) = LVC.ListItems.Item(Index).Key Then Exit Sub
        Next
    End If
    Dim X As Integer
    X = OpenSocket("Sock", Me)
    ReDim Preserve Network(Sock.UBound)
    ReDim Preserve Buffer(Sock.UBound)
    ReDim Preserve Busy(Sock.UBound)
    
    Network(X) = LVC.ListItems.Item(Index).Key
    Dim CI() As String
    CI = Split(LVC.ListItems.Item(Index), ":")
    LVC.ListItems.Item(Index).SubItems(1) = "Connecting"
    LVC.ListItems.Item(Index).SubItems(2) = "0"
    LVC.ListItems.Item(Index).SubItems(3) = "0"
    LVC.ListItems.Item(Index).SubItems(4) = "0"
    DoEvents
    Sock(X).Connect CI(0), CI(1)
End Sub


Private Sub mnuConnects_Click()
    On Error Resume Next
    ConnectNum LVC.SelectedItem.Index
End Sub

Private Sub mnuDaServer_Click(Index As Integer)
On Error Resume Next
mnuDisconnects_Click
For i = 1 To lvSL.ListItems.Count
    If lvSL.ListItems.Item(i) = mnuDaServer(Index).Caption Then
        lvSL.ListItems.Item(i) = LVC.SelectedItem
        LVC.SelectedItem = mnuDaServer(Index).Caption
    End If
Next
End Sub

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not (mnuDebug.Checked)
End Sub

Private Sub mnuDisconnects_Click()
    On Error Resume Next
    For i = 1 To Sock.UBound
        If Network(i) = LCase(LVC.SelectedItem.Key) Then
            LVC.ListItems(Network(i)).SubItems(1) = "Disconnected"
            LVC.ListItems(Network(i)).SubItems(2) = "0"
            LVC.ListItems(Network(i)).SubItems(3) = "0"
            LVC.ListItems(Network(i)).SubItems(4) = "0"
            Sock(i).Close
            Network(i) = ""
                For j = 0 To Room.UBound
                    If Room(j).CServer = i Then
                        Room(j).LeaveRoom
                    End If
                Next
                For j = 1 To LVIC.ListItems.Count
                    If LVIC.ListItems.Item(j).Tag = i Then
                            On Error Resume Next
                            For f = 1 To frmMain.Sock.UBound
                                If frmMain.Sock(f).State = 7 Then frmMain.SendData LVIC.ListItems.Item(j), NumToChr(303), Int(f)
                                If frmMain.Sock(f).State = 7 Then frmMain.SendData LVIC.ListItems.Item(j), NumToChr(207), Int(f)
                            Next
                            LVNC.ListItems.Add , LVIC.ListItems.Item(j), LVIC.ListItems.Item(j), , 6
                            LVIC.ListItems.Remove j
                    End If
                Next
            Exit For
        End If
    Next
End Sub

Private Sub mnuDLD_Click()
Shell "explorer.exe """ & Replace(DownloadDir, "\\", "\") & """", vbMaximizedFocus
End Sub

Private Sub mnuDPM_Click()
On Error Resume Next
            Fchan = PMtoChan(CStr(LVD.SelectedItem), Int(LVD.SelectedItem.Tag))
            If Fchan > 50 Then
                Fchan = OpenPM
                If Fchan < 51 Then
                    PM(Fchan).Show
                    PM(Fchan).CServer = Int(LVD.SelectedItem.Tag)
                    PM(Fchan).CName = LVD.SelectedItem
                    PM(Fchan).Caption = "(" & LVD.SelectedItem & ") Instant Message"
                End If
            Else
                    PM(Fchan).Caption = "(" & LVD.SelectedItem & ") Instant Message"
                    PM(Fchan).WindowState = vbNormal
                    PM(Fchan).ZOrder 0
                    
            End If
End Sub

Private Sub mnuGNL_Click()
    frmShare.Show , frmMain
    frmShare.lblShare.Caption = "Geting Server List!!"
    LoadGator
    Unload frmShare
End Sub

Private Sub mnuIL_Click()
    frmIgnore.Show , Me
    frmIgnore.List1.Clear
    For i = 0 To lstIgnore.ListCount - 1
        frmIgnore.List1.AddItem lstIgnore.List(i)
    Next
End Sub

Private Sub mnuNext_Click()
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index + 1).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub mnuNexter_Click()
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index + 1).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub mnuOPT_Click()
    frmOptions.Show
    frmOptions.IsOpen = True
End Sub

Private Sub mnuPlay_Click()
    On Error Resume Next
    LVS_DblClick
End Sub

Private Sub mnuPlaya_Click()
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub mnuPMHL_Click()
On Error Resume Next
            Fchan = PMtoChan(CStr(LVIC.SelectedItem), LVIC.SelectedItem.Tag)
            If Fchan > 50 Then
                Fchan = OpenPM
                If Fchan < 51 Then
                    PM(Fchan).Show
                    PM(Fchan).CServer = LVIC.SelectedItem.Tag
                    PM(Fchan).Caption = "(" & LVIC.SelectedItem & ") Instant Message"
                    PM(Fchan).CName = LVIC.SelectedItem
                End If
            Else
                    PM(Fchan).Caption = "(" & LVIC.SelectedItem & ") Instant Message"
            End If
End Sub

Private Sub mnuPrevi_Click()
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index - 1).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Private Sub mnuPrevius_Click()
    On Error Resume Next
    If LVS.SelectedItem.Index = LVS.ListItems.Count Then
        LVS.ListItems.Item(1).Selected = True
    Else
        LVS.ListItems.Item(LVS.SelectedItem.Index - 1).Selected = True
    End If
    M1.FileName = LVS.SelectedItem.Key
End Sub

Public Sub mnuPrivate_Click()
    On Error Resume Next
    RTB.ZOrder 0
    mnuPrivate.Checked = True
    For i = 0 To 50
        mnuRoom(i).Checked = False
    Next
End Sub

Private Sub mnuRL_Click()
        On Error Resume Next
        frmShare.Show
        frmMain.LVS.ListItems.Clear
        For i = 0 To List1.ListCount - 1
        DirSize List1.List(i)
        Next
        Unload frmShare
End Sub

Private Sub mnuRMHDL_Click()
    On Error Resume Next
                For i = 1 To frmMain.Sock.UBound
                    If frmMain.Sock(i).State = 7 Then frmMain.SendData LVNC.SelectedItem, NumToChr(303), Int(i)
                    'If frmMain.Sock(i).State = 7 Then frmMain.SendData LVNC.SelectedItem, NumToChr(207), Int(i)
                Next
                LVNC.ListItems.Remove LVNC.SelectedItem.Index
End Sub

Private Sub mnuRMHL_Click()
    On Error Resume Next
    
                For i = 1 To frmMain.Sock.UBound
                    If frmMain.Sock(i).State = 7 Then frmMain.SendData LVIC.SelectedItem, NumToChr(303), Int(i)
                    'If frmMain.Sock(i).State = 7 Then frmMain.SendData LVIC.SelectedItem, NumToChr(207), Int(i)
                Next
                'LVNC.ListItems.Add , LVIC.SelectedItem, LVIC.SelectedItem, , 6
                LVIC.ListItems.Remove LVIC.SelectedItem.Index
End Sub

Public Sub mnuRoom_Click(Index As Integer)
    On Error Resume Next
    mnuPrivate.Checked = False
    For i = 0 To 50
        mnuRoom(i).Checked = False
    Next
    Room(Index).ZOrder 0
    mnuRoom(Index).Checked = True
                frmMain.P(2).ZOrder 0
                For i = 1 To frmMain.TB1.Buttons.Count
                frmMain.TB1.Buttons.Item(i).Value = tbrUnpressed
                Next
                
                frmMain.TB1.Buttons.Item(3).Value = tbrPressed
End Sub

Private Sub mnuRooms_Click()
    frmRooms.Show
    For i = 1 To Sock.UBound
        If Sock(i).State = 7 Then
            SendData Chr(0), Chr(59) & Chr(3), Int(i)
            'SendData Chr(0), Chr(105) & Chr(2), Int(i)
        End If
    Next
End Sub

Private Sub mnuSC_Click()
    On Error Resume Next
    Dim xu As Integer
    xu = 0
    xu = LVC.SelectedItem.Index
    If xu = 0 Then Exit Sub
    xu = 0
    For i = 0 To Sock.UBound
        If Network(i) = LVC.SelectedItem.Key Then
            xu = i
            Exit For
        End If
    Next
    If xu = 0 Then
        MsgBox "Not connected to the " & LVC.SelectedItem.Key & " network.", vbCritical, "Error"
    Else
        Dim Cmds As String
        Cmds = InputBox("Enter command to run.", "/Command", "/join")
        If Len(Cmds) > 0 Then
            If IsCommand(Cmds, "0", xu) = False Then
                MsgBox "Unknown Command!", vbCritical, "Error"
            End If
        End If
    End If
End Sub

Private Sub mnuSite_Click()
Shell "explorer.exe ""http://www.intradream.com/inap/Help/index.html"""
End Sub

Private Sub mnuStopa_Click()
    M1.Stop
End Sub

Private Sub mnuUBrowse_Click()
    On Error Resume Next
    frmMain.LVB.ListItems.Clear
    frmMain.SendData LVU.SelectedItem, NumToChr(211), Int(LVU.SelectedItem.Tag)
End Sub

Private Sub mnuUCan_Click()
    On Error Resume Next
    If Left(LVU.SelectedItem.Key, 1) = "x" And Len(LVU.SelectedItem.Key) < 4 Then
        Transfer_Close Right(LVU.SelectedItem.Key, Len(LVU.SelectedItem.Key) - 1)
    End If
    LVU.ListItems.Remove LVU.SelectedItem.Index
End Sub

Private Sub mnuUPM_Click()
On Error Resume Next
            Fchan = PMtoChan(CStr(LVU.SelectedItem), Int(LVU.SelectedItem.Tag))
            If Fchan > 50 Then
                Fchan = OpenPM
                If Fchan < 51 Then
                    PM(Fchan).Show
                    PM(Fchan).CServer = Int(LVU.SelectedItem.Tag)
                    PM(Fchan).CName = LVU.SelectedItem
                    PM(Fchan).Caption = "(" & LVU.SelectedItem & ") Instant Message"
                End If
            Else
                    PM(Fchan).Caption = "(" & LVU.SelectedItem & ") Instant Message"
                    PM(Fchan).WindowState = vbNormal
                    PM(Fchan).ZOrder 0
                    
            End If
End Sub

Private Sub mnuUWhois_Click()
    On Error Resume Next
    SendData LVU.SelectedItem, Chr(91) & Chr(2), Int(LVU.SelectedItem.Tag)
End Sub


Private Sub mnuView_Click()
        If Me.WindowState = vbMinimized Then
            Me.WindowState = vbNormal
            Me.Show
            App.TaskVisible = True
        Else
            Me.WindowState = vbMinimized
            Me.Hide
        End If
End Sub

Private Sub Slider1_Change()
    M1.Volume = Slider1.Value
End Sub

Private Sub Slider1_Click()
    M1.Volume = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
    M1.Volume = Slider1.Value
End Sub

Public Sub Sock_Close(Index As Integer)
    On Error Resume Next
    Dim netW As String
    LVC.ListItems.Remove LVC.ListItems(Network(Index)).Index
    Sock(Index).Close
    For i = 0 To Room.UBound
        If Room(i).CServer = Index Then
            Room(i).LeaveRoom
        End If
    Next
    For i = 1 To LVIC.ListItems.Count
        If LVIC.ListItems.Item(i).Tag = Index Then
                On Error Resume Next
                For j = 1 To frmMain.Sock.UBound
                    If frmMain.Sock(j).State = 7 Then frmMain.SendData LVIC.ListItems.Item(i), NumToChr(303), Int(j)
                    If frmMain.Sock(j).State = 7 Then frmMain.SendData LVIC.ListItems.Item(i), NumToChr(207), Int(j)
                Next
                LVNC.ListItems.Add , LVIC.ListItems.Item(i), LVIC.ListItems.Item(i), , 6
                LVIC.ListItems.Remove i
        End If
    Next
    netW = Network(Index)
    Network(Index) = ""
    For i = 1 To lvSL.ListItems.Count
        If LCase(lvSL.ListItems.Item(i).SubItems(1)) = netW Then
            LVC.ListItems.Add , netW, lvSL.ListItems.Item(i)
            LVC.ListItems(netW).SubItems(1) = "Disconnected"
            LVC.ListItems(netW).SubItems(2) = "0"
            LVC.ListItems(netW).SubItems(3) = "0"
            LVC.ListItems(netW).SubItems(4) = "0"
            LVC.ListItems(netW).SubItems(5) = lvSL.ListItems.Item(i).SubItems(1)
            lvSL.ListItems.Remove i
            ConnectNum Int(LVC.ListItems(netW).Index)
            Exit Sub
        End If
    Next
End Sub

Private Sub Sock_Connect(Index As Integer)
    'On Error Resume Next
    Dim ConInfo As String
    'AddText "*Connected", vbRed

    LVC.ListItems(Network(Index)).SubItems(1) = "Connected..."
    ConInfo = Nickname & " " & Password & " " & Port & " """ & Client & """ " & Speed
    SendData ConInfo, Chr(2) & Chr(0), Index
    DoEvents

    'S1.Panels.Item(1).Text = "Connected(" & Nickname & "): Sharing " & Share & " Files"
End Sub

Sub SendShare(Index As Integer)
    For i = 1 To LVS.ListItems.Count
        SendData LVS.ListItems.Item(i).Tag, Chr(100) & Chr(0), Index
        DoEvents
    Next
End Sub
Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
    Dim Sata As String
    Sock(Index).GetData Sata
    Sata = Buffer(Index) & Sata
    Buffer(Index) = ""
    If Sata > 0 And Busy(Index) = False Then
        HandleData Sata, Index
    End If
End Sub

Private Sub Sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Sock_Close Index
End Sub

Private Sub TB1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            P(Button.Index).ZOrder 0
        Case 3
             P(Button.Index - 1).ZOrder 0
        Case 5
             P(Button.Index - 2).ZOrder 0
        Case 9
             P(Button.Index - 3).ZOrder 0
        Case 11
             P(Button.Index - 4).ZOrder 0
        Case 13
             P(Button.Index - 5).ZOrder 0
        Case 7
            LVD.ZOrder 0
        Case 9
        Case 10
    End Select
    For i = 1 To TB1.Buttons.Count
    TB1.Buttons.Item(i).Value = tbrUnpressed
    Next
    TB1.Buttons.Item(Button.Index).Value = tbrPressed
End Sub

Private Sub tmrDL_Timer()
    'gets reset every new d/l
    'when it fires it resends info for any d/ls that havnt started.
    On Error Resume Next
    For i = 1 To LVD.ListItems.Count
        If Not Left(LVD.ListItems.Item(i).Key, 1) = "x" And Len(LVD.ListItems.Item(i).Key) > 3 Then
            LVD.ListItems.Item(i).SubItems(2) = "Resolving Host"
            SendData LVD.ListItems.Item(i) & " """ & LVD.ListItems.Item(i).ToolTipText & """", Chr(203) & Chr(0), LV1.SelectedItem.Tag
            If Sock(LV1.SelectedItem.Tag).State <> 7 Then LVD.ListItems.Item(i).SubItems(2) = "Disconnected"
        End If
    Next
    For i = 1 To LVU.ListItems.Count
        If LVU.ListItems.Item(i).SubItems(2) = "Waiting" Then
            LVU.ListItems.Remove i
        ElseIf LVU.ListItems.Item(i).SubItems(2) = "Connected" Then
            Transfer_Close Int(Right(LVU.ListItems.Item(i).Key, Len(LVU.ListItems.Item(i).Key) - 1))
        End If
    Next
End Sub

Private Sub tmrSpeed_Timer()
On Error Resume Next
gSysTray.IconInSysTray
    For i = 1 To Transfer.UBound
        If TState(i) = 2 Then
            Dim X As String
            X = DoConv((LOF(i) - TSpeed(i)) / (tmrSpeed.Interval / 1000))
            LVD.ListItems("x" & i).SubItems(3) = X & "/s"
            'SetProgress LVD.ListItems("x" & i).Index, Int(LOF(i) / TSize(i) * 100)
            TSpeed(i) = LOF(i)
        ElseIf TState(i) = 12 Then
            X = DoConv((NSend(i) - LSend(i)) / (tmrSpeed.Interval / 1000))
            LVU.ListItems("x" & i).SubItems(3) = X & "/s"
            'SetProgress LVD.ListItems("x" & i).Index, Int(LOF(i) / TSize(i) * 100)
            LSend(i) = NSend(i)
        End If
    Next

End Sub

Private Sub tmrStatus_Timer()
    On Error Resume Next
    Dim s As Double
    Dim f As Double
    Dim u As Double
    Dim SV As Double
    For i = 0 To Sock.UBound
        If Sock(i).State = 7 Then
            SV = SV + 1
            'If LVC.ListItems(Network(index)).SubItems(4) = "" Then LVC.ListItems(Network(index)).SubItems(4) = 0
            'Dim check As String
            'For s = 2 To 4
            '    If LVC.ListItems(i).SubItems(s) = "" Then LVC.ListItems(i).SubItems(s) = "0"
            'Next
            s = s + LVC.ListItems(Network(i)).SubItems(4)
            f = f + LVC.ListItems(Network(i)).SubItems(3)
            u = u + LVC.ListItems(Network(i)).SubItems(2)
            S1.Panels.Item(1).Text = "Servers : " & SV
            S1.Panels.Item(2).Text = "Users : " & u
            S1.Panels.Item(3).Text = "Files : " & f
            S1.Panels.Item(4).Text = "GigaBytes : " & s
            'u f s
        End If
    Next
End Sub

Private Sub Transfer_Close(Index As Integer)
    On Error Resume Next
    Dim LOFS As Long
    LOFS = 0
    LOFS = LOF(Index)
    If LOFS > TSize(Index) - 1 And TSize(Index) > 0 Then
        LVD.ListItems("x" & Index).SubItems(2) = "Completed"
        LVD.ListItems("x" & Index).SubItems(3) = "0 B/s"
        LVD.ListItems("x" & Index).Key = ""
        GoTo TC
    End If
    If LVD.ListItems("x" & Index).SubItems(2) = "Downloading" Or InStr(LVD.ListItems("x" & Index).SubItems(2), "/") Then
        SendData "", NumToChr(219), LVD.ListItems("x" & Index).Tag
    End If
    LVD.ListItems("x" & Index).SubItems(2) = "Disconnected"
TC:
    Transfer(Index).Close
    TState(Index) = 0
    TSpeed(Index) = 0
    NSend(Index) = 0
    LSend(Index) = 0
    UFile(Index) = ""
    If LVU.ListItems("x" & Index).SubItems(2) = "Uploading" Or InStr(LVU.ListItems("x" & Index).SubItems(2), "/") Then
        SendData "", NumToChr(221), LVU.ListItems("x" & Index).Tag
    End If
    LVU.ListItems.Remove LVU.ListItems("x" & Index).Index
    LVD.ListItems("x" & Index).SubItems(3) = "0 B/s"
    LVD.ListItems("x" & Index).Key = LVD.ListItems("x" & Index).ToolTipText & LVD.ListItems("x" & Index)
    Close #Index
    If Index = 0 Then
    Transfer(0).Listen
    End If
End Sub

Private Sub Transfer_Connect(Index As Integer)
    'MsgBox "Connected " & Index
    On Error Resume Next
    tmrDL.Enabled = False
    tmrDL.Enabled = True
    TState(Index) = 17
    LVD.ListItems("x" & Index).SubItems(2) = "Connected"
    LVU.ListItems("x" & Index).SubItems(2) = "Connected"
    DoEvents
    ReDim Preserve LSend(Transfer.UBound)
    ReDim Preserve NSend(Transfer.UBound)
    ReDim Preserve UFile(Transfer.UBound)
    NSend(Index) = 0
    LSend(Index) = 0
    UFile(Index) = ""
End Sub

Private Sub Transfer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim X As Integer
    X = OpenSocket("Transfer", frmMain)
    ReDim Preserve LSend(Transfer.UBound)
    ReDim Preserve NSend(Transfer.UBound)
    ReDim Preserve UFile(Transfer.UBound)
    Transfer(X).Accept requestID
    'MsgBox Transfer(x).RemoteHostIP
    Transfer(X).SendData "1" '? this is gay
    ReDim Preserve TState(Transfer.UBound)
    TState(X) = 5 'Set to listen for get/send
    ReDim Preserve TSpeed(Transfer.UBound)
    ReDim Preserve TSize(Transfer.UBound)
    NSend(X) = 0
    LSend(X) = 0
    UFile(X) = ""
    'TSize(x) = 0
    TSpeed(X) = 0
    'Me.Caption = Transfer(x).RemoteHostIP
    
End Sub

Private Sub Transfer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Sata As String
    On Error Resume Next
    Transfer(Index).GetData Sata
Tops:
    Select Case TState(Index)
        Case 9
            Dim FSL As String
            FSL = LOF(Index)
ReCheck1:
            If IsNumeric(Mid(Sata, 1, 1)) Then
                Sata = Right(Sata, Len(Sata) - 1)
            End If
            If IsNumeric(Mid(Sata, 1, 1)) Then GoTo ReCheck1:
            If LOF(Index) = 0 Then FSL = 1
            Put #Index, 1, Sata
            TState(Index) = 2
        Case 6
            Dim SI() As String
            SI = Split(Sata, """")
            If LVD.ListItems(SI(1) & RTrim(SI(0))).SubItems(2) = "Listening" Then
                LVD.ListItems(SI(1) & RTrim(SI(0))).SubItems(2) = "Downloading"
                SendData "", NumToChr(218), LVD.ListItems(SI(1) & RTrim(SI(0))).Tag
                Open DownloadDir & "\" & RemChr(LVD.ListItems(SI(1) & RTrim(SI(0))).SubItems(1)) For Binary As Index
                LVD.ListItems(SI(1) & RTrim(SI(0))).Key = "x" & Index
                TState(Index) = 9
                TSize(Index) = LTrim(SI(2))
                FSL = LOF(Index) - 1
                If FSL < 1 Then FSL = 0
                Transfer(Index).SendData FSL
            End If
        Case 10
            SI = Split(Sata, """")
            Dim isQueue As String
            isQueue = 0 '
            isQueue = LVU.ListItems(SI(1) & Trim(SI(0))).Index
            If isQueue = 0 Then
                Transfer_Close Index
                Exit Sub
            Else
            LVU.ListItems(SI(1) & Trim(SI(0))).Key = "x" & Index
            LVU.ListItems("x" & Index).SubItems(2) = "Uploading"
            SendData "", NumToChr(218), LVD.ListItems(SI(1) & RTrim(SI(0))).Tag
                UFile(Index) = SI(1)
                LSend(Index) = (Trim(SI(2)) - 1)
                NSend(Index) = (Trim(SI(2)) - 1)
                If NSend(Index) <= 0 Then NSend(Index) = 1
                Open UFile(Index) For Binary Access Read As #253 'use #253 for uploads
                        Dim Sdta As String
                        Sdta = Space$(4096)
                        If NSend(Index) + 4096 > LOF(253) - 1 Then
                            If NSend(Index) >= LOF(253) Then
                                Close #253
                                Transfer_Close Index
                                Exit Sub
                            End If
                            Sdta = Space$(LOF(253) - NSend(Index))
                        End If
                        Dim hldSize As Long
                        hldSize = (LOF(253) - 1)
                        Get #253, NSend(Index), Sdta
                        NSend(Index) = NSend(Index) + Len(Sdta)
                        LVU.ListItems("x" & Index).SubItems(2) = DoConv(CStr(NSend(Index))) & "/" & DoConv(LOF(253))
               Close #253
                    TState(Index) = 12   'Set to 12 aka uploading
                    Transfer(Index).SendData hldSize & Sdta
                End If
        Case 15
            If Not IsNumeric(Sata) Then
                Transfer_Close Index
                Exit Sub
            End If
                LVU.ListItems("x" & Index).SubItems(2) = "Uploading"
                SendData "", NumToChr(220), LVD.ListItems(SI(1) & RTrim(SI(0))).Tag
                UFile(Index) = LVU.ListItems("x" & Index).SubItems(1)
                LSend(Index) = (Sata - 1)
                NSend(Index) = (Sata - 1)
                If NSend(Index) <= 0 Then NSend(Index) = 1
                    Open UFile(Index) For Binary Access Read As #253 'use #253 for uploads
                    Sdta = Space$(4096)
                    If NSend(Index) + 4096 > LOF(253) - 1 Then
                        If LOF(253) - NSend(Index) < 1 Then
                            LVU.ListItems("x" & Index).SubItems(2) = "Complete"
                            Transfer_Close Index
                            Close #253
                            Exit Sub
                        End If
                        Sdta = Space$(LOF(253) - NSend(Index))
                    End If
                    Get #253, NSend(Index), Sdta
                    NSend(Index) = NSend(Index) + Len(Sdta)
                    hldSize = (LOF(253) - 1)
                    Close #253
                    TState(Index) = 12   'Set to 12 aka uploading
                    Transfer(Index).SendData hldSize & Sdta
        Case 5
            If Left(Sata, 3) = "GET" Then
                TState(Index) = 10
                Sata = Right(Sata, Len(Sata) - 3)
                If Len(Sata) > 0 Then GoTo Tops:
                Exit Sub
            ElseIf Left(Sata, 4) = "SEND" Then
                TState(Index) = 6
                Sata = Right(Sata, Len(Sata) - 4)
                If Len(Sata) > 0 Then GoTo Tops:
                Exit Sub
            Else
                Transfer_Close Index
            End If
        Case 17
            If Sata = "1" Then
                TState(Index) = 1
                Dim MyCheck As Integer
                    MyCheck = 0
                    MyCheck = LVD.ListItems("x" & Index).Index
                    If MyCheck = 0 Then
                        MyCheck = LVU.ListItems("x" & Index).Index
                        If MyCheck = 0 Then
                            Transfer_Close Index
                            Exit Sub
                        Else
                            TState(Index) = 15 'Wait for Byte Offset
                            'MsgBox Nickname & " """ & LVU.ListItems("x" & Index).SubItems(1) & """ " & FileLen(LVU.ListItems("x" & Index).SubItems(1))
                            Transfer(Index).SendData "SEND"
                            Transfer(Index).SendData Nickname & " """ & LVU.ListItems("x" & Index).SubItems(1) & """ " & FileLen(LVU.ListItems("x" & Index).SubItems(1)) - 1
                        End If
                    Else
                        Transfer(Index).SendData "GET"
                        Open DownloadDir & "\" & RemChr(LVD.ListItems("x" & Index).SubItems(1)) For Binary As #Index
                        If LOF(Index) = 0 Then
                            Transfer(Index).SendData Nickname & " """ & LVD.ListItems("x" & Index).ToolTipText & """ " & CStr(LOF(Index))
                        Else
                            Transfer(Index).SendData Nickname & " """ & LVD.ListItems("x" & Index).ToolTipText & """ " & CStr(LOF(Index) - 1)
                        End If
                    End If
            End If
        Case 1
            If Len(Sata) = 0 Then Exit Sub
            TState(Index) = 2
ReCheck:
            If IsNumeric(Mid(Sata, 1, 1)) Then
                Sata = Right(Sata, Len(Sata) - 1)
            End If
            If IsNumeric(Mid(Sata, 1, 1)) Then GoTo ReCheck:
            FSL = LOF(Index)
            If LOF(Index) = 0 Then FSL = 1
            Put #1, CLng(FSL), Sata
            LVD.ListItems("x" & Index).SubItems(2) = DoConv(LOF(Index)) & "/" & DoConv(CStr(TSize(Index)))
        Case 2
            Put #Index, , Sata
            LVD.ListItems("x" & Index).SubItems(2) = DoConv(LOF(Index)) & "/" & DoConv(CStr(TSize(Index)))
    End Select
End Sub

Private Sub Transfer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Transfer_Close Index
End Sub

Private Sub Transfer_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error Resume Next
    If bytesRemaining = 0 And TState(Index) = 12 Then
        Open UFile(Index) For Binary Access Read As #253 'use #253 for uploads
        Dim Sdta As String
        Sdta = Space$(4096)
        If NSend(Index) + 4096 > LOF(253) - 1 Then
            If NSend(Index) >= LOF(253) Then
                Close #253
                Transfer_Close Index
                DoEvents
                Exit Sub
            End If
            Sdta = Space$(LOF(253) - NSend(Index))
        End If
        Get #253, NSend(Index), Sdta
        NSend(Index) = NSend(Index) + Len(Sdta)
        LVU.ListItems("x" & Index).SubItems(2) = DoConv(CStr(NSend(Index))) & "/" & DoConv(LOF(253))
        Close #253
        Transfer(Index).SendData Sdta
    End If
End Sub
