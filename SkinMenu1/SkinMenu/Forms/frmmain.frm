VERSION 5.00
Object = "*\A..\Projects\SkinMenu.vbp"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demonistration Plz Vote"
   ClientHeight    =   5415
   ClientLeft      =   4125
   ClientTop       =   3075
   ClientWidth     =   8535
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
   ScaleHeight     =   5415
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdapply 
      Caption         =   "Apply Settings"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmmain.frx":0000
      Left            =   5400
      List            =   "frmmain.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmmain.frx":0079
      Left            =   5400
      List            =   "frmmain.frx":0089
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin SkinMenu.sMenu sMenu1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   661
      Style           =   1
      ItemCount       =   68
      BeginProperty ItemsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuCaption1    =   "File"
      MenuName1       =   "mnuFile"
      MenuDescription1=   "click here to open Commands Related to files"
      MenuIdent2      =   1
      MenuCaption2    =   "&New"
      MenuName2       =   "mnuFileNew"
      MenuPicture2    =   "frmmain.frx":00CB
      MenuDescription2=   "Create New Document"
      MenuIdent3      =   1
      MenuCaption3    =   "&Open"
      MenuName3       =   "mnuFileOpen"
      MenuPicture3    =   "frmmain.frx":041D
      MenuDescription3=   "Open Preivous File"
      MenuIdent4      =   1
      MenuCaption4    =   "&Close"
      MenuName4       =   "mnuFileClose"
      MenuMaskColor4  =   12632256
      MenuPicture4    =   "frmmain.frx":076F
      MenuDescription4=   "Close Currently Open File"
      MenuIdent5      =   1
      MenuCaption5    =   "Save Options"
      MenuName5       =   "mnuFileSep0"
      MenuDescription5=   "Save Related Options"
      MenuIdent6      =   1
      MenuCaption6    =   "&Save"
      MenuName6       =   "mnuFileSave"
      MenuPicture6    =   "frmmain.frx":13C1
      MenuDescription6=   "Save This Document to HardDrive"
      MenuIdent7      =   1
      MenuCaption7    =   "Save &All"
      MenuName7       =   "mnuFileSaveAll"
      MenuPicture7    =   "frmmain.frx":1713
      MenuDescription7=   "Save All Open Document"
      MenuIdent8      =   1
      MenuCaption8    =   "Save As"
      MenuName8       =   "mnuFileSaveAs"
      MenuDescription8=   "Save This File As"
      MenuIdent9      =   1
      MenuCaption9    =   "Print Layout Options"
      MenuName9       =   "mnuFileSep3"
      MenuPicture9    =   "frmmain.frx":1A65
      MenuDescription9=   "Print Realted Options and Settings"
      MenuIdent10     =   1
      MenuCaption10   =   "Page Set&up..."
      MenuName10      =   "mnuFilePgSetup"
      MenuPicture10   =   "frmmain.frx":1DB7
      MenuIdent11     =   1
      MenuCaption11   =   "Print Pre&view..."
      MenuName11      =   "mnuFilePrntPrvw"
      MenuPicture11   =   "frmmain.frx":2109
      MenuDescription11=   "Show Print Preview"
      MenuIdent12     =   1
      MenuCaption12   =   "&Print..."
      MenuName12      =   "mnuFilePrint"
      MenuPicture12   =   "frmmain.frx":242B
      MenuDescription12=   "Print Current Document"
      MenuIdent13     =   1
      MenuCaption13   =   "-"
      MenuName13      =   "mnuFileSep2"
      MenuIdent14     =   1
      MenuCaption14   =   "E&xit"
      MenuName14      =   "mnuFileExit"
      MenuDescription14=   "Exit Here"
      MenuCaption15   =   "Edit"
      MenuName15      =   "mnuEdit"
      MenuIdent16     =   1
      MenuCaption16   =   "&Undo"
      MenuName16      =   "mnuEditUndo"
      MenuPicture16   =   "frmmain.frx":277D
      MenuIdent17     =   1
      MenuCaption17   =   "&Redo"
      MenuName17      =   "mnuEditRedo"
      MenuPicture17   =   "frmmain.frx":2ACF
      MenuIdent18     =   1
      MenuCaption18   =   "Clipboard Options"
      MenuName18      =   "mnuEditSep0"
      MenuPicture18   =   "frmmain.frx":2E21
      MenuIdent19     =   1
      MenuCaption19   =   "Cu&t"
      MenuName19      =   "mnuEditCut"
      MenuPicture19   =   "frmmain.frx":3173
      MenuIdent20     =   1
      MenuCaption20   =   "&Copy"
      MenuName20      =   "mnuEditCopy"
      MenuPicture20   =   "frmmain.frx":34C5
      MenuIdent21     =   1
      MenuCaption21   =   "&Paste"
      MenuName21      =   "mnuEditPaste"
      MenuPicture21   =   "frmmain.frx":3817
      MenuIdent22     =   1
      MenuCaption22   =   "Search Options"
      MenuName22      =   "mnuEditSep1"
      MenuPicture22   =   "frmmain.frx":3B69
      MenuIdent23     =   1
      MenuCaption23   =   "&Find"
      MenuName23      =   "mnuEditFind"
      MenuPicture23   =   "frmmain.frx":47BB
      MenuIdent24     =   1
      MenuCaption24   =   "Select &All"
      MenuName24      =   "mnuEditSelAll"
      MenuIdent25     =   1
      MenuCaption25   =   "-"
      MenuName25      =   "mnuEditSep3"
      MenuIdent26     =   1
      MenuCaption26   =   "Proper&ties"
      MenuName26      =   "mnuEditProp"
      MenuPicture26   =   "frmmain.frx":540D
      MenuCaption27   =   "Format"
      MenuName27      =   "mnuFormat"
      MenuIdent28     =   1
      MenuCaption28   =   "Format Options..."
      MenuName28      =   "mnuFormatSideBar"
      MenuIdent29     =   1
      MenuCaption29   =   "&Font"
      MenuName29      =   "mnuFmtFont"
      MenuIdent30     =   1
      MenuCaption30   =   "&Spell Check"
      MenuName30      =   "mnuFmtSplChk"
      MenuPicture30   =   "frmmain.frx":605F
      MenuIdent31     =   1
      MenuCaption31   =   "Text Format Options"
      MenuName31      =   "mnuFmtSep0"
      MenuIdent32     =   1
      MenuCaption32   =   "Bold"
      MenuName32      =   "mnuFmtBold"
      MenuPicture32   =   "frmmain.frx":6CB1
      MenuIdent33     =   1
      MenuCaption33   =   "Italic"
      MenuName33      =   "mnuFmtItalic"
      MenuPicture33   =   "frmmain.frx":7903
      MenuIdent34     =   1
      MenuCaption34   =   "Underline"
      MenuName34      =   "mnuFmtUndrln"
      MenuPicture34   =   "frmmain.frx":8555
      MenuIdent35     =   1
      MenuCaption35   =   "-"
      MenuName35      =   "mnuFileSep1"
      MenuIdent36     =   1
      MenuCaption36   =   "Sort"
      MenuName36      =   "mnuFmtSort"
      MenuPicture36   =   "frmmain.frx":91A7
      MenuIdent37     =   2
      MenuCaption37   =   "Ascending"
      MenuName37      =   "mnuFmtSortAsc"
      MenuIdent38     =   2
      MenuCaption38   =   "Descending"
      MenuName38      =   "mnuFmtSortDesc"
      MenuIdent39     =   1
      MenuCaption39   =   "-"
      MenuName39      =   "mnuFmtSep2"
      MenuIdent40     =   1
      MenuCaption40   =   "Align"
      MenuName40      =   "mnuFrmtAlgn"
      MenuIdent41     =   2
      MenuCaption41   =   "Align Options"
      MenuName41      =   "mnuFrmtAlgnSideBar"
      MenuPicture41   =   "frmmain.frx":9DF9
      MenuIdent42     =   2
      MenuCaption42   =   "&Left"
      MenuName42      =   "mnuFrmtAlgnOptn"
      MenuPicture42   =   "frmmain.frx":AA4B
      MenuIdent43     =   2
      MenuCaption43   =   "&Right"
      MenuName43      =   "mnuFrmtAlgnOptn"
      MenuPicture43   =   "frmmain.frx":B69D
      MenuIdent44     =   2
      MenuCaption44   =   "&Center"
      MenuName44      =   "mnuFrmtAlgnOptn"
      MenuPicture44   =   "frmmain.frx":C2EF
      MenuIdent45     =   2
      MenuCaption45   =   "&Justify"
      MenuName45      =   "mnuFrmtAlgnOptn"
      MenuPicture45   =   "frmmain.frx":CF41
      MenuIdent46     =   1
      MenuCaption46   =   "-"
      MenuName46      =   "mnuFmtSep3"
      MenuIdent47     =   1
      MenuCaption47   =   "Paint"
      MenuName47      =   "mnuFmtPaint"
      MenuPicture47   =   "frmmain.frx":DB93
      MenuCaption48   =   "Nested Menu"
      MenuName48      =   "mnuNstMnu"
      MenuChecked49   =   -1  'True
      MenuIdent49     =   1
      MenuCaption49   =   "Sub Menu 1 Checked"
      MenuName49      =   "mnuNstSubMnu1"
      MenuIdent50     =   1
      MenuCaption50   =   "Sub Menu 2"
      MenuName50      =   "mnuNstSubMnu2"
      MenuIdent51     =   2
      MenuCaption51   =   "Sub Menu 3"
      MenuName51      =   "mnuNstSubMnu3"
      MenuIdent52     =   2
      MenuCaption52   =   "Sub Menu 4"
      MenuName52      =   "mnuNstSubMnu4"
      MenuIdent53     =   3
      MenuCaption53   =   "Sub Menu 5"
      MenuName53      =   "mnuNstSubMnu5"
      MenuIdent54     =   3
      MenuCaption54   =   "Sub Menu 6"
      MenuName54      =   "mnuNstSubMnu6"
      MenuIdent55     =   4
      MenuCaption55   =   "Sub Menu 7"
      MenuName55      =   "mnuNstSubMnu7"
      MenuIdent56     =   4
      MenuCaption56   =   "Sub Menu 8"
      MenuName56      =   "mnuNstSubMnu8"
      MenuIdent57     =   5
      MenuCaption57   =   "Last"
      MenuName57      =   "mnuNstSubMnu8SideBar"
      MenuIdent58     =   5
      MenuCaption58   =   "Sub Menu 9"
      MenuName58      =   "mnuNstSubMnu9"
      MenuIdent59     =   5
      MenuCaption59   =   "Sub Menu 10"
      MenuName59      =   "mnuNstSubMnu10"
      MenuCaption60   =   "Window"
      MenuName60      =   "mnuWnd"
      MenuIdent61     =   1
      MenuCaption61   =   "Window"
      MenuName61      =   "mnuWndSidebar"
      MenuIdent62     =   1
      MenuCaption62   =   "New Window"
      MenuName62      =   "mnuWndwNew"
      MenuIdent63     =   1
      MenuCaption63   =   "Arrange Windows"
      MenuName63      =   "mnuWndSep0"
      MenuPicture63   =   "frmmain.frx":E7E5
      MenuIdent64     =   1
      MenuCaption64   =   "Cascade"
      MenuName64      =   "mnuWndPos"
      MenuPicture64   =   "frmmain.frx":F437
      MenuIdent65     =   1
      MenuCaption65   =   "Tile Horizontal"
      MenuName65      =   "mnuWndPos"
      MenuPicture65   =   "frmmain.frx":10089
      MenuIdent66     =   1
      MenuCaption66   =   "Tile Vertical"
      MenuName66      =   "mnuWndPos"
      MenuPicture66   =   "frmmain.frx":10CDB
      MenuIdent67     =   1
      MenuCaption67   =   "Arrange Icons"
      MenuName67      =   "mnuWndPos"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   120
      Picture         =   "frmmain.frx":1192D
      ScaleHeight     =   4425
      ScaleWidth      =   2985
      TabIndex        =   8
      Top             =   720
      Width           =   3015
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "XpStyle Skinize menu"
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
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   $"frmmain.frx":121E3
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Highlight Style:"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmmain.frx":122D4
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3600
      TabIndex        =   6
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Appaerance:"
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   780
      Width           =   1515
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdapply_Click()
sMenu1.Style = Combo1.ListIndex
sMenu1.HighLightStyle = Combo2.ListIndex
sMenu1.Refresh



End Sub

Private Sub Form_Load()
Combo1.ListIndex = sMenu1.Style
Combo2.ListIndex = sMenu1.HighLightStyle
End Sub

Private Sub sMenu1_ItemClick(Key As String)
Select Case Key
Case "mnuFileNew"
MsgBox "clicked on new", vbinfo




End Select

End Sub

Private Sub sMenu1_ItemDescription(Description As String)
Label2.Caption = Description

End Sub
