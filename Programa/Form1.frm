VERSION 5.00
Object = "{553E8CEC-F455-4A8A-B7EE-4492089A2AB5}#20.0#0"; "TS_CTRL.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8340
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin TS_CTRL.txtCampo txtCampoExtensao 
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483625
      BorderStyle     =   0
   End
   Begin TS_CTRL.txtCampo txtCampoNome 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483625
      BorderStyle     =   0
   End
   Begin TS_CTRL.txtCampo txtCampoCaminho 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483625
      BorderStyle     =   0
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      DragMode        =   1  'Automatic
      Height          =   4335
      HelpContextID   =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "coluna 2"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "coluna 3"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "coluna 4"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "coluna 5"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "coluna 6"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "coluna 7"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "coluna 1"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=260"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=260"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=260"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2752"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=260"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2752"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=260"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2752"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=260"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2752"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=260"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2752"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=260"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2752"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=260"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      BorderStyle     =   0
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      EditDropDown    =   0   'False
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "Arquivos no Banco"
      ExposeCellMode  =   2
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MousePointer    =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   1
      CellTipsWidth   =   0
      MultiSelect     =   2
      OLEDragMode     =   1
      OLEDropMode     =   2
      DataView        =   1
      AnimateWindow   =   3
      AnimateWindowDirection=   3
      AnimateWindowClose=   1
      DeadAreaBackColor=   65535
      ScrollTips      =   -1  'True
      RowDividerColor =   65535
      RowSubDividerColor=   8454143
      DirectionAfterEnter=   2
      MaxRows         =   250000
      ChildGrid       =   "0"
      ChildGrid.vt    =   8
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      ExpandColor     =   8454016
      CollapseColor   =   49344
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.fgcolor=&H8000000F&"
      _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=42,.bgcolor=&H80000010&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000F&,.bold=0,.fontsize=1200,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000011&"
      _StyleDefs(15)  =   ":id=3,.fgcolor=&H80000011&,.bold=0,.fontsize=1200,.italic=0,.underline=0"
      _StyleDefs(16)  =   ":id=3,.strikethrough=0,.charset=0"
      _StyleDefs(17)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H80000010&,.fgcolor=&H80000012&"
      _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000011&"
      _StyleDefs(20)  =   ":id=6,.fgcolor=&H80000012&"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000006&,.fgcolor=&H80000005&"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000016&"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.appearance=0"
      _StyleDefs(25)  =   ":id=10,.borderColor=&H80000000&"
      _StyleDefs(26)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(27)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(28)  =   "Splits(0).Style:id=83,.parent=1"
      _StyleDefs(29)  =   "Splits(0).CaptionStyle:id=92,.parent=4"
      _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=84,.parent=2"
      _StyleDefs(31)  =   "Splits(0).FooterStyle:id=85,.parent=3"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=86,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=88,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=87,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=89,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=90,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=91,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=93,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=94,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=98,.parent=83"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=84,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=85"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=87"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=102,.parent=83"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=84,.alignment=0"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=85"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=87"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=106,.parent=83"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=84,.alignment=0"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=85"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=87"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=110,.parent=83"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=84,.alignment=0"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=85"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=87"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=114,.parent=83"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=84,.alignment=0"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=85"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=87"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=118,.parent=83"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=115,.parent=84,.alignment=0"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=116,.parent=85"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=117,.parent=87"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=134,.parent=83"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=131,.parent=84,.alignment=0"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=132,.parent=85"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=133,.parent=87"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=138,.parent=83"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=135,.parent=84,.alignment=0"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=136,.parent=85"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=137,.parent=87"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=142,.parent=83"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=139,.parent=84,.alignment=0"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=140,.parent=85"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=141,.parent=87"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin TS_CTRL.xpcmdbutton xpcmdbutton2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Escolher Arquivo no explorador"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TS_CTRL.xpcmdbutton xpcmdbutton1 
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   7680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Importar para o Banco"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TS_CTRL.xpTitle xpTitle1 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1614
      Caption         =   "Importação de Arquivos"
      RtName          =   "BLA BLA BLA"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Extenção do arquivo (Ex: .txt)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   9
      Top             =   2400
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cole o nome do arquivo com a sua extenção aqui"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Após escolher caminho do  arquivo copie e cole aqui"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As ADODB.Connection
Dim rc  As ADODB.Recordset

Dim banco As String
Dim ACCTID As String
Dim ACCTTYPE As String
Dim NUMERO_STMTTRN As String
Dim OPERACAO As String
Dim dia As String
Dim valor As Long
Dim resumo As String
Dim CHECKNUM As String
Dim MEMO As String
Dim BRANCHID As String
Dim REFNUM As String
Dim controlaFinal As Integer
Dim parouAqui As Integer
Dim numeroLinha As Integer
Dim achou As Integer
Dim totalRegistros As Integer
Dim final As Integer

Dim query As String
'Tags Existentes

Const BANKID1 As String = "<BANKID>"
Const BANKID2 As String = "</BANKID>"
Const ACCTID1 As String = "<ACCTID>"
Const ACCTID2 As String = "</ACCTID>"
Const ACCTTYPE1 As String = "<ACCTTYPE>"
Const ACCTTYPE2 As String = "</ACCTTYPE>"
Const TRNTYPE1 As String = "<TRNTYPE>"
Const TRNTYPE2 As String = "</TRNTYPE>"
Const DTPOSTED1 As String = "<DTPOSTED>"
Const DTPOSTED2 As String = "</DTPOSTED>"
Const TRNAMT1 As String = "<TRNAMT>"
Const TRNAMT2 As String = "</TRNAMT>"
Const FITID1 As String = "<FITID>"
Const FITID2 As String = "</FITID>"
Const CHECKNUM1 As String = "<CHECKNUM>"
Const CHECKNUM2 As String = "</CHECKNUM>"
Const MEMO1 As String = "<MEMO>"
Const MEMO2 As String = "</MEMO>"
Const REFNUM1 As String = "<REFNUM>"
Const REFNUM2 As String = "</REFNUM>"
Const BRANCHID1 As String = "<BRANCHID>"
Const BRANCHID2 As String = "</BRANCHID>"

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Function conexaoBanco() As Boolean
        Set conexao = New ADODB.Connection
        Set rc = New ADODB.Recordset
        With conexao
            
            .CursorLocation = adUseClient
            .ConnectionString = "Driver={SQL Server};Server=.;Uid=sa;Pwd=254685ro;Database=BancoOFX"
            
            .Open
            
            If .State = adStateOpen Then
                conexaoBanco = True
            Else
                conexaoBanco = False
            End If
        End With
End Function

Private Sub Analise(referencia As Integer)
'<BANKID>   - banco
'<BRANCHID>
'<ACCTID>
'<ACCTTYPE>
'<TRNTYPE>  - OPERACAO
'<DTPOSTED> - DIA
'<TRNAMT>   - VALOR
'<FITID>    - RESUMO
'<CHECKNUM>
'<REFNUM>
'<MEMO>

'11 tags

numeroLinha = 0
parouAqui = 0
achou = 1
totalRegistros = 0
final = 0

DescobrindoOfinal referencia

For Var = 1 To totalRegistros

    banco = LimpaArquivo(BANKID1, BANKID2, referencia)
    BRANCHID = LimpaArquivo(BRANCHID1, BRANCHID2, referencia)
    ACCTID = LimpaArquivo(ACCTID1, ACCTID2, referencia)
    ACCTTYPE = LimpaArquivo(ACCTTYPE1, ACCTTYPE2, referencia)
    OPERACAO = LimpaArquivo(TRNTYPE1, TRNTYPE2, referencia)
    dia = LimpaArquivo(DTPOSTED1, DTPOSTED2, referencia)
    dia = Mid(dia, 7, 2) & "/" & Mid(dia, 5, 2) & "/" & Mid(dia, 1, 4)
    valor = LimpaArquivo(TRNAMT1, TRNAMT2, referencia)
    resumo = LimpaArquivo(FITID1, FITID2, referencia)
    CHECKNUM = LimpaArquivo(CHECKNUM1, CHECKNUM2, referencia)
    REFNUM = LimpaArquivo(REFNUM1, REFNUM2, referencia)
    MEMO = LimpaArquivo(MEMO1, MEMO2, referencia)

    query = "insert into Movimentacao (Banco, ACCTID, ACCTTYPE, NUMERO_STMTTRN, OPERACAO, DIA, VALOR, RESUMO, CHECKNUM, MEMO, BRANCHID, REFNUM) values ('" & banco & "','" & ACCTID & "','" & ACCTTYPE & "'," & 1 & ",'" & OPERACAO & "', '" & dia & "'," & valor & ",'" & resumo & "', '" & CHECKNUM & "', '" & MEMO & "', '" & BRANCHID & "', '" & REFNUM & "')"
    
    conexaoBanco
    
    conexao.Execute (query)
Next
    query = "select * from Movimentacao"
    Set rc = conexao.Execute(query)
    
    Set TDBGrid1.DataSource = rc
End Sub

Private Function AbrindoArquivo(referencia As Integer) As Integer
        Dim numero As Integer
        
        If referencia = 1 Then
            numero = FreeFile
            Open txtCampoCaminho.Text & "\" & txtCampoNome.Text & "." & txtCampoExtensao For Input As #numero
            AbrindoArquivo = numero
        Else
            If referencia = 2 Then
                numero = FreeFile
                Open txtCampoCaminho.Text & "\" & "novoArquivoVB6" & "." & txtCampoExtensao.Text For Input As #numero
                AbrindoArquivo = numero
            End If
        End If
End Function


Private Sub DescobrindoOfinal(referencia As Integer)
    
    Dim numero As Integer
    Dim linha As String
    
    numero = AbrindoArquivo(referencia)
    
    Do While Not EOF(numero)
        Line Input #numero, linha
        
        If InStr(1, linha, "</STMTTRN>") > 0 Then
            totalRegistros = totalRegistros + 1
        End If
    Loop
    Close #numero

End Sub

Private Function LimpaArquivo(tag1 As String, tag2 As String, referencia As Integer) As String
        Dim linha As String
        Dim numero As Integer
           
        numero = AbrindoArquivo(referencia)
        
        Do While Not EOF(numero)
            Line Input #numero, linha
                        
            numeroLinha = numeroLinha + 1
            
            If numeroLinha >= parouAqui Then
            
                linha = Replace(linha, vbTab, "")
                If InStr(1, linha, tag1) > 0 Then
                
                    linha = Replace(linha, tag1, "")
                    linha = Replace(linha, tag2, "")
                    
                    achou = 2
                    parouAqui = numeroLinha
                    numeroLinha = 0
                    Exit Do
                End If
            End If
            totalLinhas = totalLinhas + 1
        Loop
        
        If achou = 1 Then
            If EOF(numero) Then
                numeroLinha = 0
            End If
            Close #numero
            LimpaArquivo = ""
        Else
            If achou = 2 Then
                achou = 1
                Close #numero
                LimpaArquivo = linha
            End If
        End If
          
End Function

Private Sub xpcmdbutton1_Click()
    Dim arquivo As Integer
    Dim linha As String
    
    arquivo = FreeFile
    
    Open txtCampoCaminho.Text & "\" & txtCampoNome.Text & "." & txtCampoExtensao.Text For Input As #arquivo
    
    Line Input #arquivo, linha
    
    If InStr(1, linha, vbLf) > 0 Then
        
        Dim novoArquivo As Integer
        
        novoArquivo = FreeFile
        
        Open txtCampoCaminho.Text & "\" & "novoArquivoVB6" & "." & txtCampoExtensao.Text For Output As #novoArquivo
        Print #novoArquivo, Replace(linha, vbLf, Chr(10) + Chr(13))
       
        Close arquivo
        Close novoArquivo
             
        Analise 2
    Else
        Close #arquivo
        Analise 1
    End If
End Sub

Private Sub xpcmdbutton2_Click()
   Shell "C:\Windows\explorer.exe", vbNormalFocus
End Sub

