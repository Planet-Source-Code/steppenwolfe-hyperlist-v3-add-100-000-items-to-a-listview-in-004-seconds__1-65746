VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HyperList 3.0 - Test Harness"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   7320
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBench 
      Caption         =   "HyperList Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   5985
      TabIndex        =   28
      Top             =   6615
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "Listview Bench: Add 100,000 Items"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   5985
      TabIndex        =   27
      Top             =   6615
      Width           =   3165
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "HyperList Bench: Add 100,000 Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   270
      TabIndex        =   24
      Top             =   6615
      Width           =   3570
   End
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
      Height          =   6405
      Left            =   5940
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CommandButton cmdControls 
         Caption         =   "Remove Duplicates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   225
         TabIndex        =   33
         Top             =   5310
         Width           =   1725
      End
      Begin VB.CommandButton cmdControls 
         Caption         =   "Sort Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   3420
         TabIndex        =   32
         Top             =   4725
         Width           =   1410
      End
      Begin VB.CommandButton cmdControls 
         Caption         =   "Remove Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1845
         TabIndex        =   31
         Top             =   4725
         Width           =   1410
      End
      Begin VB.CommandButton cmdControls 
         Caption         =   "Add Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   225
         TabIndex        =   30
         Top             =   4725
         Width           =   1410
      End
      Begin VB.ComboBox cbStyles 
         Height          =   330
         Left            =   225
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   4095
         Width           =   1680
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Custom Header Color"
         Height          =   210
         Index           =   17
         Left            =   2790
         TabIndex        =   20
         Top             =   3510
         Width           =   1995
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Sub Item Images"
         Height          =   210
         Index           =   16
         Left            =   2790
         TabIndex        =   19
         Top             =   3135
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Flat Scroll Bar"
         Height          =   210
         Index           =   15
         Left            =   2790
         TabIndex        =   18
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Multi Select"
         Height          =   210
         Index           =   14
         Left            =   2790
         TabIndex        =   17
         Top             =   2385
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Header Hide"
         Height          =   210
         Index           =   13
         Left            =   2790
         TabIndex        =   16
         Top             =   2025
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Header Flat"
         Height          =   210
         Index           =   12
         Left            =   2790
         TabIndex        =   15
         Top             =   1650
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Header Fixed Width"
         Height          =   210
         Index           =   11
         Left            =   2790
         TabIndex        =   14
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Header Drag Drop"
         Height          =   210
         Index           =   10
         Left            =   2790
         TabIndex        =   13
         Top             =   930
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Grid Lines"
         Height          =   210
         Index           =   9
         Left            =   2790
         TabIndex        =   12
         Top             =   570
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Full Row Select"
         Height          =   210
         Index           =   8
         Left            =   270
         TabIndex        =   11
         Top             =   3444
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Fore Color"
         Height          =   210
         Index           =   7
         Left            =   270
         TabIndex        =   10
         Top             =   3081
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Font"
         Height          =   210
         Index           =   6
         Left            =   270
         TabIndex        =   9
         Top             =   2718
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Column Text"
         Height          =   210
         Index           =   5
         Left            =   270
         TabIndex        =   8
         Top             =   2355
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Column Icon"
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   7
         Top             =   1992
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Column Align"
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1629
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Check Boxes"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1266
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Border Style"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   903
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Back Color"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblCtrlResults 
         AutoSize        =   -1  'True
         Caption         =   "Result:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   34
         Top             =   5490
         Width           =   570
      End
      Begin VB.Label lblInfo 
         Caption         =   "View:"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   3870
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lvwTest 
      Height          =   5910
      Left            =   5985
      TabIndex        =   1
      Top             =   315
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   10425
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   135
      ScaleHeight     =   5910
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   315
      Width           =   5595
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   11340
      Top             =   2655
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
            Picture         =   "frmTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":157E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":16D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1832
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":198C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   11430
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":20CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2228
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2382
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":26B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   210
      Index           =   1
      Left            =   6030
      TabIndex        =   26
      Top             =   6255
      Width           =   375
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   210
      Index           =   0
      Left            =   270
      TabIndex        =   25
      Top             =   6255
      Width           =   375
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Standard ListView"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   5940
      TabIndex        =   23
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "HyperList 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   90
      Width           =   1065
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'~ HyperList 3.0 Test Harness by John Underhill (Steppenwolfe)

Private m_bReportMode               As Boolean
Private m_lPasses                   As Long
Private m_lPtr()                    As Long
Private m_lPointer                  As Long
Private m_aData2()                  As String
Private m_tSubData()                As HLISubItm
Private WithEvents cHyperList       As clsHyperList
Attribute cHyperList.VB_VarHelpID = -1
Private cTiming                     As clsTiming
Private m_HLIStc()                  As HLIStc


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Property Let p_Pointer(PropVal As Long)
    m_lPointer = PropVal
End Property

Private Property Get p_Pointer() As Long
    p_Pointer = m_lPointer
End Property

'> Hyperlist Events
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cHyperList_eHColumnClick(Column As Long)
    Debug.Print "Column: " & Column & " clicked."
End Sub

Private Sub cHyperList_eHErrCond(sRtn As String, lErr As Long)
    Debug.Print "Error " & lErr & " in routine: " & sRtn
End Sub

Private Sub cHyperList_eHItemCheck(lItem As Long)
    Debug.Print "Item: " & lItem & " checked."
End Sub

Private Sub cHyperList_eHItemClick(lItem As Long)
    Debug.Print "Item: " & lItem & " clicked."
End Sub

Private Sub cHyperList_eHIndirect(ByVal iItem As Long, _
                                  ByVal iSubItem As Long, _
                                  ByVal fMask As Long, _
                                  sText As String, _
                                  hImage As Long)

'/* Indirect callback method:
'/* if using an external database, this
'/* is where reecords would be passed
'/* by index into the list

    If m_bReportMode Then
        Select Case iSubItem
        Case 0
            sText = "Item " & Format$(iItem, "#,###,##0")
        Case 1
            sText = "Subitem 1"
        Case 2
            sText = "Subitem 2"
        End Select
    Else
        sText = "Item " & Format$(iItem, "#,###,##0")
    End If
    
End Sub

Private Sub PutStandard()

Dim lC      As Long
Dim cItem   As ListItem
        
    With m_HLIStc(0)
        For lC = LBound(.Item) To UBound(.Item)
            Set cItem = lvwTest.ListItems.Add(Text:=.Item(lC))
            cItem.SubItems(1) = .SubItem1(lC)
            cItem.SubItems(2) = .SubItem2(lC)
        Next lC
    End With
    
End Sub

Private Sub cmdBench_Click(Index As Integer)

    cTiming.Reset
    
    Select Case Index
    '/* hyperlist
    Case 0
        PutArray
        lblResults(0).Caption = "100,000 items added to HyperList in: " & _
            Format$(cTiming.Elapsed / 1000, "0.0000") & "s"
        cmdBench(1).Enabled = True
    '/* standard listview
    Case 1
        PutStandard
        lblResults(1).Caption = "100,000 items added to Standard List in: " & _
            Format$(cTiming.Elapsed / 1000, "0.0000") & "s"
        cmdBench(0).Visible = False
        cmdBench(1).Visible = False
        cmdBench(2).Visible = True
    '/* properties
    Case 2
        lvwTest.ListItems.Clear
        lvwTest.Visible = False
        frmControls.Visible = True
        cmdBench(2).Visible = False
    End Select
    
End Sub

Private Sub cmdControls_Click(Index As Integer)

    With cHyperList
        Select Case Index
        '/* add item
        Case 0
            Dim lX As Long
            Dim aSub() As String
            ReDim aSub(1)
            lX = UBound(m_HLIStc(0).Item)
            aSub(0) = "SubItem 1 " & lX
            aSub(1) = "SubItem 2 " & lX
            .ItemAdd lX, "Test Item: " & lX, 1, aSub()
            .ItemEnsureVisible lX
        '/* remove item
        Case 1
            .ItemRemove 1
        '/* sort
        Case 2
            .ItemsSort 0, False
        '/* remove duplicates
        Case 3
            .RemoveDuplicates
        End Select
    End With

End Sub

Public Sub Form_Load()

Dim lX  As Long
Dim lY  As Long
Dim lC  As Long

    With lvwTest
        .View = lvwReport
        .Checkboxes = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Header 1", .Width / 3
        .ColumnHeaders.Add 2, , "Header 2", .Width / 3
        .ColumnHeaders.Add 3, , "Header 3", (.Width / 3) - 100
    End With
    
    '/* instance list and timer
    Set cHyperList = New clsHyperList
    Set cTiming = New clsTiming
    '/* create a portable structure
    ReDim m_HLIStc(0)
    lX = picContainer.ScaleWidth / Screen.TwipsPerPixelX
    lY = picContainer.ScaleHeight / Screen.TwipsPerPixelY
    '/* apply list settings
    With cHyperList
        '/* create a default list
        .CreateList eReport, lX, lY, picContainer.hWnd, App.hInstance
        '/* large icons
        .InitImlLarge
        For lC = 1 To 9
            .ImlLargeAddIcon iml32.ListImages.Item(lC).Picture
        Next lC
        '/* header icons
        .InitImlHeader
        .ImlHeaderAddIcon iml16.ListImages.Item(1).Picture
        .ImlHeaderAddIcon iml16.ListImages.Item(2).Picture
        '/* small images
        .InitImlSmall
        For lC = 3 To 7
            .ImlSmallAddIcon iml16.ListImages.Item(lC).Picture
        Next lC
        '/* enable checkboxes
        .p_CheckBoxes = True
        '/* set viewmode
        .p_ViewMode = eReport
        '/* add columns
        .ColumnAdd 0, "Header 1", lX / 3, [cLeft]
        .ColumnAdd 1, "Header 2", lX / 3, [cLeft]
        .ColumnAdd 2, "Header 3", lX / 3, [cLeft]
        .p_ItemsSorted = True
        '.p_ViewMode = eList
        '.p_IndirectMode = True
        '.SetItemCount 100
        '.ListRefresh
    End With
    
    '/* dimension the list
    ItemsArray 100000
    
    '/* list styles
    With cbStyles
        .Text = "Styles"
        .AddItem "0 - Icon"
        .AddItem "1 - Report"
        .AddItem "2 - SmallIcon"
        .AddItem "3 - List"
        .ListIndex = 1
    End With
    
End Sub

Private Sub cbStyles_Click()

    With cHyperList
        m_bReportMode = (cbStyles.ListIndex = 1)
        .p_ViewMode = cbStyles.ListIndex
        .ListRefresh
    End With
    
End Sub

Private Sub chkOptions_Click(Index As Integer)

Dim bVal    As Boolean

    bVal = chkOptions(Index).Value = 1
    With cHyperList
        Select Case Index
        '/* backcolor
        Case 0
            .p_BackColor = IIf(bVal, &HC0FFC0, vbWhite)
        '/* borderstyle
        Case 1
            .p_BorderStyle = IIf(bVal, bLine, bThin)
        '/* checkboxes
        Case 2
            .p_CheckBoxes = bVal
        '/* alignment
        Case 3
            .p_ColumnAlign(1) = IIf(bVal, ccEnter, cLeft)
        '/* column icon
        Case 4
            .p_ColumnIcon(1) = IIf(bVal, 1, 3)
        '/* column text
        Case 5
            .p_ColumnText(1) = IIf(bVal, "Column 1", "New Text")
        '/* font
        Case 6
            Dim sF As New StdFont
            sF.Charset = 3
            sF.Name = IIf(bVal, "Verdana", "Arial")
            Set .p_Font = sF
        '/* forecolor
        Case 7
            .p_ForeColor = IIf(bVal, &H8000&, vbBlack)
        '/* full row select
        Case 8
            .p_FullRowSelect = bVal
        '/* grid lines
        Case 9
            .p_GridLines = bVal
        '/* header grag drop
        Case 10
            .p_HeaderDragDrop = bVal
        '/* header fixed width
        Case 11
            .p_HeaderFixedWidth = bVal
        '/* header flat
        Case 12
            .p_HeaderFlat = bVal
        '/* hide headers
        Case 13
            .p_HeaderHide = bVal
        '/* multi select
        Case 14
            .p_FullRowSelect = bVal
            .p_MultiSelect = bVal
        '/* flat scroll bar
        Case 15
            .p_ScrollBarFlat = bVal
        '/* subitem images
        Case 16
            .p_SubItemImages = bVal
        '/* header color
        Case 17
            .p_ColumnIcon(1) = -1
            .p_ColumnIcon(2) = -1
            .p_ColumnIcon(3) = -1
            .p_HeaderCustom = bVal
            .p_HeaderColor = &HC9CEAC
            .p_HeaderForeColor = vbWhite
        End Select
        .ListRefresh
    End With

End Sub


'> Array maintenance
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ItemsArray(ByVal lCount As Long)
'/* build an array using copymemory
'/* This represents the data structure
'/* generated by your application before
'/* acquisition into listview.

Dim lPos            As Long
Dim lOffset         As Long
Dim lPtr            As Long
Dim lLb             As Long
Dim lUb             As Long
Dim lX              As Long
Dim aData1()        As String
Dim aData3()        As String
Dim sSubData1()     As String
Dim sSubData2()     As String

    '/* init temp array
    ReDim aData1(0 To 9)

    '/* base array
    aData1(0) = "Alpha"
    aData1(1) = "Beta"
    aData1(2) = "Gamma"
    aData1(3) = "Delta"
    aData1(4) = "Epsilon"
    aData1(5) = "Zeta"
    aData1(6) = "Eta"
    aData1(7) = "Theta"
    aData1(8) = "Iota"
    aData1(9) = "Kappa"
    
    '/* generate items array
    ReDim m_aData2(0 To lCount - 1)
    '/* merge arrays to size
    For lPos = lUb To (lCount - 1) Step 10
        '/* create a 'scratch array' to avoid pointer duplication
        aData3 = aData1
        For lOffset = 0 To 9
            '/* copy the pointer to the dest array
            CopyMemBv VarPtr(m_aData2(lOffset + lPos)), VarPtr(aData3(lOffset)), &H4
            '/* deallocate the string
            CopyMemBr ByVal VarPtr(aData3(lOffset)), 0&, &H4
        Next lOffset
    Next lPos
    
    '/* generate subitems array
    lLb = LBound(m_aData2)
    lUb = UBound(m_aData2)
    lPos = 0
    With m_HLIStc(0)
        ReDim .SubItem1(lLb To lUb)
        ReDim .SubItem2(lLb To lUb)
        ReDim .lIcon(lLb To lUb)
        Do
            .lIcon(lPos) = 0
            .SubItem1(lPos) = "SubItem1 " & lPos
            .SubItem2(lPos) = "SubItem2 " & lPos
            lPos = lPos + 1
        Loop Until lPos > lUb
    End With

    '/* copy items array
    CopyMemBv GetASPtr(m_HLIStc(0).Item), GetASPtr(m_aData2), &H4
    '/* resolve pointer
    CopyMemBr ByVal GetASPtr(m_aData2), 0&, &H4
    
End Sub

Private Sub PutArray()
'/* forward struct pointer into library

On Error GoTo Handler

    If Not StructValid Then GoTo Handler
    CopyMemBr m_lPointer, ByVal GetAVPtr(m_HLIStc), 4&
    cHyperList.p_StructPtr = m_lPointer
    cHyperList.LoadArray
    cHyperList.SetItemCount UBound(m_HLIStc(0).Item) + 1
    
Handler:
    On Error GoTo 0
    
End Sub

Private Function StructValid() As Boolean
'/* test data structure

On Error GoTo Handler

    If Not ArrayCheck(m_HLIStc(0).Item) Then GoTo Handler
    '/* success
    StructValid = True

Handler:
    On Error GoTo 0

End Function

Private Sub ResetArray()

    If ArrayCheck(m_aData2) Then
        Erase m_aData2
    End If
    ItemsArray m_lPasses
    
End Sub

Private Function ArrayCheck(ByRef sArray As Variant) As Boolean
'/* validity test

On Error GoTo Handler

    '/* an array
    If IsArray(sArray) Then
        On Error Resume Next
        '/* dimensioned
        If IsError(UBound(sArray)) Then
            GoTo Handler
        End If
    Else
        GoTo Handler
    End If

    ArrayCheck = True

Handler:
    On Error GoTo 0

End Function

Private Function ArrayExists(ByRef sArray() As String) As Boolean

On Error Resume Next

    If IsError(UBound(sArray)) Then
        GoTo Handler
    End If

    '/* success
    ArrayExists = True

Handler:
    On Error GoTo 0

End Function


'> Timing
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Reset()
'/* reset timer

    cTiming.Reset

End Sub

Private Function Results() As Long
'/* get timer results

Dim lC As Long

    lC = cTiming.Elapsed
    Results = lC

End Function

Private Sub SortTest()

Dim lC  As Long

    For lC = 0 To UBound(m_aData2) Step 1000
        Debug.Print m_aData2(lC)
    Next lC

End Sub

Private Sub Form_Unload(Cancel As Integer)

    CopyMemBr ByVal GetAVPtr(m_HLIStc), 0&, &H4
    Erase m_HLIStc
    Set cHyperList = Nothing
    Set cTiming = Nothing
    
End Sub
