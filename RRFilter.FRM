VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   11880
   DrawStyle       =   5  'Transparent
   Icon            =   "RRFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraStoreFilter 
      BackColor       =   &H8000000A&
      Caption         =   "      Enter Filter Name        "
      Height          =   1380
      Left            =   6360
      TabIndex        =   65
      Top             =   825
      Width           =   2145
      Begin VB.CommandButton cmdAccCan 
         Caption         =   "Cancel"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   68
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdAccCan 
         Caption         =   "Accept"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtFilterName 
         Height          =   300
         Left            =   135
         MaxLength       =   16
         TabIndex        =   66
         Text            =   "txtFilterName"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdStoreFilter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Store Filter"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdFixPic 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fix Picture"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Info "
      Height          =   2115
      Left            =   180
      TabIndex        =   53
      Top             =   6000
      Width           =   1395
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabRGB"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   330
         TabIndex        =   62
         Top             =   1770
         Width           =   765
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabRGB"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   465
         TabIndex        =   61
         Top             =   1515
         Width           =   450
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabRGB"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   465
         TabIndex        =   60
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabRGB"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   465
         TabIndex        =   59
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label LabGen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   58
         Top             =   495
         Width           =   210
      End
      Begin VB.Label LabGen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "W"
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   57
         Top             =   495
         Width           =   210
      End
      Begin VB.Label LabWH 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabH"
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   56
         Top             =   705
         Width           =   600
      End
      Begin VB.Label LabWH 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabW"
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   55
         Top             =   705
         Width           =   600
      End
      Begin VB.Label LabXY 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabXY"
         Height          =   255
         Left            =   105
         TabIndex        =   54
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.Frame fraPIC 
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   4305
      TabIndex        =   51
      Top             =   1725
      Width           =   2670
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   285
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   52
         Top             =   285
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3405
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdResetMults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reset factors"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fraPICBC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "QBColor "
      Height          =   1425
      Left            =   10335
      TabIndex        =   47
      Top             =   75
      Width           =   1305
      Begin VB.PictureBox picQBColor 
         Height          =   1035
         Left            =   210
         Picture         =   "RRFilter.frx":0442
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   48
         Top             =   210
         Width           =   690
      End
      Begin VB.Label LabBC 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   990
         TabIndex        =   49
         Top             =   210
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdResetPIC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reset Picture"
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdUndoLast 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Undo Last"
      Height          =   495
      Left            =   9120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apply Filter"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   840
      Width           =   1095
   End
   Begin VB.VScrollBar vertSB 
      Height          =   6255
      Left            =   11520
      TabIndex        =   43
      Top             =   1620
      Width           =   255
   End
   Begin VB.HScrollBar horzSB 
      Height          =   255
      Left            =   4140
      TabIndex        =   42
      Top             =   7920
      Width           =   7335
   End
   Begin VB.Frame fraMat 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Matrix "
      Height          =   2010
      Left            =   180
      TabIndex        =   10
      Top             =   1395
      Width           =   3285
      Begin VB.VScrollBar vscrMat 
         Height          =   405
         Left            =   2835
         Max             =   -100
         Min             =   100
         TabIndex        =   38
         Top             =   795
         Value           =   -9
         Width           =   285
      End
      Begin VB.CommandButton cmdMatReset 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reset Matrix"
         Height          =   420
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1500
         Width           =   600
      End
      Begin VB.PictureBox picSum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   60
         Picture         =   "RRFilter.frx":2284
         ScaleHeight     =   1020
         ScaleWidth      =   525
         TabIndex        =   36
         Top             =   480
         Width           =   525
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   210
         Width           =   330
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   1
         Left            =   1245
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   195
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   195
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   3
         Left            =   1995
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   195
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   4
         Left            =   2370
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   195
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   5
         Left            =   885
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   585
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   6
         Left            =   1260
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   570
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   7
         Left            =   1590
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   570
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   8
         Left            =   1995
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   555
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   9
         Left            =   2355
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   570
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   10
         Left            =   885
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   915
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   11
         Left            =   1260
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   900
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   12
         Left            =   1605
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   900
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   13
         Left            =   2025
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   900
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   14
         Left            =   2385
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   930
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   15
         Left            =   915
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1290
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   16
         Left            =   1245
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1275
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   17
         Left            =   1590
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1260
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   18
         Left            =   1995
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1260
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   19
         Left            =   2340
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1260
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   20
         Left            =   930
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1590
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   21
         Left            =   1230
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1620
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   22
         Left            =   1590
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1590
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   23
         Left            =   1965
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1590
         Width           =   270
      End
      Begin VB.TextBox txtMat 
         Height          =   285
         Index           =   24
         Left            =   2340
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1605
         Width           =   270
      End
   End
   Begin VB.ListBox LisUser 
      Height          =   2010
      Left            =   1740
      TabIndex        =   9
      Top             =   6120
      Width           =   1515
   End
   Begin VB.ListBox LisManSantos 
      Height          =   1620
      Left            =   1740
      TabIndex        =   8
      Top             =   4080
      Width           =   1515
   End
   Begin VB.ListBox LisUnLimReal 
      Height          =   1620
      Left            =   180
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CheckBox chkQB 
      BackColor       =   &H00FFFFFF&
      Caption         =   "QB"
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   900
      Width           =   615
   End
   Begin VB.TextBox txtf 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtdf 
      Height          =   315
      Left            =   4425
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtmf2 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtmf1 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   900
      Width           =   555
   End
   Begin VB.CheckBox chkAbs 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abs"
      Height          =   315
      Left            =   315
      TabIndex        =   1
      Top             =   900
      Width           =   570
   End
   Begin VB.PictureBox picEquation 
      Height          =   735
      Left            =   180
      Picture         =   "RRFilter.frx":3F76
      ScaleHeight     =   675
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   120
      Width           =   6435
   End
   Begin VB.Shape Shap 
      BackColor       =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   11520
      Top             =   7920
      Width           =   270
   End
   Begin VB.Shape SshpPIC 
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   4140
      Top             =   1620
      Width           =   7335
   End
   Begin VB.Label LabUser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label LabManSantos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manuel Santos"
      Height          =   435
      Left            =   1740
      TabIndex        =   40
      Top             =   3615
      Width           =   705
   End
   Begin VB.Label LabUnLimReal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unlimited Realities"
      Height          =   465
      Left            =   165
      TabIndex        =   39
      Top             =   3585
      Width           =   750
   End
   Begin VB.Menu mnuLoadPicture 
      Caption         =   "Load Picture"
   End
   Begin VB.Menu mnuSavePicture 
      Caption         =   "Save Picture"
   End
   Begin VB.Menu mnuLoadFilters 
      Caption         =   "Load User Filters"
   End
   Begin VB.Menu mnuSaveFilters 
      Caption         =   "Save User Filters"
      Begin VB.Menu mnuAddFilterToExisting 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuSaveNewFilterFile 
         Caption         =   "New"
      End
   End
   Begin VB.Menu mnuVBASM 
      Caption         =   "VB"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RRFilter.frm
' Interactive Filter Design  by Robert Rayment   25/11/01

Option Base 1
DefLng A-W
DefSng X-Z


Private Sub Form_Load()

Caption = "Interactive Filter Design V2 by Robert Rayment"

'===================================================================
'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

' This for save last Folder, on exit to INI file
AppName$ = "IFilter"
GetInifile
'===================================================================

INIT

UR ' Unlimited Realities filters
MS ' Manuel Santos filters

NumUserFilters = 0

Loadmcode PathSpec$ & "IFilter.bin", IFilterMC()

ASM = False

End Sub

Private Sub Form_Resize()

Top = 0
Left = 0
Width = 11999
Height = 8999

End Sub


'#######  STORE FILTER ###############################

Private Sub cmdStoreFilter_Click()

fraStoreFilter.Visible = True

End Sub

Private Sub cmdAccCan_Click(Index As Integer)

' FILTER NAME

If Index = 0 And txtFilterName <> "" Then

   ' STORE FILTER

   NumUserFilters = NumUserFilters + 1

   If NumUserFilters = 1 Then
      ReDim USName$(1)
      ReDim zUSFactors(6, 1), USMatrix(5, 5, 1)
   Else
      ReDim Preserve USName$(NumUserFilters)
      ReDim Preserve zUSFactors(6, NumUserFilters), USMatrix(5, 5, NumUserFilters)
   End If

   USName$(NumUserFilters) = txtFilterName.Text

   Form1.LisUser.Clear
   For i = 1 To NumUserFilters
      Form1.LisUser.AddItem USName$(i)
   Next i

   If chkAbs.Value = False Then
      Fac$ = "0,"
   Else
      Fac$ = "1,"
   End If
   Fac$ = Fac$ & txtmf1.Text & ","
   Fac$ = Fac$ & txtmf2.Text & ","
   Fac$ = Fac$ & txtdf.Text & ","
   Fac$ = Fac$ & txtf.Text & ","
   If chkQB.Value = False Then
      Fac$ = Fac$ + "0"
   Else
      Fac$ = Fac$ + "1"
   End If
   
   FillFactors zUSFactors(), NumUserFilters, Fac$
   
   iy = 1
   For nIndex = 0 To 24 Step 5
      tMat$ = txtMat(nIndex).Text
      For i = nIndex + 1 To nIndex + 4
         tMat$ = tMat$ & "," & txtMat(i).Text
      Next i
      FillMatrix USMatrix(), iy, NumUserFilters, tMat$
      iy = iy + 1
   Next nIndex

Else  ' CANCEL or ""
End If

fraStoreFilter.Visible = False

End Sub


'#####  RESET FACTORS ########################################

Private Sub cmdResetMults_Click()

'Public zmf1, zmf2, zdf, zf ' Multiplying factors (Single)
'Abs, zmf1, zmf2, zdf, zf|QB ' factors

chkAbs.Value = Unchecked
zmf1 = 1
zmf2 = 1
zdf = 1
zf = 0
chkQB.Value = Unchecked
txtmf1.Text = "1"
txtmf2.Text = "1"
txtdf.Text = "1"
txtf.Text = "1"

End Sub

Private Sub chkQB_Click()

If chkQB.Value = Checked Then
   txtf.Text = "0"
   zf = 0
End If

End Sub

'####### RESET MAT #################################

Private Sub cmdMatReset_Click()

' RESET MATRIX
For MatIndex = 0 To 24
   iy = 1 + MatIndex \ 5
   ix = 1 + MatIndex - 5 * (MatIndex \ 5)
   If MatIndex = 12 Then
      Mat(ix, iy) = 1
      txtMat(MatIndex).Text = "1"
   Else
      Mat(ix, iy) = 0
      txtMat(MatIndex).Text = "0"
   End If
Next MatIndex

MatIndex = 12
vscrMat.Value = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
'======================================================================
' Save last folders to ini file
IniSpec$ = PathSpec$ & AppName$ & ".ini"
Open IniSpec$ For Output As #1

Print #1, LoadDir$
Print #1, SaveDir$
Close #1
'======================================================================
End
End Sub

'#######  LIST SELECTIONS ##############################

Private Sub LisManSantos_Click()

'ReDim MSName$(8)
'ReDim zMSFactors(6, 8), MSMatrix(5, 5, 8)

'ReDim MSName$(Item)
'ReDim zMSFactors(6, Item), MSMatrix(ix 1-5, iy 1-5, Item)
'chkAbs zmf1, zmf2, zdf, zf, chkQB ' 6 zFactors

Item = LisManSantos.ListIndex + 1
ItemName$ = LisManSantos.List(LisManSantos.ListIndex)

If zMSFactors(1, Item) = 0 Then chkAbs.Value = Unchecked Else chkAbs.Value = Checked
zmf1 = zMSFactors(2, Item): txtmf1.Text = Str$(zmf1)
zmf2 = zMSFactors(3, Item): txtmf2.Text = Str$(zmf2)
zdf = zMSFactors(4, Item): txtdf.Text = Str$(zdf)
zf = zMSFactors(5, Item): txtf.Text = Str$(zf)
If zMSFactors(6, Item) = 0 Then chkQB.Value = Unchecked Else chkQB.Value = Checked

ReDim Mat(5, 5)
For iy = 1 To 5
For ix = 1 To 5
   Mat(ix, iy) = MSMatrix(ix, iy, Item)
   txtMat((ix - 1) + (iy - 1) * 5).Text = Trim$(Str$(Mat(ix, iy)))
Next ix
Next iy

vscrMat.Value = Mat(3, 3)
MatIndex = 12

End Sub

Private Sub LisUnLimReal_Click()

'ReDim URName$(8)
'ReDim zURFactors(6, 8), URMatrix(5, 5, 8)

'ReDim URName$(Item)
'ReDim zURFactors(6, Item), URMatrix(ix 1-5, iy 1-5, Item)
'chkAbs zmf1, zmf2, zdf, zf, chkQB ' 6 zFactors

Item = LisUnLimReal.ListIndex + 1
ItemName$ = LisUnLimReal.List(LisUnLimReal.ListIndex)

If zURFactors(1, Item) = 0 Then chkAbs.Value = Unchecked Else chkAbs.Value = Checked
zmf1 = zURFactors(2, Item): txtmf1.Text = Str$(zmf1)
zmf2 = zURFactors(3, Item): txtmf2.Text = Str$(zmf2)
zdf = zURFactors(4, Item): txtdf.Text = Str$(zdf)
zf = zURFactors(5, Item): txtf.Text = Str$(zf)
If zURFactors(6, Item) = 0 Then chkQB.Value = Unchecked Else chkQB.Value = Checked

ReDim Mat(5, 5)
For iy = 1 To 5
For ix = 1 To 5
   Mat(ix, iy) = URMatrix(ix, iy, Item)
   txtMat((ix - 1) + (iy - 1) * 5).Text = Trim$(Str$(Mat(ix, iy)))
Next ix
Next iy

vscrMat.Value = Mat(3, 3)
MatIndex = 12

End Sub


Private Sub LisUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'ReDim USName$(8)
'ReDim zUSFactors(6, 8), USMatrix(5, 5, 8)

'ReDim USName$(Item)
'ReDim zUSFactors(6, Item), USMatrix(ix 1-5, iy 1-5, Item)
'chkAbs zmf1, zmf2, zdf, zf, chkQB ' 6 zFactors


If Button = 1 Then ' LC
   
   If NumUserFilters = 0 Then Exit Sub
   
   Item = LisUser.ListIndex + 1
   
   ItemName$ = LisUser.List(LisUser.ListIndex)

   If zUSFactors(1, Item) = 0 Then chkAbs.Value = Unchecked Else chkAbs.Value = Checked
   zmf1 = zUSFactors(2, Item): txtmf1.Text = Str$(zmf1)
   zmf2 = zUSFactors(3, Item): txtmf2.Text = Str$(zmf2)
   zdf = zUSFactors(4, Item): txtdf.Text = Str$(zdf)
   zf = zUSFactors(5, Item): txtf.Text = Str$(zf)
   If zUSFactors(6, Item) = 0 Then chkQB.Value = Unchecked Else chkQB.Value = Checked

   ReDim Mat(5, 5)
   For iy = 1 To 5
   For ix = 1 To 5
      Mat(ix, iy) = USMatrix(ix, iy, Item)
      txtMat((ix - 1) + (iy - 1) * 5).Text = Trim$(Str$(Mat(ix, iy)))
   Next ix
   Next iy

   vscrMat.Value = Mat(3, 3)
   MatIndex = 12

ElseIf Button = 2 Then ' RC
   
   Item = LisUser.ListIndex + 1
   
   If Item <= 0 Then
      
      Message$ = "No item selected"
      ynYES = 0
      ynNO = 0
      ynOK = 1
      RRMsg Message$, ynYES, ynNO, ynOK
      Exit Sub
   
   Else
      
      Message$ = "Delete Selected Filter ?"
      ynYES = 1
      ynNO = 1
      ynOK = 0
      RRMsg Message$, ynYES, ynNO, ynOK
      
      'resp = MsgBox("", 4 + 32, "Delete Selected Filter ?")
      If ynNO = 1 Then     'NO pressed
      
         Exit Sub
   
      ElseIf ynYES = 1 Then 'YES pressed
      
         If Item > 1 And Item = NumUserFilters Then
            
            NumUserFilters = NumUserFilters - 1
            ReDim Preserve USName$(NumUserFilters)
            ReDim Preserve zUSFactors(6, NumUserFilters), USMatrix(5, 5, NumUserFilters)
            Form1.LisUser.Clear
            For i = 1 To NumUserFilters
               Form1.LisUser.AddItem USName$(i)
            Next i
      
         ElseIf NumUserFilters = 1 Then
         
            Erase USName$, zUSFactors, USMatrix
            Form1.LisUser.Clear
            NumUserFilters = 0
      
         ElseIf Item >= 1 Then
         
            For i = Item To NumUserFilters - 1
               USName$(i) = USName$(i + 1)
               For j = 1 To 6
                  zUSFactors(j, i) = zUSFactors(j, i + 1)
               Next j
               For j = 1 To 5
               For k = 1 To 5
                  USMatrix(j, k, i) = USMatrix(j, k, i + 1)
               Next k
               Next j
            Next i
         
            NumUserFilters = NumUserFilters - 1
            Form1.LisUser.Clear
            For i = 1 To NumUserFilters
               Form1.LisUser.AddItem USName$(i)
            Next i
      
         End If
      
      End If
   
   End If

End If

End Sub

Private Sub mnuVBASM_Click()

ASM = Not ASM
If ASM Then
   mnuVBASM.Caption = "ASM"
Else
   mnuVBASM.Caption = "VB"
End If

End Sub

'######  SHOW PIC INFO #############################

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   LabXY.Caption = Str$(X) & "," & Str$(Y)
   
   LongCul = GetPixel(PIC.HDC, X, Y)
   
   Red = LongCul And &HFF&
   Green = (LongCul And &HFF00&) / &H100&
   Blue = (LongCul And &HFF0000) / &H10000

   LabRGB(0) = Str$(Red)
   LabRGB(1) = Str$(Green)
   LabRGB(2) = Str$(Blue)
   LabRGB(3) = Str$(LongCul)

End Sub

'##### TEXT INPUTS #####################################

Private Sub txtdf_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc(".")
Case 8, 127
Case Is >= 48
   If KeyAscii > 57 Then KeyAscii = 0
Case 13
      If txtdf.Text = "" Then txtdf.Text = "1"
      If txtdf.Text = "0" Then txtdf.Text = "1"
      KeyAscii = 0
Case Else
   KeyAscii = 0
End Select

End Sub

Private Sub txtf_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc(".")
Case 8, 127
Case Is >= 48
   If KeyAscii > 57 Then KeyAscii = 0
Case 13
      txtf.Text = Trim$(txtf.Text)
      
      If txtf.Text = "" And chkQB.Value = Unchecked Then
         txtf.Text = "0"
      ElseIf txtf.Text <> "" And chkQB.Value = Checked Then
         chkQB.Value = Unchecked
      End If
      
      KeyAscii = 0
Case Else
   KeyAscii = 0
End Select

End Sub

Private Sub txtMat_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case KeyAscii
Case Asc("-"), Asc("+") ',Asc(".")
Case 8, 127
Case Is >= 48
   If KeyAscii > 57 Then KeyAscii = 0
Case 13
      If txtMat(Index).Text = "" Then txtMat(Index).Text = "0"
      If Index + 1 <= 24 Then
      txtMat(Index + 1).SetFocus
      txtMat(Index + 1).SelStart = Len(txtMat(Index + 1).Text)
      End If
      
      KeyAscii = 0
Case Else
   KeyAscii = 0
End Select

End Sub

Private Sub txtMat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

MatIndex = Index
vscrMat.Value = Val(txtMat(MatIndex).Text)

End Sub

Private Sub txtmf1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc(".")
Case 8, 127
Case Is >= 48
   If KeyAscii > 57 Then KeyAscii = 0
Case 13
      If txtmf1.Text = "" Then txtmf1.Text = "0"
      KeyAscii = 0
Case Else
   KeyAscii = 0
End Select

End Sub

Private Sub txtmf2_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case Asc("-"), Asc("+"), Asc(".")
Case 8, 127
Case Is >= 48
   If KeyAscii > 57 Then KeyAscii = 0
Case 13
      If txtmf2.Text = "" Then txtmf2.Text = "0"
      KeyAscii = 0
Case Else
   KeyAscii = 0
End Select

End Sub

Private Sub vscrMat_Change()

iy = 1 + MatIndex \ 5
ix = 1 + MatIndex - 5 * (MatIndex \ 5)
Mat(ix, iy) = vscrMat.Value

txtMat(MatIndex).Text = Trim$(Str$(Mat(ix, iy)))

End Sub

'###  FILL MAT FROM txtMAT ###############################

Public Sub FillMat()

For MIndex = 0 To 24
   iy = 1 + MIndex \ 5
   ix = 1 + MIndex - 5 * (MIndex \ 5)
   Mat(ix, iy) = Val(txtMat(MIndex).Text)
Next MIndex

End Sub

'##### SET QB COLORS ################################

Private Sub picQBColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   LongCul = GetPixel(picQBColor.HDC, X, Y)
   
   QBLongColor = LongCul
   
   QBRed = LongCul And &HFF&
   QBGreen = (LongCul And &HFF00&) / &H100&
   QBBlue = (LongCul And &HFF0000) / &H10000

   LabBC.BackColor = RGB(QBRed, QBGreen, QBBlue)

End Sub

'#### INIT ######################################

Private Sub INIT()

'Public PicFrameW, PicFrameH ' Size of PIC frame container
PicFrameW = 480
PicFrameH = 408

With fraPIC
   .Top = 112
   .Left = 280
   .Width = PicFrameW
   .Height = PicFrameH
End With
With PIC
   .Top = 0
   .Left = 0
   .Width = 15 * PicFrameW
   .Height = 15 * PicFrameH
End With


ReDim Mat(5, 5)
For iy = 1 To 5
For ix = 1 To 5
   Mat(ix, iy) = 1
   With txtMat((ix - 1) + (iy - 1) * 5)
      .Text = Trim$(Str$(Mat(ix, iy)))
      .MaxLength = 3
      .Width = 384
      .Height = 285
      .Top = (iy - 1) * 320 + 224
      .Left = (ix - 1) * 384 + 672 + 160
   End With
Next ix
Next iy

zmf1 = 0
txtmf1.Text = Str$(zmf1)
txtmf1.MaxLength = 6

zmf2 = 1
txtmf2.Text = Str$(zmf2)
txtmf2.MaxLength = 6

zdf = 1
txtdf.Text = Str$(zdf)
txtdf.MaxLength = 6

zf = 0
txtf.Text = Str$(zf)
txtf.MaxLength = 6

chkAbs.Value = Unchecked
chkQB.Value = Unchecked
MatIndex = 12

vscrMat.Max = -1000
vscrMat.Min = 1000
vscrMat.LargeChange = 1
vscrMat.SmallChange = 1
vscrMat.Value = 1

DisableControls

QBRed = 0
QBGreen = 160
QBBlue = 160
QBLongColor = RGB(QBRed, QBGreen, QBBlue)
LabBC.BackColor = QBLongColor

fraStoreFilter.Visible = False

End Sub

Private Sub DisableControls()

cmdApply.Enabled = False
cmdResetPIC.Enabled = False
cmdUndoLast.Enabled = False
cmdFixPic.Enabled = False

End Sub

Private Sub EnableControls()

cmdApply.Enabled = True
cmdResetPIC.Enabled = True
cmdUndoLast.Enabled = True
cmdFixPic.Enabled = True

End Sub

'###### SCROLL PICTURE #######################

Private Sub HorzSB_Change()
   
   PIC.Left = -horzSB.Value * 15
   PIC.Refresh

End Sub

Private Sub VertSB_Change()
   
   PIC.Top = -vertSB.Value * 15
   PIC.Refresh

End Sub

'### COMMANDS PICTURE #################################

Private Sub cmdApply_Click()
   
' APPLY FILTER

If txtmf1.Text <> "" Then zmf1 = Val(txtmf1.Text) Else zmf1 = 0
If txtmf2.Text <> "" Then zmf2 = Val(txtmf2.Text) Else zmf2 = 0
If txtdf.Text <> "" Then zdf = Val(txtdf.Text) Else zdf = 1
If zdf = 0 Then
   zdf = 1
   txtdf.Text = "1"
End If
If txtf.Text <> "" Then zf = Val(txtf.Text) Else zf = 0

FillMat

' 2 -> 3 for Undo Last before applying filter
CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 2), PalSize

ApplyFilter

End Sub

Private Sub cmdFixPic_Click()

' Make blocks 1 & 3 = block 2

' 2 -> 1
CopyMemory PalBGR(1, 1, 1, 1), PalBGR(1, 1, 1, 2), PalSize
' 2 -> 3
CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 2), PalSize

End Sub

Private Sub cmdResetPIC_Click()
   
   ' REST Copy 1->2 1->3 Show 1
   CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 1), PalSize
   CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 1), PalSize
   ShowPalBGR 1

End Sub

Private Sub cmdUndoLast_Click()
   
   ' UNDO LAST  Copy 3-2 Show 2
   CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 3), PalSize
   ShowPalBGR 2

End Sub

'###### MENUS #############################################

Private Sub mnuSaveFilters_Click()

If NumUserFilters = 0 Then
      Beep
      Message$ = "No User Filters"
      ynYES = 0
      ynNO = 0
      ynOK = 1
      RRMsg Message$, ynYES, ynNO, ynOK
      Exit Sub
End If

End Sub

Private Sub mnuLoadFilters_Click()

'======================================================================
Title$ = "Load Filter File"
Choice$ = "Filters (*.ufi)|*.ufi"
InitDir$ = LoadDir$
LoadSave = True
LSDialog Title$, Choice$, InitDir$, LoadSave
' Output:  Public LoadFileSpec$, SaveFileSpec$
'======================================================================

If LoadFileSpec$ <> "" Then

   On Error GoTo LoadFilterError

   ' LOAD FILTER(S) from file
   Open LoadFileSpec$ For Input As #1
   
   NumUserFilters = 1
   
   Do Until EOF(1)

      If NumUserFilters = 1 Then
         ReDim USName$(1)
         ReDim zUSFactors(6, 1), USMatrix(5, 5, 1)
      Else
         ReDim Preserve USName$(NumUserFilters)
         ReDim Preserve zUSFactors(6, NumUserFilters), USMatrix(5, 5, NumUserFilters)
      End If

      Line Input #1, USName$(NumUserFilters)
      Line Input #1, Fac$
      FillFactors zUSFactors(), NumUserFilters, Fac$
   
      iy = 1
      For nIndex = 0 To 24 Step 5
         Line Input #1, tMat$
         FillMatrix USMatrix(), iy, NumUserFilters, tMat$
         iy = iy + 1
      Next nIndex
      
      NumUserFilters = NumUserFilters + 1
   
   Loop
   
   NumUserFilters = NumUserFilters - 1
   
   Form1.LisUser.Clear
   For i = 1 To NumUserFilters
      Form1.LisUser.AddItem USName$(i)
   Next i

End If

Close

Exit Sub

'=============
LoadFilterError:

Close

MsgBox ("Filter File Error")
            
Erase USName$, zUSFactors, USMatrix
Form1.LisUser.Clear
NumUserFilters = 0

End Sub

Private Sub mnuSaveNewFilterFile_Click()

'======================================================================
Title$ = "Save To New Filter File"
Choice$ = "Pictures (*.ufi)|*.ufi"
InitDir$ = LoadDir$
LoadSave = False     ' Save
LSDialog Title$, Choice$, InitDir$, LoadSave, "ufi"
' Output:  Public LoadFileSpec$, SaveFileSpec$
'======================================================================
' User Filters
'Public USName$()
'Public zUSFactors(), USMatrix()
'Public NumUserFilters

If SaveFileSpec$ <> "" Then
   
   Open SaveFileSpec$ For Output As #1
   
   SaveAll

End If

End Sub

Private Sub mnuAddFilterToExisting_Click()

'======================================================================
Title$ = "Add To Existing Filter File"
Choice$ = "Pictures (*.ufi)|*.ufi"
InitDir$ = LoadDir$
LoadSave = False     ' Save
LSDialog Title$, Choice$, InitDir$, LoadSave, "ufi"
' Output:  Public LoadFileSpec$, SaveFileSpec$
'======================================================================

If SaveFileSpec$ <> "" Then

   Open SaveFileSpec$ For Append As #1
   
   SaveAll

End If

End Sub

Private Sub SaveAll()

' Save to #1

   For i = 1 To NumUserFilters
      
      Print #1, USName$(i)
      
      Fac$ = ""
      For j = 1 To 6
         If j < 6 Then
            Fac$ = Fac$ + Trim$(Str$(zUSFactors(j, i))) & ","
         Else
            Fac$ = Fac$ + Trim$(Str$(zUSFactors(j, i)))
         End If
      Next j
   
      Print #1, Fac$
      
      For iy = 1 To 5
      tMat$ = ""
      For ix = 1 To 5
         If ix < 5 Then
            tMat$ = tMat$ & Trim$(Str$(USMatrix(ix, iy, i))) & ","
         Else
            tMat$ = tMat$ & Trim$(Str$(USMatrix(ix, iy, i)))
         End If
      Next ix
      
      Print #1, tMat$
      
      Next iy
   
   Next i
   
   Close #1
   
End Sub


Private Sub mnuSavePicture_Click()

'======================================================================
Title$ = "Save Picture file"
Choice$ = "Pictures (*.bmp)|*.bmp"
InitDir$ = LoadDir$
LoadSave = False     ' Save
LSDialog Title$, Choice$, InitDir$, LoadSave, "bmp"
' Output:  Public LoadFileSpec$, SaveFileSpec$
'======================================================================

If SaveFileSpec$ <> "" Then

   'SavePicture PIC.Picture, SaveFileSpec$
   SavePicture PIC.Image, SaveFileSpec$

End If

End Sub

Private Sub mnuLoadPicture_Click()

On Error GoTo PICError
   
'======================================================================
Title$ = "Load a Picture file"
Choice$ = "Pictures (*.bmp *.gif *.jpg *.wmf *.ico *.cur)|*.bmp;*.gif;*.jpg;*.wmf;*.ico;*.cur|All files|*.*"
InitDir$ = LoadDir$
LoadSave = True
LSDialog Title$, Choice$, InitDir$, LoadSave
' Output:  Public LoadFileSpec$, SaveFileSpec$
'======================================================================

If LoadFileSpec$ <> "" Then
   
   'LabFileName.Caption = LoadFileSpec$
   
   PIC.Visible = False
   PIC.Picture = LoadPicture(LoadFileSpec$)
   
   PICW = PIC.Width / 15
   PICH = PIC.Height / 15
   
   'iXp = PICW \ 2
   'iYp = PICH \ 2
   
   LabWH(0).Caption = Str$(PICW)
   LabWH(1).Caption = Str$(PICH)
   
   ' NBNB can be changed for larger pictures if you want!!
   If PICW < 16 Or PICH < 16 Or PICW > 1024 Or PICH > 1024 Then
      MsgBox ("Both Width & Height of picture must be >= 16 and <= 1024")
      PIC.Picture = LoadPicture()
      Exit Sub
   End If
   
   SetPICScrollBars

   '---------------------
   GeneratePalBGRs
   '---------------------
   ShowPalBGR 3
   '---------------------

   EnableControls


   'CheckNums 2
   'CheckNums 3

   'shpFlash.BackColor = QBColor(10)
   
   MousePointer = vbHourglass
   Refresh
   
End If

MousePointer = vbDefault
Exit Sub

'=========
PICError:

MsgBox ("Error loading picture")
On Error GoTo 0

End Sub

'###### END MENUS #############################################

Private Sub GeneratePalBGRs()

GETDIBS PIC.Image

' Mem storage

'PalBGR for 32 bitcount

' BLUE     GREEN     RED       ALPHA
'(1,1,1,1)(2,1,1,1),(3,1,1,1),(4,1,1,1),,,(1,X,1,1),(2,X,1,1),(3,X,1,1),(4,X,1,1)

'(1,1,Y,1)(2,1,Y,1),(3,1,Y,1),(4,1,1,1),,,(1,X,Y,1),(2,X,Y,1),(3,X,Y,1),(4,X,Y,1)
'-----------------------------------------------------------------------------------
'(1,1,1,2)(2,1,1,2),(3,1,1,2),(4,1,1,2),,,(1,X,1,2),(2,X,1,2),(3,X,1,2),(4,X,1,2)

'(1,1,Y,2)(2,1,Y,1),(3,1,Y,1),(4,1,1,1),,,(1,X,Y,2),(2,X,Y,2),(3,X,Y,2),(4,X,Y,2)
'-----------------------------------------------------------------------------------
'(1,1,1,3)(2,1,1,3),(3,1,1,3),(4,1,1,3),,,(1,X,1,3),(2,X,1,3),(3,X,1,3),(4,X,1,3)

'(1,1,Y,3)(2,1,Y,3),(3,1,Y,3),(4,1,Y,3),,,(1,X,Y,3),(2,X,Y,3),(3,X,Y,3),(4,X,Y,3)
'-----------------------------------------------------------------------------------

' Save 2 copies of palette picture
PalSize = 4 * PICW * PICH     ' Bytes
CopyMemory PalBGR(1, 1, 1, 2), PalBGR(1, 1, 1, 1), PalSize
CopyMemory PalBGR(1, 1, 1, 3), PalBGR(1, 1, 1, 1), PalSize

End Sub

Private Sub GETDIBS(ByVal PICIM As Long)

' PICM is PIC.Image - handle to PIC picbox

On Error GoTo DIBError

'Get info on picture loaded into PIC
GetObjectAPI PICIM, Len(bmp), bmp

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

' Set up bm struc for GetDIBits & StretchDIBits
With bm.bmiH
   .biSize = 40
   .biwidth = bmp.bmWidth
   .biheight = bmp.bmHeight
   .biPlanes = 1
   .biBitCount = 32          ' Sets up BGRA pixels
   .biCompression = 0
   BytesPerScanLine = ((((.biwidth * .biBitCount) + 31) \ 32) * 4)
   PadBytesPerScanLine = _
       BytesPerScanLine - (((.biwidth * .biBitCount) + 7) \ 8)
   .biSizeImage = BytesPerScanLine * Abs(.biheight)
End With

' Not sure about necessity of  31) \ 32)  &  7) \ 8)
   
' Set PalBGR to receive color bytes BGRA
ReDim PalBGR(4, PICW, PICH, 3) As Byte

' Load color bytes to 1st third of PalBGR
ret = GetDIBits(NewDC, PICIM, 0, PICH, PalBGR(1, 1, 1, 1), bm, 1)

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC
Exit Sub
'==========
DIBError:
  MsgBox "Error"
  On Error GoTo 0
End Sub

Private Sub SetPICScrollBars()
 
   If PICW > PicFrameW Then
      horzSB.Enabled = True
      horzSB.Min = 0
      horzSB.Max = PICW - PicFrameW
      horzSB.LargeChange = PICW / 10
      horzSB.Value = 0
   Else
      horzSB.Enabled = False
   End If
   If PICH > PicFrameH Then
      vertSB.Enabled = True
      vertSB.Min = 0
      vertSB.Max = PICH - PicFrameH
      vertSB.LargeChange = PICH / 10
      vertSB.Value = 0
   Else
      vertSB.Enabled = False
   End If

End Sub

'###### CALLING FORM MESSAGE BOX ######################

Private Sub RRMsg(Message$, ynYES, ynNO, ynOK)

' Simple message box avoiding BEEP

'ynYES 0,1   0 no cmdYNO(0) YES
'ynNO  0,1   0 no cmdYNO(1) NO
'ynOK  0,1   0 no cmdYNO(2) OK

If ynYES = 0 Then frmMsg.cmdYNO(0).Visible = False
If ynNO = 0 Then frmMsg.cmdYNO(1).Visible = False
If ynOK = 0 Then frmMsg.cmdYNO(2).Visible = False

frmMsg.Caption = Message$

frmMsg.Show vbModal

End Sub

Private Sub Loadmcode(InFile$, MCCode() As Byte)
'Load machine code into InCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox (InFile$ & " missing")
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox (InFile$ & " missing")
   DoEvents
   Unload Form1
   End
End If
ReDim MCCode(MCSize&)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

