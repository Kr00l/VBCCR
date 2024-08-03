VERSION 5.00
Begin VB.Form UserEditingForm 
   Caption         =   "CellEditing Demo"
   ClientHeight    =   6900
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13830
   ScaleHeight     =   6900
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Allow incremental search"
      Height          =   1335
      Left            =   7920
      TabIndex        =   12
      Top             =   5400
      Width           =   2295
      Begin VB.OptionButton Option9 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option8 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calendar input text editable"
      Height          =   1335
      Left            =   5520
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
      Begin VB.OptionButton Option6 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Edit on return key"
      Height          =   1335
      Left            =   3120
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
      Begin VB.OptionButton Option5 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ValidateEdit (Cancel=True)"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   2775
      Begin VB.OptionButton Option3 
         Caption         =   "Discard changes silently"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Remain active for grid only"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Remain active for whole form"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   11040
      TabIndex        =   15
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Other control to test validation"
      Top             =   4800
      Width           =   10575
   End
   Begin VBFlexGridDemo.VBFlexGrid VBFlexGrid1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7435
      MouseTrack      =   -1  'True
      Rows            =   25
      Cols            =   13
      AllowUserEditing=   -1  'True
      AllowUserResizing=   3
      MergeCells      =   1
      FormatString    =   "UserEditingForm.frx":0000
      AllowReaderMode =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Editing mode OFF"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "UserEditingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private Const CC_RGBINIT As Long = &H1
Private Type TCHOOSECOLOR
lStructSize As Long
hWndOwner As LongPtr
hInstance As LongPtr
RGBResult As Long
lpCustColors As LongPtr
Flags As Long
lCustData As LongPtr
lpfnHook As LongPtr
lpTemplateName As LongPtr
End Type
#If VBA7 Then
Private Declare PtrSafe Function ChooseColor Lib "comdlg32" Alias "ChooseColorW" (ByRef lpChooseColor As TCHOOSECOLOR) As Long
#Else
Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorW" (ByRef lpChooseColor As TCHOOSECOLOR) As Long
#End If
Private Const COL_NORMAL As Long = 1
Private Const COL_ONLYNUMBERS As Long = 2
Private Const COL_CALENDARVALIDATION As Long = 3
Private Const COL_LOCKED As Long = 4
Private Const COL_REDBKCOLOR As Long = 5
Private Const COL_NOTALLOWED As Long = 6
Private Const COL_NOCLOSEBYNAVIGATIONKEY As Long = 7
Private Const COL_SINGLELINE As Long = 8
Private Const COL_MERGEDCELLS As Long = 9
Private Const COL_COMBODROPDOWN As Long = 10
Private Const COL_COMBOEDITABLE As Long = 11
Private Const COL_COMBOBUTTON As Long = 12

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
Dim i As Long
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    VBFlexGrid1.TextMatrix(i, 0) = i
Next i
VBFlexGrid1.MergeCol(COL_MERGEDCELLS) = True
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    VBFlexGrid1.TextMatrix(i, COL_NORMAL) = Chr(64 + i)
    VBFlexGrid1.TextMatrix(i, COL_LOCKED) = VBFlexGrid1.TextMatrix(i, COL_NORMAL)
    VBFlexGrid1.TextMatrix(i, COL_REDBKCOLOR) = VBFlexGrid1.TextMatrix(i, COL_NORMAL)
    VBFlexGrid1.TextMatrix(i, COL_NOTALLOWED) = VBFlexGrid1.TextMatrix(i, COL_NORMAL)
    VBFlexGrid1.TextMatrix(i, COL_NOCLOSEBYNAVIGATIONKEY) = VBFlexGrid1.TextMatrix(i, COL_NORMAL)
    VBFlexGrid1.TextMatrix(i, COL_SINGLELINE) = VBFlexGrid1.TextMatrix(i, COL_NORMAL) & "_" & Chr(64 + i + 1)
Next i
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    VBFlexGrid1.TextMatrix(i, COL_ONLYNUMBERS) = i
Next i
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1
    VBFlexGrid1.TextMatrix(i, COL_CALENDARVALIDATION) = VBA.DateAdd("d", i, Int(Now()))
Next i
VBFlexGrid1.ColComboMode(COL_CALENDARVALIDATION) = FlexComboModeCalendar
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1 - 1 Step 2
    VBFlexGrid1.TextMatrix(i, COL_MERGEDCELLS) = Chr(64 + i)
    VBFlexGrid1.TextMatrix(i + 1, COL_MERGEDCELLS) = Chr(64 + i)
Next i
Dim ComboItems As String
For i = 1 To 10
    ComboItems = ComboItems & i & vbTab & VBA.Choose(i, "Arnold", "Bob", "Charlie", "David", "Elena", "Felix", "Greg", "Hanna", "Ivan", "Jacob")
    ComboItems = ComboItems & vbTab & "Hint " & i & "/1" & vbTab & "Hint " & i & "/2"
    If i < 10 Then ComboItems = ComboItems & "|"
Next i
VBFlexGrid1.ColComboMode(COL_COMBODROPDOWN) = FlexComboModeDropDown
VBFlexGrid1.ColComboItems(COL_COMBODROPDOWN) = ComboItems ' "Arnold|Bob|Charlie|David|Elena|Felix|Greg|Hanna|Ivan|Jacob"
VBFlexGrid1.ColComboHeader(COL_COMBODROPDOWN) = "Id" & vbTab & "Name" & vbTab & "Info 1" & vbTab & "Info 2"
VBFlexGrid1.ColComboBoundColumn(COL_COMBODROPDOWN) = 1 ' zero-based column
VBFlexGrid1.ColLookup(COL_COMBODROPDOWN) = ";NULL|0;|1;Arnold|2;Bob|3;Charlie|4;David|5;Elena|6;Felix|7;Greg|8;Hanna|9;Ivan|10;Jacob"
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1 - 2 Step 3
    VBFlexGrid1.TextMatrix(i, COL_COMBODROPDOWN) = "1" ' Arnold
    VBFlexGrid1.TextMatrix(i + 1, COL_COMBODROPDOWN) = "2" ' Bob
    VBFlexGrid1.TextMatrix(i + 2, COL_COMBODROPDOWN) = "3" ' Charlie
Next i
VBFlexGrid1.ColComboMode(COL_COMBOEDITABLE) = FlexComboModeEditable
VBFlexGrid1.ColComboItems(COL_COMBOEDITABLE) = ComboItems ' "Arnold|Bob|Charlie|David|Elena|Felix|Greg|Hanna|Ivan|Jacob"
VBFlexGrid1.ColComboHeader(COL_COMBOEDITABLE) = "Id" & vbTab & "Name" & vbTab & "Info 1" & vbTab & "Info 2"
VBFlexGrid1.ColComboBoundColumn(COL_COMBOEDITABLE) = 1 ' zero-based column
For i = VBFlexGrid1.FixedRows To VBFlexGrid1.Rows - 1 - 2 Step 3
    VBFlexGrid1.TextMatrix(i, COL_COMBOEDITABLE) = "Arnold"
    VBFlexGrid1.TextMatrix(i + 1, COL_COMBOEDITABLE) = "Bob"
    VBFlexGrid1.TextMatrix(i + 2, COL_COMBOEDITABLE) = "Charlie"
Next i
VBFlexGrid1.ColComboMode(COL_COMBOBUTTON) = FlexComboModeButton
VBFlexGrid1.ColComboItems(COL_COMBOBUTTON) = vbNullString
VBFlexGrid1.AutoSize 0, VBFlexGrid1.Cols - 1, FlexAutoSizeModeColWidth, FlexAutoSizeScopeAll
End Sub

Private Sub VBFlexGrid1_DividerDblClick(ByVal Row As Long, ByVal Col As Long)
If Row = -1 Then
    VBFlexGrid1.AutoSize Col, , FlexAutoSizeModeColWidth, , , , CBool(VBFlexGrid1.ClipMode = FlexClipModeExcludeHidden)
ElseIf Col = -1 Then
    VBFlexGrid1.AutoSize Row, , FlexAutoSizeModeRowHeight, , , , CBool(VBFlexGrid1.ClipMode = FlexClipModeExcludeHidden)
End If
End Sub

Private Sub VBFlexGrid1_RowColChange()
' The combo cue can only be displayed on the current cell.
If VBFlexGrid1.Row >= VBFlexGrid1.FixedRows Then
    Select Case VBFlexGrid1.Col
        Case COL_CALENDARVALIDATION, COL_COMBODROPDOWN, COL_COMBOEDITABLE
            VBFlexGrid1.ComboCue = FlexComboCueDropDown
        Case COL_COMBOBUTTON
            VBFlexGrid1.ComboCue = FlexComboCueButton
        Case Else
            VBFlexGrid1.ComboCue = FlexComboCueNone
    End Select
Else
    VBFlexGrid1.ComboCue = FlexComboCueNone
End If
End Sub

Private Sub VBFlexGrid1_BeforeEdit(Row As Long, Col As Long, ByVal Reason As FlexEditReasonConstants, Cancel As Boolean)
' This event is for evaluation if the cell can be edited.
' Nothing has been initialized yet. So EditRow/EditCol can't be used. Instead they are passed in the parameters.
' Row and Col parameters are ByRef so they can be changed, if necessary.
' The Reason parameter is a value indicating why this event was called.
' EditReason property is not appropriate as it contains the value from the last edit which was not canceled in this event.
' EditReason can be -1 as an alias for a failed edit attempt (canceled here) or the grid was never edited before.
If Row < VBFlexGrid1.FixedRows Or Col < VBFlexGrid1.FixedCols Then
    ' Fixed cells can't be edited by the end-user. (only by code)
    ' However, here it can be ensured that this is not possible at all.
    ' Cancel = True
End If
If Col = COL_NOTALLOWED Then
    ' The last col we want to be in a special range which is not allowed to be edited.
    Cancel = True
End If
End Sub

Private Sub VBFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long, ByVal Changed As Boolean)
' This event is fired when the edit control is destroyed. It can be useful to resort the grid for example.
' EditRow/EditCol is already reset to -1. That's why they got preserved in the Row/Col parameters in this event.
' Changed parameter is True when changes were comitted.
' EditCloseMode property can be used to find out why the editing was closed.
End Sub

Private Sub VBFlexGrid1_EnterEdit()
' This event will be called always when entering edit mode. Edit control is already displayed.
Label1.Caption = "Editing mode ON (Row:" & VBFlexGrid1.EditRow & " Col: " & VBFlexGrid1.EditCol & ")"
End Sub

Private Sub VBFlexGrid1_LeaveEdit()
' This event will be called always when exiting edit mode. Edit control is still displayed.
' EditCloseMode property can be used to find out why the editing is about to be closed.
Label1.Caption = "Editing mode OFF"
End Sub

Private Sub VBFlexGrid1_EditSetupStyle(dwStyle As Long, dwExStyle As Long)
' Edit control is not created, used to modify window styles.
Select Case VBFlexGrid1.EditCol
    Case COL_ONLYNUMBERS
        Const ES_NUMBER As Long = &H2000
        dwStyle = dwStyle Or ES_NUMBER
    Case COL_LOCKED
        Const ES_READONLY As Long = &H800
        dwStyle = dwStyle Or ES_READONLY
    Case COL_SINGLELINE, COL_CALENDARVALIDATION
        Const ES_MULTILINE As Long = &H4, ES_AUTOVSCROLL As Long = &H40, ES_AUTOHSCROLL As Long = &H80
        ' If 'SingleLine' is True then the whole flex grid is single lined. ES_MULTILINE is not predefined in that case.
        ' So it is better to check for ES_MULTILINE before removing it.
        If (dwStyle And ES_MULTILINE) = ES_MULTILINE Then
            dwStyle = dwStyle And Not (ES_MULTILINE Or ES_AUTOVSCROLL)
            dwStyle = dwStyle Or ES_AUTOHSCROLL
        End If
End Select
End Sub

Private Sub VBFlexGrid1_EditSetupWindow(BackColor As stdole.OLE_COLOR, ForeColor As stdole.OLE_COLOR)
' Edit control is created but not yet displayed.
Select Case VBFlexGrid1.EditCol
    Case COL_REDBKCOLOR
        BackColor = vbRed
    Case COL_CALENDARVALIDATION
        If Option6.Value = True Then
            ' FlexComboModeCalendar now behaves like FlexComboModeDropDown when the edit control has ES_READONLY.
            ' It means always immediately popup of the calendar and commit on a date click.
            VBFlexGrid1.EditLocked = True
        End If
End Select
End Sub

Private Sub VBFlexGrid1_EditQueryClose(ByVal CloseMode As FlexEditCloseModeConstants, Cancel As Boolean)
Select Case VBFlexGrid1.EditCol
    Case COL_NOCLOSEBYNAVIGATIONKEY
        If CloseMode = FlexEditCloseModeNavigationKey Then Cancel = True
End Select
End Sub

Private Sub VBFlexGrid1_Validate(Cancel As Boolean)
' This must be handled when validation of the edit control should be for the whole form.
If Option1.Value = True Then
    If VBFlexGrid1.hWndEdit <> 0 Then ' Check if editing is active.
        ' Try to commit. The method 'CommitEdit' will fire the ValidateEdit event.
        ' Doing this way will prevent double validation in case a MsgBox is shown in the ValidateEdit event.
        Cancel = Not VBFlexGrid1.CommitEdit() ' Call VBFlexGrid1_ValidateEdit(Cancel)
    End If
End If
End Sub

Private Sub VBFlexGrid1_ValidateEdit(Cancel As Boolean)
' If validation fails the control will remain in edit mode.
' EditCloseMode property is not meaningful yet.
Select Case VBFlexGrid1.EditCol
    Case COL_CALENDARVALIDATION
        Dim Text As String
        Text = Trim$(VBFlexGrid1.EditText)
        If Not Text = vbNullString Then
            If InStr(Text, vbCrLf) Then ' Only single line entries are valid.
                Cancel = True
            Else
                Cancel = Not IsDate(Text)
            End If
            If Cancel = False Then
                ' Ensure unique date format before commit. (override possible custom format of the text box)
                VBFlexGrid1.EditText = VBFlexGrid1.ComboCalendarValue
            End If
        End If
End Select
If Cancel = True Then
    If Option3.Value = True Then
        VBFlexGrid1.CancelEdit
        Cancel = False ' Ensuring 'VBFlexGrid1_Validate' will not be blocked.
    Else
        If Cancel = True Then Beep ' Give user a minimal feedback.
    End If
End If
End Sub

Private Sub VBFlexGrid1_ComboButtonClick()
Static CustomColors(0 To 15) As Long, CustomColorsInitialized As Boolean
Select Case VBFlexGrid1.EditCol
    Case COL_COMBOBUTTON
        Dim CHCLR As TCHOOSECOLOR
        With CHCLR
        .lStructSize = LenB(CHCLR)
        .hWndOwner = Me.hWnd
        .hInstance = App.hInstance
        .Flags = CC_RGBINIT
        If CustomColorsInitialized = False Then
            Dim i As Long, IntValue As Integer
            For i = 0 To 15
                IntValue = 255 - (i * 16)
                CustomColors(i) = RGB(IntValue, IntValue, IntValue)
            Next i
            CustomColorsInitialized = True
        End If
        .lpCustColors = VarPtr(CustomColors(0))
        .RGBResult = WinColor(VBFlexGrid1.Cell(FlexCellBackColor, VBFlexGrid1.EditRow, VBFlexGrid1.EditCol))
        End With
        If ChooseColor(CHCLR) <> 0 Then
            VBFlexGrid1.Cell(FlexCellBackColor, VBFlexGrid1.EditRow, VBFlexGrid1.EditCol) = CHCLR.RGBResult
            VBFlexGrid1.CommitEdit
        Else
            VBFlexGrid1.CancelEdit
        End If
End Select
End Sub

Private Sub Option4_Click()
VBFlexGrid1.DirectionAfterReturn = FlexDirectionAfterReturnNone
End Sub

Private Sub Option5_Click()
VBFlexGrid1.DirectionAfterReturn = FlexDirectionAfterReturnEdit
End Sub

Private Sub Option8_Click()
VBFlexGrid1.AllowIncrementalSearch = False
End Sub

Private Sub Option9_Click()
VBFlexGrid1.AllowIncrementalSearch = True
End Sub
