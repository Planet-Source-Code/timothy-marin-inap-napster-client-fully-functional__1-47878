Attribute VB_Name = "basProgress"
Option Explicit
' Windows API Types
Private Const ProgressBarColumn = 6
Private Const ProgressBarColor = &HFF

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
' Windows API Functions
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Windows API Constants
Private Const WM_PAINT = &HF
Private Const GWL_WNDPROC = (-4)
Private Const SRCCOPY = &HCC0020
Private Const NOTSRCCOPY = &H330008
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Private Const LVIR_BOUNDS = 0
' Private Variables
Private OldListViewProc As Long ' Old WndProc function for default function calling
Private lvListView As ListView  ' Reference to control

Public Sub InitPBinLV(Control As ListView)
    ' Set the global reference so the user doesn't have to send it anymore
    Set lvListView = Control
    ' Set our own custom WndProc function and save the address of the old one
    OldListViewProc = SetWindowLong(lvListView.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub UpdateProgress(PB As Control, Percent As Integer, Selected As Boolean)
    Dim intPercentWidth As Integer
    
    ' Set all the needed properties
    ' Size the control to the size of the columnheader used for the progressbar
    PB.Height = lvListView.ListItems(1).Height - 2
    PB.Width = lvListView.ColumnHeaders(ProgressBarColumn).Width - 9
    ' Clear the picture
    ' Determine the width of the bar in pixels based of the percentage
    intPercentWidth = (Percent / 100) * PB.Width
    ' Draw main bar
    If Selected = True Then
        PB.BackColor = vbHighlight
    End If
PB.SetPercent Str(Percent)
End Sub

Private Sub DrawProgress(Index As Integer, ByVal Progress As Integer)
    Dim intI As Long
    Dim tmpRect As RECT
    Dim lngOffset As Long
    
    If lvListView.ColumnHeaders(ProgressBarColumn).Width > 15 Then
        ' Get item rect
        tmpRect.Left = LVIR_BOUNDS
        SendMessage lvListView.hWnd, LVM_GETITEMRECT, Index - 1, tmpRect
        ' Update the temparary picture
        If lvListView.SelectedItem.Index = Index Then
            UpdateProgress lvListView.Parent.Picture1, Progress, True
        Else
            UpdateProgress lvListView.Parent.Picture1, Progress, False
        End If
        ' Find the left side offset for column 'ProgressSubitem'
        For intI = 1 To ProgressBarColumn - 1
            lngOffset = lngOffset + lvListView.ColumnHeaders(intI).Width
        Next
        lngOffset = lngOffset + 4
        ' BitBlt the picture on the listview
        BitBlt GetDC(lvListView.hWnd), lngOffset, tmpRect.Top + 1, lvListView.ColumnHeaders(ProgressBarColumn).Width + 1, tmpRect.Bottom - tmpRect.Top, lvListView.Parent.Picture1.hDC, 0, 0, SRCCOPY
    End If
End Sub

Public Sub SetProgress(Index As Integer, ByVal Progress As Integer)
    If Progress >= 100 Then
        Progress = 100
    ElseIf Progress <= 0 Then
        Progress = 0
    Else
        Progress = Progress
    End If
    ' Draw the changes
    DrawProgress Index, Progress
End Sub


Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        ' Paint message received
        Case WM_PAINT
            ' Do the normal painting function for the listview
            WindowProc = CallWindowProc(OldListViewProc, hWnd, uMsg, wParam, lParam)
            ' Now do the custom painting function over top of the normal one
        Case Else
            ' We only want to handle the paint message to just do the default processing
            WindowProc = CallWindowProc(OldListViewProc, hWnd, uMsg, wParam, lParam)
    End Select
End Function

