VERSION 5.00
Begin VB.UserControl PB 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   255
   ScaleWidth      =   2295
   ToolboxBitmap   =   "PB.ctx":0000
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "PB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Dim HPER As Long
Const m_def_ProgressColor = &HFF
'Property Variables:
Dim m_ProgressColor As OLE_COLOR
'Event Declarations:
Event DblClick() 'MappingInfo=P,P,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = P.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    P.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = P.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    P.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = P.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set P.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = P.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    P.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub P_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&HFF
Public Property Get ProgressColor() As OLE_COLOR
    ProgressColor = m_ProgressColor
End Property

Public Property Let ProgressColor(ByVal New_ProgressColor As OLE_COLOR)
    m_ProgressColor = New_ProgressColor
    PropertyChanged "ProgressColor"
End Property

Private Sub UserControl_Initialize()
    HPER = 0
    SetPercent 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ProgressColor = m_def_ProgressColor
    HPER = 0
    SetPercent 0
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    P.BackColor = PropBag.ReadProperty("BackColor", &H80FF&)
    P.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    Set P.Font = PropBag.ReadProperty("Font", Ambient.Font)
    P.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_ProgressColor = PropBag.ReadProperty("ProgressColor", m_def_ProgressColor)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    P.Width = UserControl.Width
    P.Height = UserControl.Height
    SetPercent HPER
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", P.BackColor, &H80FF&)
    Call PropBag.WriteProperty("ForeColor", P.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Font", P.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", P.BorderStyle, 1)
    Call PropBag.WriteProperty("ProgressColor", m_ProgressColor, m_def_ProgressColor)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub
Public Function Refresh()
    SetPercent HPER
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SetPercent(Percent As Long) As Variant
    HPER = Percent
    P.DrawMode = vbCopyPen
    P.Cls
    ' Determine the width of the bar in pixels based of the percentage
    intPercentWidth = (Percent / 100) * P.Width
    P.Line (0, 0)-(intPercentWidth, P.Height - 1), ProgressColor, BF
    Dim num$
    num$ = Format$(Percent, "###") + "%"
    If num$ = "%" Then num$ = "0%"
    P.CurrentX = (P.Width / 2) - P.TextWidth(num$) / 2
    P.CurrentY = (P.ScaleHeight - P.TextHeight(num$)) / 2
    P.Print num$ 'print percent
    P.Refresh
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = P.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set P.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = P.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = P.hDC
End Property

