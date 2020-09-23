VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   3  'Windows Default
   Tag             =   "Reconstructor 2.0"
   Begin VB.Frame fraDestination 
      Caption         =   "Destination Image"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save 24-bit Destination Image"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Text            =   "256"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "256"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Height:"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdResample 
      Caption         =   "&Resample"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame fraResampling 
      Caption         =   "Resampling"
      Height          =   2895
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin VB.HScrollBar hsc2 
         Height          =   220
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   12
         Top             =   840
         Value           =   1
         Width           =   2175
      End
      Begin VB.ComboBox cboSub 
         Height          =   315
         ItemData        =   "frmMain.frx":1042
         Left            =   120
         List            =   "frmMain.frx":106D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.HScrollBar hsc1 
         Height          =   220
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   11
         Top             =   600
         Value           =   1
         Width           =   2175
      End
      Begin VB.PictureBox picBC 
         Height          =   2175
         Left            =   120
         Picture         =   "frmMain.frx":1144
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   139
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "BC-spline scheme (drag to set the control point)"
         Top             =   600
         Width           =   2145
      End
      Begin VB.ComboBox cboResampler 
         Height          =   315
         ItemData        =   "frmMain.frx":F8DA
         Left            =   120
         List            =   "frmMain.frx":F8F0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lbl3 
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         ToolTipText     =   "One and only parameter for Cardinal spline. Corresponds to -C."
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl2 
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl1 
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source Image"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cboSource 
         Height          =   315
         ItemData        =   "frmMain.frx":F99C
         Left            =   120
         List            =   "frmMain.frx":F99E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblSource 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Reconstructor by Peter Scale 2003

' This sample application shows some kinds of area interpolation.
' It is not optimized for speed but for understanding.
' Due to no effective optimizations these resamplers are reference,
' e.g., very strict.

Option Explicit

' structure for OpenFile/SaveFile dialog
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As String
    lpstrFileTitle As String
    nMaxFileTitle As String
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' API function to show SaveFile dialog
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Sub ShowProgress(ByVal nPos&, ByVal nMax&)
    frmDestination.Refresh
    frmDestination.Line (0, nPos)-(frmDestination.ScaleWidth, nPos), vbRed
    Caption = FormatPercent(nPos / nMax, 0) & " " & Tag
End Sub

Private Sub cboResampler_Click()
    ' Enable or disable controls we need
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
    picBC.Visible = False
    hsc1.Visible = False
    hsc2.Visible = False
    cboSub.Visible = False
    Select Case cboResampler.ListIndex
        Case BicubicCardinal
            ' This will set a=-0.5 (C=0.5; B=0)
            picBC_MouseMove vbLeftButton, 0, 72, 122
            lbl1.Visible = True
            lbl2.Visible = True
            lbl3.Visible = True
            picBC.Visible = True
        Case BicubicBSpline
            ' B=1; C=0
            picBC_MouseMove vbLeftButton, 0, 16, 9
            lbl1.Visible = True
            lbl2.Visible = True
            picBC.Visible = True
        Case BicubicBCSpline
            ' B=1/3; C=1/3
            picBC_MouseMove vbLeftButton, 0, 53, 85
            lbl1.Visible = True
            lbl2.Visible = True
            picBC.Visible = True
        Case WindowedSinc
            cboSub.ListIndex = wLanczos
            hsc1.Value = 3
            hsc1_Scroll
            cboSub.Visible = True
            hsc1.Visible = True
            lbl1.Visible = True
    End Select
    frmPreview.Paint
End Sub

Private Sub cboSource_Click()
    Dim tBM As BITMAP, sPic As StdPicture
    Dim CDC&, CDC1&
    On Error GoTo Out
    With frmSource
        ' Load chosen picture
        .Picture = LoadPicture(App.Path & "\Images\" & cboSource.Text)
        ' Get informations about loaded bitmap
        GetObjectAPI .Picture, Len(tBM), tBM
        ' Show info about source
        lblSource.Caption = "Width: " & tBM.bmWidth & "    Height: " & tBM.bmHeight & "    BPP: " & tBM.bmBitsPixel
        ' If non 24bpp image loaded convert to it
        If tBM.bmBitsPixel <> 24 Then
            ' Create 24bpp empty (black) image
            Set sPic = CreatePicture(tBM.bmWidth, tBM.bmHeight, 24)
            CDC = CreateCompatibleDC(0) ' Temporary devices
            CDC1 = CreateCompatibleDC(0)
            DeleteObject SelectObject(CDC, .Picture) ' Source bitmap
            DeleteObject SelectObject(CDC1, sPic) ' Converted bitmap
            ' Copy between two different depths
            BitBlt CDC1, 0, 0, tBM.bmWidth, tBM.bmHeight, CDC, 0, 0, vbSrcCopy
            DeleteDC CDC: DeleteDC CDC1 ' Erase devices
            .Picture = sPic ' Set visible image
        End If
        .Move Left, Top + Height
        .Show vbModeless, Me
        SetFocus
        cmdResample.Enabled = True
    End With
Out:
End Sub

Private Sub cboSub_Click()
    sinc_window = cboSub.ListIndex
    If sinc_window = wGauss Or sinc_window = wKaiser Then
        hsc2.Value = 6
        hsc2.Visible = True
        lbl2.Visible = True
    Else
        hsc2.Visible = False
        lbl2.Visible = False
    End If
    frmPreview.Paint
End Sub

Private Sub cmdResample_Click()
    Dim tBM As BITMAP
    ' Check if destination dimensions are correct
    If IsNumeric(txtWidth) = False Or IsNumeric(txtHeight) = False Then GoTo Out
    If txtWidth < 2 Or txtWidth > 2048 Or txtHeight < 2 Or txtHeight > 2048 Then GoTo Out
    GetObjectAPI frmDestination.Picture, Len(tBM), tBM
    If tBM.bmWidth <> txtWidth Or tBM.bmHeight <> txtHeight Then
        ' If there are new dimensions then create new 24-bit bitmap
        frmDestination.Picture = CreatePicture(txtWidth, txtHeight, 24)
    End If
    frmDestination.Move frmSource.Left + frmSource.Width, Top + Height
    frmDestination.Show vbModeless, Me
    SetFocus
    ' Resample source image into destination image
    DoResample cboResampler.ListIndex, frmDestination.Picture, frmSource.Picture
    frmDestination.Refresh
    cmdSave.Enabled = True
    Caption = Tag
    Exit Sub
Out: MsgBox "Invalid destination dimensions", vbCritical
End Sub

Private Sub cmdSave_Click()
    Dim OFN As OPENFILENAME
    With OFN
        .lpstrFile = String$(260, 0)
        .nMaxFile = Len(.lpstrFile)
        .lpstrFilter = "Windows Bitmaps (*.bmp)" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
        .lpstrDefExt = "bmp"
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .Flags = 6
        .lStructSize = Len(OFN)
        ' SaveFile dialog
        If GetSaveFileName(OFN) = 0 Then Exit Sub
        ' Save image to the file
        SavePicture frmDestination.Picture, Left$(.lpstrFile, InStr(1, .lpstrFile, vbNullChar) - 1)
    End With
End Sub

Private Sub Form_Load()
    Dim strFile$
    ' Select resampler in the list
    cboResampler.ListIndex = BicubicBCSpline
    ' Load names of available images into ComboBox
    strFile = Dir(App.Path & "\Images\*.*")
    While Len(strFile)
        cboSource.AddItem strFile
        strFile = Dir
    Wend
    Caption = Tag
    Show
    frmPreview.Move Left + Width, Top
    frmPreview.Show vbModeless, Me
    SetFocus
End Sub

Private Sub hsc1_Change()
    hsc1_Scroll
End Sub

Private Sub hsc1_Scroll()
    sinc_size = hsc1.Value
    lbl1.Caption = "Size: " & sinc_size
    lbl1.Refresh
    frmPreview.Paint
End Sub

Private Sub hsc2_Change()
    hsc2_Scroll
End Sub

Private Sub hsc2_Scroll()
    param_d = hsc2.Value
    lbl2.Caption = "d = " & param_d
    lbl2.Refresh
    frmPreview.Paint
End Sub

Private Sub picBC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBC_MouseMove Button, Shift, X, Y
End Sub

Private Sub picBC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        picBC.Cls
        ' Determine max and min values by picture and resampler
        If X < 16 Then X = 16
        If Y < 9 Then Y = 9
        If X > 129 Then X = 129
        If Y > 122 Then Y = 122
        Select Case cboResampler.ListIndex
            Case BicubicBSpline
                X = 16: Y = 9
            Case BicubicCardinal
                Y = 122
        End Select
        ' Draw circle around the point
        picBC.Circle (X, Y), 3, vbRed
        ' Set a, B, C and show them
        cubic_c = (X - 16) / 113
        cubic_a = -cubic_c
        lbl3.Caption = "a = " & Round(cubic_a, 2)
        lbl2.Caption = "C = " & Round(cubic_c, 2)
        cubic_b = 1 - (Y - 9) / 113
        lbl1.Caption = "B = " & Round(cubic_b, 2)
        Refresh
        frmPreview.Paint
    End If
End Sub

Private Sub picBC_Paint()
    picBC_MouseMove vbLeftButton, 0, cubic_c * 113 + 16, (1 - cubic_b) * 113 + 9
End Sub
