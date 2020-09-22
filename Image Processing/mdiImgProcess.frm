VERSION 5.00
Begin VB.MDIForm mdiImgProcess 
   BackColor       =   &H8000000C&
   Caption         =   "Image Processing by Albert Nicholas"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8310
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Picture..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "F&ilters"
      Begin VB.Menu mnuFilterLight 
         Caption         =   "Lighten"
      End
      Begin VB.Menu mnuFilterDark 
         Caption         =   "Darken"
      End
      Begin VB.Menu mnuFilterGS 
         Caption         =   "Grayscale"
      End
      Begin VB.Menu mnuFilterInvert 
         Caption         =   "Invert Color"
      End
      Begin VB.Menu mnuFilterBlur 
         Caption         =   "Blur"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      Begin VB.Menu mnuWinCas 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWinHor 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuWinVer 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuWinArr 
         Caption         =   "Arrange Icons"
      End
   End
End
Attribute VB_Name = "mdiImgProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================
' Image Processing
' Author : Albert Nicholas
' Email  : nicho_tedja@yahoo.com
' ===================================

' The last project version takes a very long time to process an image
' Now, by using Windows API, this project enables you to process an image
' in a much shorter time.
'
' I apologize for those who have waited long for this projects
' You can use any concepts of this project, but PLEASE do not make a new project
' using these codes. Thank You

' Enclosed:
' Sample01.jpg
' Sample02.jpg
' just for sample pictures, in case you do not own any :)

Private Sub MDIForm_Load()
    indeks = 0
    currDir = "C:\My Documents"
End Sub

Private Sub mnuFileExit_Click()
    Dim i As Integer
    For i = 0 To indeks
        DeleteObject hBMPSour(i)
        DeleteDC hDCSour(i)
        DeleteObject hBMPDest(i)
        DeleteDC hDCDest(i)
            'THESE ARE IMPORTANT THINGS TO DO
            'Destroy all spaces and bitmaps to clean up memory
    Next i
    End
End Sub

Private Sub mnuFileOpen_Click()
    Load frmDlgOpen
    frmDlgOpen.Show 1
    If iCancel = True Then Exit Sub
    picforms(indeks).Show
    picforms(indeks).Caption = fpath
    picforms(indeks).Tag = indeks
    indeks = indeks + 1
End Sub

Private Sub mnuFilterBlur_Click()
    Call Blurring(ActiveForm.Tag)
End Sub

Private Sub mnuFilterDark_Click()
    Call Darken(ActiveForm.Tag)
End Sub

Private Sub mnuFilterGS_Click()
    Call Grayscaling(ActiveForm.Tag)
End Sub

Private Sub mnuFilterInvert_Click()
    Call Inverting(ActiveForm.Tag)
End Sub

Private Sub mnuFilterLight_Click()
    Call Lighten(ActiveForm.Tag)
End Sub

Private Sub mnuWinArr_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWinCas_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWinHor_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWinVer_Click()
    Me.Arrange vbTileVertical
End Sub
