VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Principale 
   Caption         =   "Pengubah nama untuk file Midi"
   ClientHeight    =   6180
   ClientLeft      =   2670
   ClientTop       =   2070
   ClientWidth     =   6645
   Icon            =   "Nom Midi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   6645
   Begin VB.CommandButton B_Changer 
      Caption         =   "Changer Nom"
      Height          =   396
      Left            =   5280
      TabIndex        =   7
      Top             =   5184
      Width           =   1212
   End
   Begin VB.TextBox T_NouveauNom 
      Height          =   348
      Left            =   48
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5184
      Width           =   5148
   End
   Begin VB.ListBox L_Proposition 
      Height          =   2400
      Left            =   48
      TabIndex        =   5
      Top             =   2496
      Width           =   6444
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   372
      Left            =   48
      TabIndex        =   4
      Top             =   5604
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   327682
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   5472
   End
   Begin MCI.MMControl MMControl1 
      Height          =   468
      Left            =   1932
      TabIndex        =   2
      Top             =   5568
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   820
      _Version        =   393216
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.ListBox L_Fic 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   6450
   End
   Begin VB.TextBox T_Rep 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6450
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5712
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label L_Propositions 
      Caption         =   "Propositions de nom :"
      Height          =   204
      Left            =   96
      TabIndex        =   8
      Top             =   2256
      Width           =   1644
   End
   Begin VB.Label L_TempsEcoule 
      Height          =   732
      Left            =   4728
      TabIndex        =   3
      Top             =   5244
      Width           =   1632
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Principale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFichiers() As String
Dim iNbFic As Integer
Dim sRep As String
Dim iMorceauEnCours As Integer
Dim bAuSuivant As Boolean

Private Sub B_Changer_Click()
    On Error GoTo Erreur
    If InStr(1, T_NouveauNom.Text, ".mid", vbTextCompare) = 0 Then
        T_NouveauNom = Trim(T_NouveauNom) & ".mid"
    End If
    'MMControl1.Command = "Close"
    Name L_Fic.List(L_Fic.ListIndex) As T_NouveauNom.Text
    'MMControl1.filename = T_NouveauNom
    'MMControl1.Command = "open"
    'MMControl1.Command = "play"
    L_Fic.List(L_Fic.ListIndex) = T_NouveauNom.Text
    Exit Sub
    
Erreur:
    Select Case Err.Number
        Case 52
            MsgBox "Ce nom de fichier contient des caractères non valides." & vbLf & "Veuillez modifier le nom du fichier.", vbExclamation, T_NouveauNom & " invalide !"
            Exit Sub
        Case 53
            MsgBox "Ce nom de fichier n'est pas valide." & vbLf & "Veuillez en trouver un autre.", vbExclamation, T_NouveauNom & " invalide !"
            Exit Sub
        Case 55
            MsgBox "Le parcours de tous les titres possibles n'est fini.", vbExclamation, "Modification trop précoce"
            Exit Sub
        Case 58
            MsgBox "Ce nom de fichier existe déjà." & vbLf & "Veuillez en trouver un autre.", vbExclamation, T_NouveauNom & " éxiste déjà !"
            Exit Sub
        Case Else
            On Error GoTo 0
            Resume
        End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error GoTo Annuler
    CommonDialog1.Filter = "*.mid|*.mid"
    CommonDialog1.ShowOpen
    sRep = sIsoleRep(CommonDialog1.FileName)
    T_Rep = sRep
    LitFichiers iNbFic, sFichiers, sRep
    ChDir sRep
    For i = 0 To iNbFic - 1
        L_Fic.AddItem sFichiers(i)
    Next
    Slider1.SelectRange = False
    Show
    Joue
Exit Sub

Annuler:
    If Err.Number = 32755 Then
        End
    Else
        On Error GoTo 0
        Resume
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        With T_Rep
            .Left = 0
            .Top = 0
            .Width = Width - 100
        End With
        With MMControl1
            .Left = (Width - .Width) / 2
            .Top = Height - .Height - 500
        End With
        With Slider1
            .Left = 0
            .Top = MMControl1.Top
            .Width = MMControl1.Left - .Left - 100
        End With
        With L_TempsEcoule
            .Left = MMControl1.Left + MMControl1.Width + 100
            .Top = MMControl1.Top
            .Width = Width - .Left - 100
        End With
        With T_NouveauNom
            .Left = 0
            .Top = MMControl1.Top - .Height - 100
            .Width = Width - 350 - B_Changer.Width
            '.Height = MMControl1.Height + 1000
        End With
        With B_Changer
            .Left = T_NouveauNom.Left + T_NouveauNom.Width + 100
            .Top = MMControl1.Top - .Height - 100
            '.Width = Width - 100
            '.Height = MMControl1.Height + 1000
        End With
        With L_Fic
            .Left = 0
            .Top = T_Rep.Top + T_Rep.Height + 50
            .Width = Width - 100
'            .Height = (L_TitreEnCours.Top - .Top) / 2 - 100
        End With
        With L_Proposition
            .Left = 0
'            .Top = L_Fic.Top + 50
            .Width = Width - 100
'            '.Height = (L_TitreEnCours.Top - .Top) / 2 - 100
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "Close"
    End
End Sub

Sub Joue()
    Dim i As Integer
    Dim sTitre As String
    
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "Sequencer"
    iMorceauEnCours = 0
    MMControl1.AutoEnable = False
    MMControl1.PlayEnabled = True
    MMControl1.StopEnabled = True
    MMControl1.PauseEnabled = True
    MMControl1.PrevEnabled = True
    MMControl1.NextEnabled = True
    Do
        T_NouveauNom.Text = sFichiers(iMorceauEnCours)
        MMControl1.FileName = sFichiers(iMorceauEnCours)
        L_Fic.ListIndex = iMorceauEnCours
        T_NouveauNom.Text = ChercheTitre(sFichiers(iMorceauEnCours))
        If sTitre <> "" Then T_NouveauNom.Text = sTitre
        MMControl1.Command = "Close"        ' Au cas où...
        MMControl1.Command = "Open"
        MMControl1.Command = "Play"
        Timer1.Enabled = True
        Slider1.Min = 0
        If MMControl1.Length > 0 Then Slider1.Max = MMControl1.Length
        Slider1.TickFrequency = MMControl1.Length / 10

        bAuSuivant = False
        Me.MousePointer = 11
        DoEvents
        ChercheTitres sFichiers(iMorceauEnCours), L_Proposition
        Me.MousePointer = 0
        Do
            DoEvents
        Loop Until bAuSuivant
        MMControl1.Command = "Close"
        Timer1.Enabled = False
    Loop
End Sub

Private Sub L_Fic_Click()
    Debug.Print "Sélection directe"
    bAuSuivant = True
    iMorceauEnCours = L_Fic.ListIndex
End Sub

Private Sub L_Proposition_Click()
    T_NouveauNom.Text = L_Proposition.List(L_Proposition.ListIndex)
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    Debug.Print "NotifyCode"; NotifyCode
    If NotifyCode = 1 Then
        If MMControl1.Position = MMControl1.Length Then
            bAuSuivant = True
            iMorceauEnCours = iMorceauEnCours + 1
            If iMorceauEnCours >= iNbFic Then
                MMControl1_StopClick False
            End If
        End If
    End If
End Sub

Private Sub MMControl1_NextClick(Cancel As Integer)
    Debug.Print "Suivant"
    bAuSuivant = True
    iMorceauEnCours = iMorceauEnCours + 1
    If iMorceauEnCours >= iNbFic Then
        iMorceauEnCours = 0
    End If
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    Debug.Print "Lecture"
    bAuSuivant = True
    MMControl1.Refresh
End Sub

Private Sub MMControl1_PrevClick(Cancel As Integer)
    Debug.Print "Précédent"
    bAuSuivant = True
    iMorceauEnCours = iMorceauEnCours - 1
    If iMorceauEnCours < 0 Then
        iMorceauEnCours = iNbFic - 1
    End If
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
    Debug.Print "Stop"
    bAuSuivant = True
    iMorceauEnCours = 0
    Unload Me
    End
End Sub

Private Sub Timer1_Timer()
    Static slPos As Long
    Static lDurée As Long      ' Durée en seconde du morceau

    If slPos = MMControl1.Position Then     ' La pause est actionnée
    ElseIf slPos > MMControl1.Position Then        ' Les touches "Suivant" ou "Précédent" sont actionnées
        lDurée = 1
    Else                                                                ' Ecoulement normal
        lDurée = lDurée + 1
        L_TempsEcoule.Caption = MMControl1.Position & " octets" & vbCr _
                & Int(MMControl1.Position / 7.7) & " s sur " & Int(MMControl1.Length / (MMControl1.Position / lDurée)) & " s"
        Slider1.Value = MMControl1.Position
    End If
    slPos = MMControl1.Position
End Sub

