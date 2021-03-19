VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Fixation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2625
      ScaleWidth      =   4665
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Timer Timer_fixation 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1920
         Top             =   1200
      End
      Begin VB.Timer itt 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.PictureBox Blanco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   11640
      ScaleHeight     =   1185
      ScaleWidth      =   1185
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Pic_L 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1080
      ScaleHeight     =   2385
      ScaleWidth      =   2265
      TabIndex        =   24
      Top             =   5880
      Width           =   2295
      Begin VB.PictureBox Componente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   37
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Componente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   1080
         Width           =   255
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmp1 
         Height          =   615
         Left            =   1200
         TabIndex        =   27
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   873
         _cy             =   1085
      End
   End
   Begin VB.PictureBox Pic_R 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   3600
      ScaleHeight     =   2385
      ScaleWidth      =   2385
      TabIndex        =   25
      Top             =   5760
      Width           =   2415
      Begin VB.PictureBox Componente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Componente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   1440
         Width           =   255
      End
   End
   Begin VB.PictureBox Instrucciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10680
      ScaleHeight     =   1545
      ScaleWidth      =   3105
      TabIndex        =   29
      Top             =   720
      Width           =   3135
      Begin VB.CommandButton Empezar 
         Caption         =   "Empezar"
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox AT2_I 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2625
      ScaleWidth      =   3225
      TabIndex        =   28
      Top             =   840
      Width           =   3255
      Begin VB.Image Comp_Test10 
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   600
      End
      Begin VB.Image Comp_Test9 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   300
      End
      Begin VB.Image Comp_Test6 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   300
      End
      Begin VB.Image Comp_Test8 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   300
      End
      Begin VB.Image Comp_Test4 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   300
      End
      Begin VB.Image Comp_Test7 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   300
      End
      Begin VB.Image Comp_Test3 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   300
      End
      Begin VB.Image Comp_Test5 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   300
      End
      Begin VB.Image Comp_Test2 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   300
      End
      Begin VB.Image Comp_Test9 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image Comp_Test8 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image Comp_Test10 
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Index           =   0
         Left            =   960
         Stretch         =   -1  'True
         Top             =   960
         Width           =   600
      End
      Begin VB.Image Comp_Test6 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   0
         Width           =   300
      End
      Begin VB.Image Comp_Test7 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   300
      End
      Begin VB.Image Comp_Test5 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Comp_Test1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   720
         Width           =   255
      End
      Begin VB.Image Comp_Test1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   720
         Width           =   255
      End
      Begin VB.Image Comp_Test1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Comp_Test1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Comp_Test2 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   300
      End
      Begin VB.Image Comp_Test3 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   300
      End
      Begin VB.Image Comp_Test4 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.Timer Fixation_Test 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   4920
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   8040
      TabIndex        =   13
      Top             =   4800
      Width           =   5895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   360
         ScaleHeight     =   2535
         ScaleWidth      =   9375
         TabIndex        =   14
         Top             =   0
         Width           =   9375
         Begin VB.CommandButton Command3 
            Caption         =   "No acepto / Salir"
            Height          =   375
            Left            =   3360
            TabIndex        =   23
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Al hacer click ACEPTO y ENTIENDO el presente consentimiento informado"
            Height          =   1215
            Left            =   840
            TabIndex        =   22
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "año"
            Height          =   255
            Left            =   3120
            TabIndex        =   20
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "mes"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "dias"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -120
            TabIndex        =   18
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "de donde"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "numero de cedula"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del paciente"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ingrese sus datos, por favor"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Iniciar"
         Height          =   495
         Left            =   1440
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Experimento.frx":0000
         Left            =   2040
         List            =   "Experimento.frx":000A
         TabIndex        =   9
         Text            =   "Femenino"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Experimento.frx":0023
         Left            =   3720
         List            =   "Experimento.frx":0025
         TabIndex        =   6
         Text            =   "Año"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Experimento.frx":0027
         Left            =   2640
         List            =   "Experimento.frx":004F
         TabIndex        =   5
         Text            =   "Enero"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Experimento.frx":00B8
         Left            =   2040
         List            =   "Experimento.frx":0119
         TabIndex        =   4
         Text            =   "Dia"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "De dónde es la cédula:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Género:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Numero de cédula:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fechad de nacimiento:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nombre completo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Usr_Name As String, Usr_Age As Integer, Usr_Gender As String, Usr_Id As String, Usr_IdCity As String, Usr_Cod As String, Fase As Integer
Public Created_ExlRegEve As Boolean
Public Evento As String, Area As String
Public IndxReg As Integer
Public REExcel As Object, REBook As Object, RESheet As Object, REruta As String
Public Pos As Integer
Public Clicks As Integer
Public Aciertos As Integer
Public void As String, vertical As String, horizontal As String, diagonal2 As String, diagonal1 As String, vertical_diagonal1 As String, horizontal_diagonal2 As String
Public CntPosI As Integer, CntPosD As Integer
Public Click_Counter As Integer
Dim Pos_4(3) As String
Dim Component_Pos(3) As String
Dim Component_2Pos(1) As String
Dim EntrenamientoF2 As Boolean
Dim Instancia As String, Participante As String, Duracion As String
Dim Reg_CompActual As Integer, YaReg_Fond As Boolean, Noharegistrado As Boolean, NotRegistered As Boolean, CrearExcel As Boolean
Dim Comp_Test1_Pos(3) As String
Dim Reversion As Integer, Rnd1to8 As Integer


Sub Oculta_Covers()
    Dim i As Integer
        While i < Pic_Cover.Count
            Pic_Cover(i).BackColor = &HFFFFFF
            Componente(i).Visible = True
            i = i + 1
        Wend
    i = 0
End Sub

Sub Muestra_Covers()
    Dim i As Integer
        While i < Pic_Cover.Count
            Pic_Cover(i).BackColor = &H808080
            Componente(i).Visible = False
            i = i + 1
        Wend
    i = 0
End Sub

Sub Randomizar()
    Dim Random As Integer
    Random = Right(Format(Time, "hh:mm:ss"), 1)
    Random = (Random) + (Left(Right(Format(Time, "hh:mm:ss"), 2), 1))
        If Random Mod 2 = 0 Then
            Random = 2
        Else
            Random = 1
        End If
        Pos = Random
        If Pos = 1 Then
            CntPosD = CntPosD + 1       '1=der 2=izq ?
        Else
            CntPosI = CntPosI + 1
        End If
        If CntPosI = 3 Then
            CntPosI = 0
            If Pos = 1 Then
                Pos = 2
            Else
                Pos = 1
            End If
        End If
        If CntPosD = 3 Then
            CntPosD = 0
            If Pos = 1 Then
                Pos = 2
            Else
                Pos = 1
            End If
        End If
End Sub

Sub Rnd4()
    Dim Listo As Boolean
    Rnd1to8 = 0
    While Listo = False
        If Rnd1to8 > 8 Or Rnd1to8 < 1 Then
            Rnd1to8 = Int(Right(Format(Time, "hh:mm:ss"), 1))
            Rnd1to8 = Int(Right(Rnd1to8 - ((8 * Rnd) + 1), 1))
        Else
            Listo = True
        End If
    Wend
End Sub

Sub Mostrar()
    Select Case Rnd1to8
    Case 1
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "diagonal2"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "vertical"
    
        Componente(2).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "diagonal1"
    
        Componente(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "horizontal"
    Case 2
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "vertical"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "diagonal2"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "diagonal1"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "horizontal"
    Case 3
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "diagonal2"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "vertical"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "horizontal"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "diagonal1"
    Case 4
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "vertical"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "diagonal2"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "horizontal"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "diagonal1"
    Case 5
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "diagonal1"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "horizontal"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "diagonal2"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "vertical"
    Case 6
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "horizontal"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "diagonal1"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "diagonal2"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "vertical"
    Case 7
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "diagonal1"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "horizontal"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "vertical"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "diagonal2"
    Case 8
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "horizontal"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "diagonal1"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "vertical"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "diagonal2"
    End Select
End Sub

Sub Mostrar_Test1()
    Select Case Click_Counter
    Case 0
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(0) = "horizontal"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(1) = "diagonal2"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(2) = "vertical"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(3) = "diagonal1"
    Case 1
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(0) = "diagonal2"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(1) = "horizontal"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(2) = "diagonal1"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(3) = "vertical"
    Case 2
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(0) = "vertical"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(1) = "horizontal"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(2) = "diagonal2"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(3) = "diagonal1"
   Case 3
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(0) = "diagonal1"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(1) = "horizontal"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(2) = "vertical"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(3) = "diagonal2"
   Case 4
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(0) = "horizontal"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(1) = "vertical"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(2) = "diagonal2"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(3) = "diagonal1"
   Case 5
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(0) = "diagonal2"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(1) = "vertical"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(2) = "diagonal1"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(3) = "horizontal"
   Case 6
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(0) = "vertical"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(1) = "diagonal2"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(2) = "horizontal"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(3) = "diagonal1"
   Case 7
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(0) = "diagonal1"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(1) = "diagonal2"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(2) = "vertical"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(3) = "horizontal"
   Case 8
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(0) = "horizontal"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(1) = "diagonal1"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(2) = "diagonal2"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(3) = "vertical"
   Case 9
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(0) = "diagonal2"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(1) = "diagonal1"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(2) = "vertical"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(3) = "horizontal"
   Case 10
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(0) = "vertical"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(1) = "diagonal1"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(2) = "horizontal"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(3) = "diagonal2"
   Case 11
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test1_Pos(0) = "diagonal1"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test1_Pos(1) = "vertical"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test1_Pos(2) = "diagonal2"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test1_Pos(3) = "horizontal"
    End Select
    Click_Counter = Click_Counter + 1
End Sub

Sub Mostrar_Test2()
    If Pos = 1 Then
        Comp_Test2(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test2(0).Visible = True
        Component_2Pos(0) = "vertical"

        Comp_Test2(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test2(1).Visible = True
        Component_2Pos(1) = "horizontal"
    Else
        Comp_Test2(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test2(1).Visible = True
        Component_2Pos(1) = "vertical"

        Comp_Test2(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test2(0).Visible = True
        Component_2Pos(0) = "horizontal"
    End If
End Sub

Sub Mostrar_Test3()
    If Pos = 1 Then
        Comp_Test3(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test3(0).Visible = True
        Component_2Pos(0) = "vertical"

        Comp_Test3(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test3(1).Visible = True
        Component_2Pos(1) = "diagonal1"
    Else
        Comp_Test3(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test3(1).Visible = True
        Component_2Pos(1) = "vertical"

        Comp_Test3(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test3(0).Visible = True
        Component_2Pos(0) = "diagonal1"
    End If
End Sub

Sub Mostrar_Test4()
    If Pos = 1 Then
        Comp_Test4(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test4(0).Visible = True
        Component_2Pos(0) = "diagonal2"

        Comp_Test4(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test4(1).Visible = True
        Component_2Pos(1) = "horizontal"
    Else
        Comp_Test4(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test4(1).Visible = True
        Component_2Pos(1) = "diagonal2"

        Comp_Test4(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test4(0).Visible = True
        Component_2Pos(0) = "horizontal"
    End If
End Sub

Sub Mostrar_Test5()
    If Pos = 1 Then
        Comp_Test5(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test5(0).Visible = True
        Component_2Pos(0) = "diagonal2"

        Comp_Test5(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test5(1).Visible = True
        Component_2Pos(1) = "diagonal1"
    Else
        Comp_Test5(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test5(1).Visible = True
        Component_2Pos(1) = "diagonal2"

        Comp_Test5(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test5(0).Visible = True
        Component_2Pos(0) = "diagonal1"
    End If
End Sub

Sub Mostrar_Test6()
    If Pos = 1 Then
        Comp_Test6(0).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test6(0).Visible = True
        Component_2Pos(0) = "void"

        Comp_Test6(1).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test6(1).Visible = True
        Component_2Pos(1) = "horizontal"
    Else
        Comp_Test6(1).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test6(1).Visible = True
        Component_2Pos(1) = "void"

        Comp_Test6(0).Picture = LoadPicture(App.Path & "\data\img\horizontal.jpg")
        Comp_Test6(0).Visible = True
        Component_2Pos(0) = "horizontal"
    End If
End Sub

Sub Mostrar_Test7()
    If Pos = 1 Then
        Comp_Test7(0).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test7(0).Visible = True
        Component_2Pos(0) = "void"

        Comp_Test7(1).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test7(1).Visible = True
        Component_2Pos(1) = "vertical"
    Else
        Comp_Test7(1).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test7(1).Visible = True
        Component_2Pos(1) = "void"

        Comp_Test7(0).Picture = LoadPicture(App.Path & "\data\img\vertical.jpg")
        Comp_Test7(0).Visible = True
        Component_2Pos(0) = "vertical"
    End If
End Sub

Sub Mostrar_Test8()
    If Pos = 1 Then
        Comp_Test8(0).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test8(0).Visible = True
        Component_2Pos(0) = "void"

        Comp_Test8(1).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test8(1).Visible = True
        Component_2Pos(1) = "diagonal2"
    Else
        Comp_Test8(1).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test8(1).Visible = True
        Component_2Pos(1) = "void"

        Comp_Test8(0).Picture = LoadPicture(App.Path & "\data\img\diagonal2.jpg")
        Comp_Test8(0).Visible = True
        Component_2Pos(0) = "diagonal2"
    End If
End Sub
Sub Mostrar_Test9()
    If Pos = 1 Then
        Comp_Test9(0).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test9(0).Visible = True
        Component_2Pos(0) = "void"

        Comp_Test9(1).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test9(1).Visible = True
        Component_2Pos(1) = "diagonal1"
    Else
        Comp_Test9(1).Picture = LoadPicture(App.Path & "\data\img\void.jpg")
        Comp_Test9(1).Visible = True
        Component_2Pos(1) = "void"

        Comp_Test9(0).Picture = LoadPicture(App.Path & "\data\img\diagonal1.jpg")
        Comp_Test9(0).Visible = True
        Component_2Pos(0) = "diagonal1"
    End If
End Sub

Sub Mostrar_Test10()
    If Pos = 1 Then
        Comp_Test10(0).Picture = LoadPicture(App.Path & "\data\img\vertical_diagonal1.jpg")
        Comp_Test10(0).Visible = True
        Component_2Pos(0) = "vertical_diagonal1"

        Comp_Test10(1).Picture = LoadPicture(App.Path & "\data\img\horizontal_diagonal2.jpg")
        Comp_Test10(1).Visible = True
        Component_2Pos(1) = "horizontal_diagonal2"
    Else
        Comp_Test10(1).Picture = LoadPicture(App.Path & "\data\img\vertical_diagonal1.jpg")
        Comp_Test10(1).Visible = True
        Component_2Pos(1) = "vertical_diagonal1"

        Comp_Test10(0).Picture = LoadPicture(App.Path & "\data\img\horizontal_diagonal2.jpg")
        Comp_Test10(0).Visible = True
        Component_2Pos(0) = "horizontal_diagonal2"
    End If
End Sub
















Private Sub Componente_Click(Index As Integer)
    YaReg_Fond = False  'prueba
    Evento = "Click"
    Area = Component_Pos(Index)
    Reg_Event
    If EntrenamientoF2 = False Then
    Instancia = "EntrenamientoF1"
    Fase = 1
        If Component_Pos(Index) = "vertical" Or Component_Pos(Index) = "diagonal2" Then
        wmp1.URL = App.Path & "\Sounds\correct.wav"
        Aciertos = Aciertos + 1
    Else
        wmp1.URL = App.Path & "\Sounds\wrong.wav"
        Aciertos = 0
    End If
    If Aciertos = 11 Then
        Aciertos = 0
        Instancia = "Test 1"
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_1
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Rnd4
        Mostrar
        itt.Enabled = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
    End If
Else
    If Component_Pos(Index) = "horizontal" Or Component_Pos(Index) = "diagonal1" Then
        wmp1.URL = App.Path & "\Sounds\correct.wav"
        Aciertos = Aciertos + 1
    Else
        wmp1.URL = App.Path & "\Sounds\wrong.wav"
        Aciertos = 0
    End If
    If Aciertos = 11 Then
        Fixation.Visible = True
        Aciertos = 0
        Instancia = "Test 1"
        Click_Counter = 0
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_1
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Rnd4
        Mostrar
        itt.Enabled = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
    End If
End If
End Sub

Private Sub Comp_Test1_Click(Index As Integer)
    Evento = "Click"
    Area = Comp_Test1_Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 2"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_2
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        itt.Enabled = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Mostrar_Test1
    End If
End Sub

Private Sub Comp_test2_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 3"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_3
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test2
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test3_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 4"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_4
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test3
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test4_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 5"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_5
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test4
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test5_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 6"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_6
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test5
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test6_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 7"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_7
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test6
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test7_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 8"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_8
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test7
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test8_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 9"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_9
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test8
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test9_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        Instancia = "Test 10"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Test_10
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test9
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test10_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 5 Then
        If Reversion >= 1 Then
            Reg_Usr
            MsgBox "El experimento ha terminado, ¡muchas gracias por tu participación!"
            REExcel.Quit
            End
        End If
        Instancia = "EntrenamientoF2"
        Fase = 2
        Clicks = 0
        Reversion = Reversion + 1
        Fixation.Visible = True
        Hide_All
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        EntrenamientoF2 = True
        Fase1
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
        Clicks = Clicks + 1
        Randomizar
        Mostrar_Test10
        itt.Enabled = True
    End If
End Sub





Private Sub command2_click()
    'Command2.Visible = False
    'Command3.Visible = False
    Instrucciones.Visible = True
End Sub

Private Sub empezar_click() ' Este es el boton de acepto y entiendo el concentimiento
    Fixation.Visible = True
    Instancia = "EntrenamientoF1"
    Fase = 1
    Fixation_Test.Enabled = True    'prueba
    'Timer_fixation.Enabled = True
    Hide_All
    SetCursorPos ((Screen.Width / 15) / 2), ((Screen.Height / 15) / 2)
    Fase1
    Muestra_Covers  'prueba
End Sub

Private Sub Command3_Click() ' este es el boton de no acepto y salir
    End
End Sub

Private Sub Command1_Click() ' este es el boton de iniciar que esta en el formulario de registro
    'Calcula la edad del paciente
    Dim y As Integer
    Usr_Age = Right(Date, 4) - Combo3.Text
    'carga las variables con los datos
    Usr_Name = Text1.Text
    Usr_Id = Text2.Text
    Usr_IdCity = Text3.Text
    Usr_Gender = Combo4.Text
    'pone los datos en el formulario de consentimiento
    Label5.Caption = Usr_Name
    Label6.Caption = Usr_Id
    Label7.Caption = Usr_IdCity
    Label8.Caption = Left(Date, 2)
    Label9.Caption = Left(Right(Date, 7), 2)
    Label10.Caption = Right(Date, 4)
    'muestra el consentimiento informado
    Frame2.Visible = True
    'Command1.Visible = False
    Reg_Event
End Sub

Private Sub Form_Load()
    Iniciar
    'muesta el formulario de registro
    Frame1.Visible = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    REExcel.Quit
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Instancia = "EntrenamientoF1" Or Instancia = "EntrenamientoF2" Then
        YaReg_Fond = False  'prueba
        Muestra_Covers
        If Noharegistrado = False Then
            Evento = "Salió a "
            Area = "Fondo"
            Reg_Event
            Noharegistrado = True
        End If
    End If
End Sub

Private Sub AT2_I_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'YaReg_Fond = False
    Noharegistrado = False
        If NotRegistered = False Then
            Evento = "Salió a "
            Area = "Fondo"
            Reg_Event
            NotRegistered = True
        End If
End Sub

Private Sub Comp_Test1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Comp_Test1_Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Comp_Test10_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    NotRegistered = False
    If Noharegistrado = False Then
        Evento = "Miró"
        Area = Component_2Pos(Index)
        Reg_Event
        Noharegistrado = True
    End If
End Sub

Private Sub Blanco_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Muestra_Covers
End Sub

Private Sub Fixation_Test_Timer()
    Fixation_Test.Enabled = False
    Fixation.Visible = False
End Sub

Private Sub itt_Timer()
    Blanco.Visible = False
    itt.Enabled = False
End Sub

Private Sub Pic_Cover_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case (Index)
        Case 0
            Pic_Cover(0).BackColor = &HFFFFFF
            Pic_Cover(1).BackColor = &H808080
            Pic_Cover(2).BackColor = &H808080
            Pic_Cover(3).BackColor = &H808080
            Componente(0).Visible = True
            Componente(1).Visible = False
            Componente(2).Visible = False
            Componente(3).Visible = False
        Case 1
            Pic_Cover(1).BackColor = &HFFFFFF
            Pic_Cover(0).BackColor = &H808080
            Pic_Cover(2).BackColor = &H808080
            Pic_Cover(3).BackColor = &H808080
            Componente(0).Visible = False
            Componente(1).Visible = True
            Componente(2).Visible = False
            Componente(3).Visible = False
        Case 2
            Pic_Cover(2).BackColor = &HFFFFFF
            Pic_Cover(0).BackColor = &H808080
            Pic_Cover(1).BackColor = &H808080
            Pic_Cover(3).BackColor = &H808080
            Componente(0).Visible = False
            Componente(1).Visible = False
            Componente(2).Visible = True
            Componente(3).Visible = False
        Case 3
            Pic_Cover(3).BackColor = &HFFFFFF
            Pic_Cover(0).BackColor = &H808080
            Pic_Cover(1).BackColor = &H808080
            Pic_Cover(2).BackColor = &H808080
            Componente(0).Visible = False
            Componente(1).Visible = False
            Componente(2).Visible = False
            Componente(3).Visible = True
    End Select
    Noharegistrado = False
    If Instancia = "EntrenamientoF1" Or Instancia = "EntrenamientoF2" Then
        If YaReg_Fond = False Then
            Reg_CompActual = Index
            YaReg_Fond = True
            Evento = "Miró "
            Area = Component_Pos(Index)
            Reg_Event
        End If
        If Index <> Reg_CompActual Then
            YaReg_Fond = True
            Evento = "Miró "
            Area = Component_Pos(Index)
            Reg_CompActual = Index
            Reg_Event
        End If
    End If
End Sub

Private Sub Pic_L_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Muestra_Covers
End Sub

Private Sub Pic_R_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Muestra_Covers
End Sub

Public Sub Fase1()
    Rnd4
    Muestra_Covers
    Mostrar
    Pic_R.Visible = True
    Pic_L.Visible = True
    Pic_Cover(0).Visible = True
    Pic_Cover(1).Visible = True
    Pic_Cover(2).Visible = True
    Pic_Cover(3).Visible = True
End Sub

Public Sub Test_1()
    Mostrar_Test1
    AT2_I.Visible = True
    Comp_Test1(0).Visible = True
    Comp_Test1(1).Visible = True
    Comp_Test1(2).Visible = True
    Comp_Test1(3).Visible = True
End Sub

Public Sub Test_2()
    Randomizar
    Mostrar_Test2
    AT2_I.Visible = True
End Sub

Public Sub Test_3()
    Randomizar
    Mostrar_Test3
    AT2_I.Visible = True
End Sub

Public Sub Test_4()
    Randomizar
    Mostrar_Test4
    AT2_I.Visible = True
End Sub

Public Sub Test_5()
    Randomizar
    Mostrar_Test5
    AT2_I.Visible = True
End Sub

Public Sub Test_6()
    Randomizar
    Mostrar_Test6
    AT2_I.Visible = True
End Sub

Public Sub Test_7()
    Randomizar
    Mostrar_Test7
    AT2_I.Visible = True
End Sub

Public Sub Test_8()
    Randomizar
    Mostrar_Test8
    AT2_I.Visible = True
End Sub

Public Sub Test_9()
    Randomizar
    Mostrar_Test9
    AT2_I.Visible = True
End Sub

Public Sub Test_10()
    Randomizar
    Mostrar_Test10
    AT2_I.Visible = True
End Sub

Public Sub Iniciar()
    Hide_All
    'carga las rutas de las imágenes
    vertical = (App.Path & "\data\img\vertical.jpg")
    horizontal = (App.Path & "\data\img\horizontal.jpg")
    diagonal2 = (App.Path & "\data\img\diagonal2.jpg")
    diagonal1 = (App.Path & "\data\img\diagonal1.jpg")
    void = (App.Path & "\data\img\void.jpg")
    vertical_diagonal1 = (App.Path & "\data\img\vertical_diagonal1.jpg")
    horizontal_diagonal2 = (App.Path & "\data\img\horizontal_diagonal2.jpg")

        'Acomoda la ventana y los formularios
    With Form1
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = 0
        .Top = 0
    End With
    With Frame1
        .Left = Form1.Width / 2 - Frame1.Width / 2
        .Top = Form1.Height / 2 - Frame1.Height / 2
    End With
    With Frame2
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = 0
        .Top = 0
    End With
    With Picture2
        .Height = 12000
        .Width = 19500
        .Left = Form1.Width / 2 - Picture2.Width / 2
        .Top = Form1.Height / 2 - Picture2.Height / 2
        .Picture = LoadPicture(App.Path & "\data\consentimiento.jpg")
    End With
     With Instrucciones
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = Form1.Width / 2 - Instrucciones.Width / 2
        .Top = Form1.Height / 2 - Instrucciones.Height / 2
        .Picture = LoadPicture(App.Path & "\data\Instrucciones.jpg")
    End With
    With Label5
        .Width = 5500
        .Left = 1000
        .Top = 1600
    End With
    With Label6
        .Width = 2500
        .Left = 12000
        .Top = 1600
    End With
    With Label7
        .Width = 2500
        .Left = 14050
        .Top = 1600
    End With
    With Label8
        .Width = 2500
        .Left = 9120
        .Top = 10700
    End With
    With Label9
        .Width = 2500
        .Left = 11090
        .Top = 10700
    End With
    With Label10
        .Width = 2500
        .Left = 13500
        .Top = 10700
    End With
    With Command2
        .Top = 9000
        .Left = Picture2.Width / 2 - Command2.Width / 2
    End With
    With Empezar
        .Top = 9000
        .Left = Instrucciones.Width / 2 - Empezar.Width / 2
    End With
    With Command3
        .Top = 200
        .Left = Picture2.Width - (Command3.Width + 200)
    End With
    With Blanco
        .Height = Screen.Height
        .Width = Screen.Width
        .Top = 0
        .Left = 0
    End With
    
    With Fixation
        .Height = Screen.Height
        .Width = Screen.Width
        .Top = 0
        .Left = 0
    End With
    
    'acomoda las áreas de presentación del entrenamiento donde se cargaran los estimulos
    With Pic_L
        .Height = (Screen.Height / 5) * 4
        .Width = Pic_L.Height / 2
        .Left = Screen.Width / 4 - Pic_L.Width / 2
        .Top = Screen.Height / 4 - (Screen.Height / 6)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Pic_R
        .Height = (Screen.Height / 5) * 4
        .Width = Pic_R.Height / 2
        .Left = ((Screen.Width / 4) * 3) - Pic_R.Width / 2
        .Top = Screen.Height / 4 - (Screen.Height / 6)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de las areas de presentación de los demás test
    With AT2_I
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = 0
        .Top = 0
    End With
    
    'posicionamiento de los Componentes y covers del entrenamiento
    With Componente(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(0).Width / 2)
        .Top = Pic_L.Height / 4 - (Componente(0).Height / 2)
    End With
    With Componente(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(1).Width / 2)
        .Top = (Pic_L.Height / 4 * 3) - (Componente(0).Height / 2)
    With Componente(2)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(0).Width / 2)
        .Top = Pic_L.Height / 4 - (Componente(0).Height / 2)
    End With
    With Componente(3)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_R.Width / 2 - (Componente(3).Width / 2)
        .Top = (Pic_L.Height / 4 * 3) - (Componente(0).Height / 2)
    End With
    With Pic_Cover(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(0).Width / 2)
        .Top = Pic_L.Height / 4 - (Componente(0).Height / 2)
    End With
    With Pic_Cover(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(1).Width / 2)
        .Top = (Pic_L.Height / 4 * 3) - (Componente(0).Height / 2)
        End With
    With Pic_Cover(2)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_L.Width / 2 - (Componente(0).Width / 2)
        .Top = Pic_L.Height / 4 - (Componente(0).Height / 2)
        End With
    With Pic_Cover(3)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Left = Pic_R.Width / 2 - (Componente(3).Width / 2)
        .Top = (Pic_L.Height / 4 * 3) - (Componente(0).Height / 2)
    End With
        
    'posicionamiento de los componentes del test 1
    With Comp_Test1(0)  'arriba izquierda
        .Height = Pic_L.Width
        .Width = Pic_L.Width
        .Top = Screen.Height / 4 - Comp_Test1(0).Height / 2
        .Left = Screen.Width / 4 - Comp_Test1(0).Width / 2
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test1(1)  'abajo izquierda
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 4) * 3 - (Comp_Test1(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test1(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test1(2)  'arriba derecha
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = Screen.Height / 4 - (Comp_Test1(0).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - Comp_Test1(0).Width / 2
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test1(3)  'abajo derecha
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 4) * 3 - (Comp_Test1(0).Height / 2)
        .Left = (Screen.Width / 4) * 3 - (Comp_Test1(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de las areas presentación y los componentes del test 2
    With Comp_Test2(0)
        .Height = Pic_L.Width
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test2(1)
        .Height = Pic_L.Width
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
        
    'posicionamiento de los componentes del test 3
    With Comp_Test3(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test3(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de los componentes del test 4
    With Comp_Test4(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test4(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de los componentes del test 5
    With Comp_Test5(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test5(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de los componentes del test 6
    With Comp_Test6(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test6(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de los componentes del test 7
    With Comp_Test7(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test7(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'posicionamiento de los componentes del test 8
    With Comp_Test8(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test8(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
     'posicionamiento de los componentes del test 9
    With Comp_Test9(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test9(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
     'posicionamiento de los componentes del test 10
    With Comp_Test10(0)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (Screen.Width / 4) - (Comp_Test2(0).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    With Comp_Test10(1)
        .Height = Pic_L.Height / 2
        .Width = Pic_L.Width
        .Top = (Screen.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = ((Screen.Width / 4) * 3) - (Comp_Test2(1).Width / 2)
        .Appearance = 1
        .BorderStyle = 1
    End With
    
    'carga los años del formulario
    Dim yi As Integer, yf As Integer, i As Integer
        yi = Right(Date, 4)
        yf = yi - 50
    While yf < yi
        Combo3.List(i) = yi
        i = i + 1
        yi = yi - 1
    Wend
        Combo3.Text = "Año"
        ' indica que no se ha creado el excel para registrar eventos
        Created_ExlRegEve = False
        IndxReg = 0
    End With
End Sub

Public Sub Hide_All()
    Blanco.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Pic_L.Visible = False
    Pic_R.Visible = False
    AT2_I.Visible = False
    Instrucciones.Visible = False
    Pic_Cover(0).Visible = False
    Pic_Cover(1).Visible = False
    Pic_Cover(2).Visible = False
    Pic_Cover(3).Visible = False
    Componente(0).Visible = False
    Componente(1).Visible = False
    Componente(2).Visible = False
    Componente(3).Visible = False
    Comp_Test1(0).Visible = False
    Comp_Test1(1).Visible = False
    Comp_Test1(2).Visible = False
    Comp_Test1(3).Visible = False
    Comp_Test2(0).Visible = False
    Comp_Test2(1).Visible = False
    Comp_Test3(0).Visible = False
    Comp_Test3(1).Visible = False
    Comp_Test4(0).Visible = False
    Comp_Test4(1).Visible = False
    Comp_Test5(0).Visible = False
    Comp_Test5(1).Visible = False
    Comp_Test6(0).Visible = False
    Comp_Test6(1).Visible = False
    Comp_Test7(0).Visible = False
    Comp_Test7(1).Visible = False
    Comp_Test8(0).Visible = False
    Comp_Test8(1).Visible = False
    Comp_Test9(0).Visible = False
    Comp_Test9(1).Visible = False
    Comp_Test10(0).Visible = False
    Comp_Test10(1).Visible = False
End Sub

Public Sub Reg_Usr()
    ' registra el usuario en la database
    Dim Excel As Object, ruta As String
    Dim LibroExcel As Object
    Dim HojaExcel As Object
    Dim UltimaDisponible As Integer
    UltimaDisponible = 1
    Set Excel = CreateObject("Excel.Application")
    Set LibroExcel = Excel.Workbooks
    Excel.Visible = False
    Excel.Application.DisplayAlerts = False
    ruta = App.Path + ("\data\participantes.xlsx")
    Set LibroExcel = Excel.Workbooks.Open(FileName:=ruta, ReadOnly:=False)
    Set HojaExcel = LibroExcel.Sheets(1)
    'encuentra ultima celda disponible
        While HojaExcel.range("A" & UltimaDisponible) <> ""
            UltimaDisponible = UltimaDisponible + 1
        Wend
        Usr_Cod = (UltimaDisponible - 1) & Left(Usr_Name, 1) & Left(Usr_Gender, 1) & Usr_Age
        ' escribe registro
        With HojaExcel
            .range("A" & UltimaDisponible) = Usr_Cod
            .range("B" & UltimaDisponible) = Usr_Name
            .range("C" & UltimaDisponible) = Usr_Age
            .range("D" & UltimaDisponible) = Usr_Gender
            .range("E" & UltimaDisponible) = Date & " - " & Time
            .range("F" & UltimaDisponible) = Usr_Id
        End With
        LibroExcel.Save
        Excel.Quit
End Sub

Public Sub Reg_Event()
    IndxReg = IndxReg + 1
    If CrearExcel = False Then
        If Created_ExlRegEve = False Then
            Set REExcel = CreateObject("Excel.Application")
            Set REBook = REExcel.Workbooks.Add
            Set RESheet = REBook.Worksheets(1)
            Usr_Cod = Usr_Name
            REruta = App.Path & "\data\" & Usr_Cod & ".xlsx"
            RESheet.range("A1").Value = "PARTICIPANTE"
            RESheet.range("B1").Value = "HORA"
            RESheet.range("C1").Value = "DURACIÓN"
            RESheet.range("D1").Value = "EVENTO"
            RESheet.range("E1").Value = "AREA"
            RESheet.range("F1").Value = "INSTANCIA"
            RESheet.range("G1").Value = "FASE"
            REBook.SaveAs REruta
            Created_ExlRegEve = True
        End If
            'Agrega registro
    CrearExcel = True
    Else
        With RESheet
            .range("A" & IndxReg).Value = Usr_Name
            .range("B" & IndxReg).Value = Format(Time, "hh:nn:ss") & "." & Right(Format(Timer, "#0.000"), 3)
            .range("C" & IndxReg).Value = Duracion
            .range("D" & IndxReg).Value = Evento
            .range("E" & IndxReg).Value = Area
            .range("F" & IndxReg).Value = Instancia
            .range("G" & IndxReg).Value = Fase
        End With
    End If
    'Guardar el excel y cerrarlo
    REBook.Save
End Sub
