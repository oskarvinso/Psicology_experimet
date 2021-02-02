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
   Begin VB.PictureBox Blanco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   11880
      ScaleHeight     =   1185
      ScaleWidth      =   1185
      TabIndex        =   35
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox Fixation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6480
      ScaleHeight     =   2625
      ScaleWidth      =   4665
      TabIndex        =   34
      Top             =   5640
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Timer itt 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   480
         Top             =   840
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   2280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   1080
         Y2              =   1440
      End
   End
   Begin VB.PictureBox AT2_D 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1080
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   38
      Top             =   6720
      Width           =   975
      Begin VB.Image Comp_Test6 
         Height          =   510
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   510
      End
      Begin VB.Image Comp_Test5 
         Height          =   255
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Comp_Test4 
         Height          =   255
         Index           =   1
         Left            =   720
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Comp_Test3 
         Height          =   255
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Comp_Test2 
         Height          =   255
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox AT2_I 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   37
      Top             =   6720
      Width           =   975
      Begin VB.Image Comp_Test6 
         Height          =   510
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   510
      End
      Begin VB.Image Comp_Test5 
         Height          =   255
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Comp_Test4 
         Height          =   255
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Comp_Test3 
         Height          =   255
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Comp_Test2 
         Height          =   255
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer Fixation_Test 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   7080
   End
   Begin VB.Timer Timer_fixation 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8880
      Top             =   6600
   End
   Begin VB.PictureBox Pic_L 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   2265
      TabIndex        =   28
      Top             =   3960
      Width           =   2295
      Begin WMPLibCtl.WindowsMediaPlayer wmp1 
         Height          =   615
         Left            =   1200
         TabIndex        =   36
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
      Begin VB.Image Componente 
         Height          =   255
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Componente 
         Height          =   255
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Pic_Abajo_Der 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2520
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
      Begin VB.Image Comp_Test1 
         Height          =   255
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic_Arriba_Der 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2400
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   26
      Top             =   840
      Width           =   1095
      Begin VB.Image Comp_Test1 
         Height          =   255
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic_Abajo_Izq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   840
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
      Begin VB.Image Comp_Test1 
         Height          =   255
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   11400
      TabIndex        =   14
      Top             =   1320
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
         TabIndex        =   15
         Top             =   120
         Width           =   9375
         Begin VB.CommandButton Command3 
            Caption         =   "No acepto / Salir"
            Height          =   375
            Left            =   3360
            TabIndex        =   24
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Al hacer click ACEPTO y ENTIENDO el presente consentimiento informado"
            Height          =   1215
            Left            =   840
            TabIndex        =   23
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "año"
            Height          =   255
            Left            =   3120
            TabIndex        =   21
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "mes"
            Height          =   255
            Left            =   2280
            TabIndex        =   20
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
            Left            =   480
            TabIndex        =   19
            Top             =   1920
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del paciente"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ingrese sus datos por favor"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Iniciar"
         Height          =   495
         Left            =   1440
         TabIndex        =   11
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Experimento.frx":0000
         Left            =   2040
         List            =   "Experimento.frx":000A
         TabIndex        =   10
         Text            =   "Femenino"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Experimento.frx":0023
         Left            =   3720
         List            =   "Experimento.frx":0025
         TabIndex        =   7
         Text            =   "Año"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Experimento.frx":0027
         Left            =   2640
         List            =   "Experimento.frx":004F
         TabIndex        =   6
         Text            =   "Mes"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Experimento.frx":00B8
         Left            =   2040
         List            =   "Experimento.frx":0119
         TabIndex        =   5
         Text            =   "Dia"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "De donde es la cedula:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sexo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Numero cedula:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fecha De nacimiento:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
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
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.PictureBox Pic_Arriba_Izq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   600
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   720
      Width           =   1005
      Begin VB.Image Comp_Test1 
         Height          =   255
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic_R 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2520
      ScaleHeight     =   2385
      ScaleWidth      =   2385
      TabIndex        =   29
      Top             =   3960
      Width           =   2415
      Begin VB.Image Componente 
         Height          =   255
         Index           =   2
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Componente 
         Height          =   255
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Pic_Cover 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Usr_Name As String, Usr_Age As Integer, Usr_Gender As String, Usr_Id As String, Usr_IdCity As String, Usr_Cod As String
Public Created_ExlRegEve As Boolean
Public Evento As String, Area As String
Public IndxReg As Integer
Public REExcel As Object, REBook As Object, RESheet As Object, REruta As String
Public Pos As Integer
Public Clicks As Integer
Public Aciertos As Integer
Public Triangulo As String, Cuadrado As String, Rojo As String, Verde As String, Triangulo_Verde As String, Cuadrado_Rojo As String
Public CntPosI As Integer, CntPosD As Integer
Public Click_Counter As Integer
Dim Pos_4(3) As String
Dim Component_Pos(3) As String
Dim Component_2Pos(1) As String
Dim EntrenamientoF2 As Boolean
Dim Instancia As String
Dim NoHaRegistrado As Boolean
Dim Comp_Test1_Pos(3) As String
Dim Reversion As Integer



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

Sub Mostrar()
    If Pos = 1 Then
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "triangulo"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "rojo"
    
        Componente(2).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "cuadrado"
    
        Componente(3).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "verde"
    Else
        Componente(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Componente(0).Visible = True
        Component_Pos(0) = "cuadrado"

        Componente(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Componente(1).Visible = True
        Component_Pos(1) = "verde"

        Componente(2).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Componente(2).Visible = True
        Component_Pos(2) = "triangulo"

        Componente(3).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Componente(3).Visible = True
        Component_Pos(3) = "rojo"
    End If
End Sub

Sub Mostrar_Test2()
    If Pos = 1 Then
        Comp_Test2(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test2(0).Visible = True
        Component_2Pos(0) = "triangulo"

        Comp_Test2(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test2(1).Visible = True
        Component_2Pos(1) = "cuadrado"
    Else
        Comp_Test2(1).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test2(1).Visible = True
        Component_2Pos(1) = "triangulo"

        Comp_Test2(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test2(0).Visible = True
        Component_2Pos(0) = "cuadrado"
    End If
End Sub

Sub Mostrar_Test3()
    If Pos = 1 Then
        Comp_Test3(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test3(0).Visible = True
        Component_2Pos(0) = "triangulo"

        Comp_Test3(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test3(1).Visible = True
        Component_2Pos(1) = "verde"
    Else
        Comp_Test3(1).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test3(1).Visible = True
        Component_2Pos(1) = "triangulo"

        Comp_Test3(0).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test3(0).Visible = True
        Component_2Pos(0) = "verde"
    End If
End Sub

Sub Mostrar_Test4()
    If Pos = 1 Then
        Comp_Test4(0).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test4(0).Visible = True
        Component_2Pos(0) = "rojo"

        Comp_Test4(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test4(1).Visible = True
        Component_2Pos(1) = "cuadrado"
    Else
        Comp_Test4(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test4(1).Visible = True
        Component_2Pos(1) = "rojo"

        Comp_Test4(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test4(0).Visible = True
        Component_2Pos(0) = "cuadrado"
    End If
End Sub

Sub Mostrar_Test5()
    If Pos = 1 Then
        Comp_Test5(0).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test5(0).Visible = True
        Component_2Pos(0) = "rojo"

        Comp_Test5(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test5(1).Visible = True
        Component_2Pos(1) = "verde"
    Else
        Comp_Test5(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test5(1).Visible = True
        Component_2Pos(1) = "rojo"

        Comp_Test5(0).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test5(0).Visible = True
        Component_2Pos(0) = "verde"
    End If
End Sub

Sub Mostrar_Test6()
    If Pos = 1 Then
        Comp_Test6(0).Picture = LoadPicture(App.Path & "\data\img\triangulo_verde.jpg")
        Comp_Test6(0).Visible = True
        Component_2Pos(0) = "triangulo_verde"

        Comp_Test6(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado_rojo.jpg")
        Comp_Test6(1).Visible = True
        Component_2Pos(1) = "cuadrado_rojo"
    Else
        Comp_Test6(1).Picture = LoadPicture(App.Path & "\data\img\triangulo_verde.jpg")
        Comp_Test6(1).Visible = True
        Component_2Pos(1) = "triangulo_verde"

        Comp_Test6(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado_rojo.jpg")
        Comp_Test6(0).Visible = True
        Component_2Pos(0) = "cuadrado_rojo"
    End If
End Sub

Sub Mostrar_Test1()
    Select Case Click_Counter
    Case 0
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(0) = "cuadrado"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(1) = "rojo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(2) = "triangulo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(3) = "verde"
    Case 1
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(0) = "rojo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(1) = "cuadrado"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(2) = "verde"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(3) = "triangulo"
    Case 2
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(0) = "triangulo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(1) = "cuadrado"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(2) = "rojo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(3) = "verde"
   Case 3
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(0) = "verde"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(1) = "cuadrado"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(2) = "triangulo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(3) = "rojo"
   Case 4
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(0) = "cuadrado"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(1) = "triangulo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(2) = "rojo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(3) = "verde"
   Case 5
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(0) = "rojo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(1) = "triangulo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(2) = "verde"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(3) = "cuadrado"
   Case 6
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(0) = "triangulo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(1) = "rojo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(2) = "cuadrado"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(3) = "verde"
   Case 7
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(0) = "verde"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(1) = "rojo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(2) = "triangulo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(3) = "cuadrado"
   Case 8
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(0) = "cuadrado"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(1) = "verde"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(2) = "rojo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(3) = "triangulo"
   Case 9
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(0) = "rojo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(1) = "verde"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(2) = "triangulo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(3) = "cuadrado"
   Case 10
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(0) = "triangulo"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(1) = "verde"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(2) = "cuadrado"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(3) = "rojo"
   Case 11
        Comp_Test1(0).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
        Comp_Test1_Pos(0) = "verde"
        Comp_Test1(1).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
        Comp_Test1_Pos(1) = "triangulo"
        Comp_Test1(2).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
        Comp_Test1_Pos(2) = "rojo"
        Comp_Test1(3).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
        Comp_Test1_Pos(3) = "cuadrado"
    End Select
    Click_Counter = Click_Counter + 1
End Sub






Sub Randomizar_4()
    Dim ActualPos As Integer
    Dim Figura As String
    Dim Random As String
    Dim Ya_Trian As Boolean, Ya_Cuad As Boolean, Ya_Rojo As Boolean, Ya_Verde As Boolean
    ActualPos = 0

    While ActualPos < 4
        Random = Right(Format(Time, "hh:mm:ss"), 1)
        Random = Right(Random * Int(Right(Rnd, 1)), 1)
            If Random = 0 Or Random = 2 Then
                If Ya_Trian = False Then
                    Figura = "Triangulo"
                    Ya_Trian = True
                    Pos_4(ActualPos) = Figura
                    Comp_Test1(ActualPos).Picture = LoadPicture(App.Path & "\data\img\triangulo.jpg")
                    ActualPos = ActualPos + 1
                End If
            End If
                If Random = 1 Or Random = 3 Or Random = 8 Then
                    If Ya_Cuad = False Then
                    Figura = "Cuadrado"
                    Ya_Cuad = True
                    Pos_4(ActualPos) = Figura
                    Comp_Test1(ActualPos).Picture = LoadPicture(App.Path & "\data\img\cuadrado.jpg")
                    ActualPos = ActualPos + 1
                End If
            End If
                If Random = 4 Or Random = 6 Or Random = 9 Then
                If Ya_Rojo = False Then
                Figura = "Rojo"
                Ya_Rojo = True
                Pos_4(ActualPos) = Figura
                Comp_Test1(ActualPos).Picture = LoadPicture(App.Path & "\data\img\rojo.jpg")
                ActualPos = ActualPos + 1
            End If
        End If
            If Random = 5 Or Random = 7 Then
                If Ya_Verde = False Then
                    Figura = "Verde"
                    Ya_Verde = True
                    Pos_4(ActualPos) = Figura
                    Comp_Test1(ActualPos).Picture = LoadPicture(App.Path & "\data\img\verde.jpg")
                    ActualPos = ActualPos + 1
                End If
            End If
    Wend
End Sub







Private Sub Componente_Click(Index As Integer)
    Evento = "Click"
    Area = Component_Pos(Index)
    Reg_Event
If EntrenamientoF2 = False Then
    Instancia = "EntrenamientoF1"
    If Component_Pos(Index) = "triangulo" Or Component_Pos(Index) = "rojo" Then
        wmp1.URL = App.Path & "\Sounds\Correct.wav"
        Aciertos = Aciertos + 1
    Else
        wmp1.URL = App.Path & "\Sounds\Wrong.wav"
        Aciertos = 0
    End If
    If Aciertos = 12 Then           'lo puse abajo en vez de de primeras y parece haber mejorado el conteo de 12, antes eran 13.
        Aciertos = 0
        Instancia = "Test 1"
        Fixation.Visible = True
        Hide_All
        Test_1
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Randomizar
        Mostrar
        itt.Enabled = True
    End If
Else
    If Component_Pos(Index) = "cuadrado" Or Component_Pos(Index) = "verde" Then
        wmp1.URL = App.Path & "\Sounds\correct.wav"
        Aciertos = Aciertos + 1
    Else
        wmp1.URL = App.Path & "\Sounds\wrong.wav"
        Aciertos = 0
    End If
    If Aciertos = 12 Then           'lo puse abajo en vez de de primeras y parece haber mejorado el conteo de 12, antes eran 13.
        Fixation.Visible = True
        Aciertos = 0
        Instancia = "Test 1"
        Click_Counter = 0
        Hide_All
        Test_1
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Randomizar
        Mostrar
        itt.Enabled = True
    End If
End If
End Sub

Private Sub Comp_Test1_Click(Index As Integer)
    Evento = "Click"
    Area = Comp_Test1_Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        Instancia = "Test 2"
        Clicks = 0
        Fixation.Visible = True     'telón que se cierra
        Hide_All
        Test_2
        Fixation_Test.Enabled = True
    Else
        Blanco.Visible = True
        itt.Enabled = True
        Mostrar_Test1
        Clicks = Clicks + 1
    End If

End Sub

Private Sub Comp_test2_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        Instancia = "Test 3"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        Test_3
        Fixation_Test.Enabled = True
    Else
        Clicks = Clicks + 1
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Randomizar
        Mostrar_Test2
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test3_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        Instancia = "Test 4"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        Test_4
        Fixation_Test.Enabled = True
    Else
        Clicks = Clicks + 1
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Randomizar
        Mostrar_Test3
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test4_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        Instancia = "Test 5"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        Test_5
        Fixation_Test.Enabled = True
    Else
        Clicks = Clicks + 1
        Blanco.Visible = True       'oculta todo, suena el pito durante 0.5 segundos, e inicia el siguiente ensayo.
        Randomizar
        Mostrar_Test4
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test5_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        Instancia = "Test 6"
        Clicks = 0
        Fixation.Visible = True
        Hide_All
        Test_6
        Fixation_Test.Enabled = True
    Else
        Clicks = Clicks + 1
        Blanco.Visible = True
        Randomizar
        Mostrar_Test5
        itt.Enabled = True
    End If
End Sub

Private Sub Comp_Test6_Click(Index As Integer)
    Evento = "Click"
    Area = Component_2Pos(Index)
    Reg_Event
    If Clicks = 11 Then
        If Reversion >= 1 Then
            MsgBox "El experimento ha terminado, ¡muchas gracias por tu participación!"
            REExcel.Quit
            End
        End If
        Instancia = "EntrenamientoF2"
        Clicks = 0
        Reversion = Reversion + 1
        Fixation.Visible = True
        Hide_All
        EntrenamientoF2 = True
        Fase1
        Fixation_Test.Enabled = True
    Else
        Clicks = Clicks + 1
        Blanco.Visible = True
        Randomizar
        Mostrar_Test6
        itt.Enabled = True
    End If
End Sub

Private Sub Command2_Click() ' Este es el boton de acepto y entiendo el concentimiento
    Fixation.Visible = True
    'guarda los datos del usuario en un excel
    Reg_Usr
    Timer_fixation.Enabled = True
    'oculta todo
    Hide_All
    'inicia fase 1 del experimento
    Fase1
    Instancia = "EntrenamientoF1"
End Sub

Private Sub Command3_Click() ' este es el boton de no acepto y salir
    End
End Sub

Private Sub Command1_Click() ' este es el boton de iniciar que esta en el formulario de registro
    'Calcula la edad del paciente
    Dim Y As Integer
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
End Sub

Private Sub Form_Load()
    Iniciar
    'muesta el formulario de registro
    Frame1.Visible = True
End Sub

Private Sub Componente_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Instancia = "EntrenamientoF1" Then
    If NoHaRegistrado = False Then
        Evento = "Miró"
        Area = Component_Pos(Index)
        Reg_Event
        NoHaRegistrado = True
    End If
End If
End Sub

Private Sub Blanco_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Muestra_Covers
    If NoHaRegistrado = True Then
    Evento = "Salio a "
    Area = "Blanco"
    Reg_Event
    NoHaRegistrado = False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Muestra_Covers
    If NoHaRegistrado = True Then
    Evento = "Salio a "
    Area = "Fondo"
    Reg_Event
    NoHaRegistrado = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    REExcel.Quit
End Sub

Private Sub Fixation_Test_Timer()
    Fixation_Test.Enabled = False
    Fixation.Visible = False
End Sub

Private Sub itt_Timer()
    Blanco.Visible = False
    itt.Enabled = False
End Sub

Private Sub Pic_Cover_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pic_Cover(Index).BackColor = &HFFFFFF
    Componente(Index).Visible = True
End Sub

Private Sub Pic_L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Muestra_Covers
    If NoHaRegistrado = True Then
    Evento = "Salio a "
    Area = "Area izquierda"
    Reg_Event
    NoHaRegistrado = False
    End If
End Sub

Private Sub Pic_R_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Muestra_Covers
    If NoHaRegistrado = True Then
    Evento = "Salio a "
    Area = "Area derecha"
    Reg_Event
    NoHaRegistrado = False
    End If
    
End Sub

Private Sub Timer_fixation_Timer()
    Fixation.Visible = False
    Timer_fixation.Enabled = False
End Sub






Public Sub Fase1()
    Randomizar
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
    Pic_Arriba_Izq.Visible = True
    Pic_Arriba_Der.Visible = True
    Pic_Abajo_Izq.Visible = True
    Pic_Abajo_Der.Visible = True
    Comp_Test1(0).Visible = True
    Comp_Test1(1).Visible = True
    Comp_Test1(2).Visible = True
    Comp_Test1(3).Visible = True
End Sub

Public Sub Test_2()
    Randomizar
    Mostrar_Test2
    AT2_I.Visible = True
    AT2_D.Visible = True
End Sub

Public Sub Test_3()
    Randomizar
    Mostrar_Test3
    AT2_I.Visible = True
    AT2_D.Visible = True
End Sub

Public Sub Test_4()
    Randomizar
    Mostrar_Test4
    AT2_I.Visible = True
    AT2_D.Visible = True
End Sub

Public Sub Test_5()
    Randomizar
    Mostrar_Test5
    AT2_I.Visible = True
    AT2_D.Visible = True
End Sub

Public Sub Test_6()
    Randomizar
    Mostrar_Test6
    AT2_I.Visible = True
    AT2_D.Visible = True
End Sub

Public Sub Iniciar()
    Hide_All
    'carga las rutas de las imágenes
    Triangulo = (App.Path & "\data\img\triangulo.jpg")
    Cuadrado = (App.Path & "\data\img\cuadrado.jpg")
    Rojo = (App.Path & "\data\img\rojo.jpg")
    Verde = (App.Path & "\data\img\verde.jpg")
    Triangulo_Verde = (App.Path & "\data\img\triangulo_verde.jpg")
    Cuadrado_Rojo = (App.Path & "\data\img\cuadrado_rojo.jpg")

        'Acomoda la ventana y los formularios
    With Form1
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = 0
        .Top = 0
    End With
    With Frame1
        .Height = Screen.Height
        .Width = Screen.Width
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
        
    'ubica el punto de fijación después de crear el Excel y antes de iniciar el entrenamiento
    With Fixation
        .Height = Screen.Height
        .Width = Screen.Width
        .Top = 0
        .Left = 0
    End With
    With Line1
        .X1 = Fixation.Width / 2
        .Y1 = (Fixation.Height / 2) - 200
        .X2 = Fixation.Width / 2
        .Y2 = (Fixation.Height / 2) + 200
    End With
    With Line2
        .X1 = (Fixation.Width / 2) - 200
        .Y1 = Fixation.Height / 2
        .X2 = (Fixation.Width / 2) + 200
        .Y2 = Fixation.Height / 2
    End With
    
        'acomoda las áreas de presentación del entrenamiento donde se cargaran los estimulos
    With Pic_L
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_L.Height
        .Left = Screen.Width / 4 - Pic_L.Width / 2
        .Top = Screen.Height / 4 - Pic_L.Height / 2
    End With
    With Pic_R
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_R.Height
        .Left = ((Screen.Width / 4) * 3) - Pic_R.Width / 2
        .Top = Screen.Height / 4 - Pic_R.Height / 2
    End With
    
        'posicionamiento de los Componentes y covers del entrenamiento
    With Componente(0)
        .Left = Pic_L.Width / 2 - (Componente(0).Width / 2)
        .Top = 100
    End With
    With Componente(1)
        .Left = Pic_L.Width / 2 - (Componente(1).Width / 2)
        .Top = Pic_L.Height - (Componente(1).Height + 100)
    End With
    With Componente(2)
        .Left = Pic_R.Width / 2 - (Componente(2).Width / 2)
        .Top = 100
    End With
    With Componente(3)
        .Left = Pic_R.Width / 2 - (Componente(3).Width / 2)
        .Top = Pic_R.Height - (Componente(3).Height + 100)
    End With
    With Pic_Cover(0)
        .Left = Pic_L.Width / 2 - (Pic_Cover(0).Width / 2)
        .Top = 50
    End With
    With Pic_Cover(1)
        .Left = Pic_L.Width / 2 - (Pic_Cover(1).Width / 2)
        .Top = Pic_L.Height - (Pic_Cover(1).Height + 50)
        End With
    With Pic_Cover(2)
        .Left = Pic_R.Width / 2 - (Pic_Cover(0).Width / 2)
        .Top = 50
        End With
    With Pic_Cover(3)
        .Left = Pic_R.Width / 2 - (Pic_Cover(1).Width / 2)
        .Top = Pic_R.Height - (Pic_Cover(1).Height + 50)
    End With

        'posicionamiento areas de presentación Componentes test 1
    With Pic_Arriba_Izq
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_Arriba_Izq.Height
        .Left = Screen.Width / 4 - Pic_Arriba_Izq.Width / 2
        .Top = Screen.Height / 4 - Pic_Arriba_Izq.Height / 2
    End With
    With Pic_Arriba_Der
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_Arriba_Der.Height
        .Left = ((Screen.Width / 4) * 3) - Pic_Arriba_Der.Width / 2
        .Top = Screen.Height / 4 - Pic_Arriba_Der.Height / 2
    End With
        With Pic_Abajo_Izq
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_Abajo_Izq.Height
        .Left = Screen.Width / 4 - Pic_Abajo_Izq.Width / 2
        .Top = ((Screen.Height / 4) * 3) - Pic_Abajo_Izq.Height / 2
    End With
    With Pic_Abajo_Der
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = Pic_Abajo_Der.Height
        .Left = ((Screen.Width / 4) * 3) - Pic_Abajo_Der.Width / 2
        .Top = ((Screen.Height / 4) * 3) - Pic_Abajo_Der.Height / 2
    End With
    With Comp_Test1(0)
        .Top = (Pic_Arriba_Izq.Height / 2) - (Comp_Test1(0).Height / 2)
        .Left = (Pic_Arriba_Izq.Width / 2) - (Comp_Test1(0).Width / 2)
    End With
    With Comp_Test1(1)
        .Top = (Pic_Abajo_Izq.Height / 2) - (Comp_Test1(1).Height / 2)
        .Left = (Pic_Abajo_Izq.Width / 2) - (Comp_Test1(1).Width / 2)
    End With
    With Comp_Test1(2)
        .Top = (Pic_Arriba_Der.Height / 2) - (Comp_Test1(2).Height / 2)
        .Left = (Pic_Arriba_Der.Width / 2) - (Comp_Test1(2).Width / 2)
    End With
    With Comp_Test1(3)
        .Top = (Pic_Abajo_Der.Height / 2) - (Comp_Test1(3).Height / 2)
        .Left = (Pic_Abajo_Der.Width / 2) - (Comp_Test1(3).Width / 2)
    End With
    
    'posicionamiento areas presentación y componentes del test 2
    With AT2_I
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_I.Height
        .Left = Screen.Width / 4 - AT2_I.Width / 2
        .Top = Screen.Height / 4 - AT2_I.Height / 2
    End With
    With AT2_D
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_D.Height
        .Left = ((Screen.Width / 4) * 3) - AT2_D.Width / 2
        .Top = Screen.Height / 4 - AT2_D.Height / 2
    End With
    With Comp_Test2(0)
        .Top = (AT2_I.Height / 2) - (Comp_Test2(0).Height / 2)
        .Left = (AT2_I.Width / 2) - (Comp_Test2(0).Width / 2)
    End With
    With Comp_Test2(1)
        .Top = (AT2_D.Height / 2) - (Comp_Test2(1).Height / 2)
        .Left = (AT2_D.Width / 2) - (Comp_Test2(1).Width / 2)
    End With
        
         'posicionamiento areas presentación y componentes del test 3
    With AT2_I
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_I.Height
        .Left = Screen.Width / 4 - AT2_I.Width / 2
        .Top = Screen.Height / 4 - AT2_I.Height / 2
    End With
    With AT2_D
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_D.Height
        .Left = ((Screen.Width / 4) * 3) - AT2_D.Width / 2
        .Top = Screen.Height / 4 - AT2_D.Height / 2
    End With
    With Comp_Test3(0)
        .Top = (AT2_I.Height / 2) - (Comp_Test3(0).Height / 2)
        .Left = (AT2_I.Width / 2) - (Comp_Test3(0).Width / 2)
    End With
    With Comp_Test3(1)
        .Top = (AT2_D.Height / 2) - (Comp_Test3(1).Height / 2)
        .Left = (AT2_D.Width / 2) - (Comp_Test3(1).Width / 2)
    End With
    
        'posicionamiento areas presentación y componentes del test 4
    With AT2_I
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_I.Height
        .Left = Screen.Width / 4 - AT2_I.Width / 2
        .Top = Screen.Height / 4 - AT2_I.Height / 2
    End With
    With AT2_D
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_D.Height
        .Left = ((Screen.Width / 4) * 3) - AT2_D.Width / 2
        .Top = Screen.Height / 4 - AT2_D.Height / 2
    End With
    With Comp_Test4(0)
        .Top = (AT2_I.Height / 2) - (Comp_Test4(0).Height / 2)
        .Left = (AT2_I.Width / 2) - (Comp_Test4(0).Width / 2)
    End With
    With Comp_Test4(1)
        .Top = (AT2_D.Height / 2) - (Comp_Test4(1).Height / 2)
        .Left = (AT2_D.Width / 2) - (Comp_Test4(1).Width / 2)
    End With
    
        'posicionamiento areas presentación y componentes del test 5
    With AT2_I
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_I.Height
        .Left = Screen.Width / 4 - AT2_I.Width / 2
        .Top = Screen.Height / 4 - AT2_I.Height / 2
    End With
    With AT2_D
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_D.Height
        .Left = ((Screen.Width / 4) * 3) - AT2_D.Width / 2
        .Top = Screen.Height / 4 - AT2_D.Height / 2
    End With
    With Comp_Test5(0)
        .Top = (AT2_I.Height / 2) - (Comp_Test5(0).Height / 2)
        .Left = (AT2_I.Width / 2) - (Comp_Test5(0).Width / 2)
    End With
    With Comp_Test5(1)
        .Top = (AT2_D.Height / 2) - (Comp_Test5(1).Height / 2)
        .Left = (AT2_D.Width / 2) - (Comp_Test5(1).Width / 2)
    End With
    
        'posicionamiento areas presentación y componentes del test 5
    With AT2_I
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_I.Height
        .Left = Screen.Width / 4 - AT2_I.Width / 2
        .Top = Screen.Height / 4 - AT2_I.Height / 2
    End With
    With AT2_D
        .Height = (Screen.Height / 2) - (Screen.Height / 8)
        .Width = AT2_D.Height
        .Left = ((Screen.Width / 4) * 3) - AT2_D.Width / 2
        .Top = Screen.Height / 4 - AT2_D.Height / 2
    End With
    With Comp_Test6(0)
        .Top = (AT2_I.Height / 2) - (Comp_Test6(0).Height / 2)
        .Left = (AT2_I.Width / 2) - (Comp_Test6(0).Width / 2)
    End With
    With Comp_Test6(1)
        .Top = (AT2_D.Height / 2) - (Comp_Test6(1).Height / 2)
        .Left = (AT2_D.Width / 2) - (Comp_Test6(1).Width / 2)
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
    IndxReg = 1
End Sub

Public Sub Hide_All()
    Blanco.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Pic_L.Visible = False
    Pic_R.Visible = False
    AT2_I.Visible = False
    AT2_D.Visible = False
    Pic_Arriba_Izq.Visible = False
    Pic_Abajo_Izq.Visible = False
    Pic_Arriba_Der.Visible = False
    Pic_Abajo_Der.Visible = False
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
            'ErrorHandler:
            '
            '    If MsgBox("la base de datos de participantes no existe desea crearla?", vbYesNo) = 6 Then
            '        MsgBox "Listo"
            '    Else
            '        MsgBox "De acuerdo entonces no se registraran los participantes"
            '    End If
End Sub

Public Sub Reg_Event()
    IndxReg = IndxReg + 1

    If Created_ExlRegEve = False Then
        Set REExcel = CreateObject("Excel.Application")
        Set REBook = REExcel.Workbooks.Add
        Set RESheet = REBook.Worksheets(1)
        REruta = App.Path & "\data\" & Usr_Cod & ".xlsx"
        RESheet.range("A1").Value = "FECHA"
        RESheet.range("B1").Value = "HORA"
        RESheet.range("C1").Value = "EVENTO"
        RESheet.range("D1").Value = "AREA"
        RESheet.range("E1").Value = "INSTANCIA"
        REBook.SaveAs REruta
        Created_ExlRegEve = True
    End If

        'Agrega registro
    With RESheet
        .range("A" & IndxReg).Value = Date
        .range("B" & IndxReg).Value = Format(Time, "hh:nn:ss") & "." & Right(Format(Timer, "#0.000"), 3)
        .range("C" & IndxReg).Value = Evento
        .range("D" & IndxReg).Value = Area
        .range("E" & IndxReg).Value = Instancia
    End With

    'Guardar el excel y cerrarlo
    REBook.Save
End Sub
