VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   17325
   StartUpPosition =   3  'Windows Default
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
      Begin VB.Image componente 
         Height          =   255
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image componente 
         Height          =   255
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic_Abajo_Der 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2520
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox Pic_Arriba_Der 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2400
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   26
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Pic_Abajo_Izq 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   840
      ScaleHeight     =   705
      ScaleWidth      =   1065
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   600
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   720
      Width           =   1005
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
      Begin VB.Image componente 
         Height          =   255
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image componente 
         Height          =   255
         Index           =   2
         Left            =   480
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
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



Public Sub Fase1()
'carga de estimulos
Pic_R.Visible = True
Pic_L.Visible = True

componente(0).Picture = LoadPicture(App.Path & "\data\img\wt.jpg")
componente(0).Visible = True


Pic_Cover(0).Visible = True
Pic_Cover(1).Visible = True
Pic_Cover(2).Visible = True
Pic_Cover(3).Visible = True
End Sub














Private Sub Form_Load()
Iniciar
'meusta el formulario de registro
Frame1.Visible = True
End Sub


Public Sub Iniciar()
Hide_All
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


'acomoda los pic donde se cargaran los estimulos
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

'posicionamiento de los covers
With Pic_Cover(0)
    .Height = (Pic_L.Height / 2) - (Pic_L.Height / 8)
    .Width = Pic_Cover(0).Height
    .Left = Pic_L.Width / 2 - (Pic_Cover(0).Width / 2)
    .Top = 0
End With
With Pic_Cover(1)
    .Height = (Pic_L.Height / 2) - (Pic_L.Height / 8)
    .Width = Pic_Cover(1).Height
    .Left = Pic_L.Width / 2 - (Pic_Cover(1).Width / 2)
    .Top = Pic_L.Height - (Pic_Cover(1).Height)
End With
With Pic_Cover(2)
    .Height = (Pic_R.Height / 2) - (Pic_R.Height / 8)
    .Width = Pic_Cover(2).Height
    .Left = Pic_R.Width / 2 - (Pic_Cover(2).Width / 2)
    .Top = 0
End With
With Pic_Cover(3)
    .Height = (Pic_R.Height / 2) - (Pic_R.Height / 8)
    .Width = Pic_Cover(3).Height
    .Left = Pic_R.Width / 2 - (Pic_Cover(3).Width / 2)
    .Top = Pic_R.Height - (Pic_Cover(3).Height)
End With


'posicionamiento de los componentes
With componente(0)
    '.Height = (Pic_L.Height / 2) - (Pic_L.Height / 16)
    '.Width = componente(0).Height
    .Left = Pic_L.Width / 2 - (componente(0).Width / 2)
    .Top = 0
End With
With componente(1)
    '.Height = (Pic_L.Height / 2) - (Pic_L.Height / 16)
    '.Width = componente(1).Height
    .Left = Pic_L.Width / 2 - (componente(1).Width / 2)
    .Top = Pic_L.Height - (componente(1).Height)
End With
With componente(2)
    '.Height = (Pic_R.Height / 2) - (Pic_R.Height / 16)
    '.Width = componente(2).Height
    .Left = Pic_R.Width / 2 - (componente(2).Width / 2)
    .Top = 0
End With
With componente(3)
    '.Height = (Pic_R.Height / 2) - (Pic_R.Height / 16)
    '.Width = componente(3).Height
    .Left = Pic_R.Width / 2 - (componente(3).Width / 2)
    .Top = Pic_R.Height - (componente(3).Height)
End With


'carga los años del formulario
Dim yi As Integer, yf As Integer, i As Integer
yi = Right(Date, 4)
yf = yi - 90
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

Private Sub Command2_Click() ' Este es el boton de acepto y entiendo el concentimiento
'guarda los datos del usuario en un excel
Reg_Usr
'oculta todo
Hide_All
'inicia face 1 del experimento
Fase1
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


Public Sub Hide_All()
Pic_Arriba_Izq.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Pic_Arriba_Izq.Visible = False
Pic_Abajo_Izq.Visible = False
Pic_Arriba_Der.Visible = False
Pic_Abajo_Der.Visible = False
Pic_L.Visible = False
Pic_R.Visible = False
Pic_Cover(0).Visible = False
Pic_Cover(1).Visible = False
Pic_Cover(2).Visible = False
Pic_Cover(3).Visible = False
componente(0).Visible = False
componente(1).Visible = False
componente(2).Visible = False
componente(3).Visible = False
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
    REBook.SaveAs REruta
    Created_ExlRegEve = True
End If

'Agrega registro
With RESheet
    .range("A" & IndxReg).Value = Date
    .range("B" & IndxReg).Value = Time
    .range("C" & IndxReg).Value = Evento
    .range("D" & IndxReg).Value = Area
End With

'Guardar el excel y cerrarlo
REBook.Save
End Sub



Private Sub Form_Unload(Cancel As Integer)
REExcel.Quit
End Sub

Private Sub Pic_Cover_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic_Cover(Index).BackStyle = 0
If Area <> ("Cover " & Index) Then
    Evento = "Mouse over"
    Area = ("Cover " & Index)
    Reg_Event
End If
End Sub


Private Sub Pic_L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Oculta_Covers
If Area <> "Estimulo de la izquierda" Then
    Evento = "Mouse over"
    Area = "Estimulo de la izquierda"
    Reg_Event
End If
End Sub
Private Sub Pic_R_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Oculta_Covers
If Area <> "Estimulo de la derecha" Then
    Evento = "Mouse over"
    Area = "Estimulo de la derecha"
    Reg_Event
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Oculta_Covers
If Area <> "Fondo" Then
    Evento = "Mouse over"
    Area = "Fondo"
    Reg_Event
End If
End Sub

Sub Oculta_Covers()
Dim i As Integer
While i < Pic_Cover.Count
    Pic_Cover(i).BackStyle = 1
    i = i + 1
Wend
i = 0
End Sub



