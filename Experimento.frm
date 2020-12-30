VERSION 5.00
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
      Height          =   4575
      Left            =   1080
      ScaleHeight     =   4575
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Usr_Name As String, Usr_Age As Integer, Usr_Gender As String, Usr_Id As String, Usr_IdCity As String

Private Sub Form_Load()
Iniciar
Fase1
End Sub

Public Sub Fase1()
MsgBox "inicia face 1"
'carga de estimulos
Pic_Arriba_Izq.Picture = LoadPicture(App.Path & "\data\stim001.jpg")
End Sub


Public Sub Iniciar()
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

'oculta todo
Hide_All

'meusta el consentimiento informado
Frame1.Visible = True

'acomoda los pic donde se cargaran los estimulos



End Sub

Private Sub Command2_Click()
Reg_Usr
Hide_All
End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Command1_Click()
'Calcula la edad del paciente
Dim y As Integer
Usr_Age = Right(Date, 4) - Combo3.Text
'carga las variables con los datos
Usr_Name = Text1.Text
Usr_Id = Text2.Text
Usr_IdCity = Text3.Text
Usr_Gender = Combo4.Text
'pone los datos en el formulario
Label5.Caption = Usr_Name
Label6.Caption = Usr_Id
Label7.Caption = Usr_IdCity
Label8.Caption = Left(Date, 2)
Label9.Caption = Left(Right(Date, 7), 2)
Label10.Caption = Right(Date, 4)
Frame2.Visible = True
End Sub


Public Sub Hide_All()
Picture1.Visible = False
Frame1.Visible = False
Frame2.Visible = False
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
' escribe registro
With HojaExcel
    .range("A" & UltimaDisponible) = (UltimaDisponible - 1) & Left(Usr_Name, 1) & Left(Usr_Gender, 1) & Usr_Age
    .range("B" & UltimaDisponible) = Usr_Name
    .range("C" & UltimaDisponible) = Usr_Age
    .range("D" & UltimaDisponible) = Usr_Gender
    .range("E" & UltimaDisponible) = Date & " - " & Time
    .range("F" & UltimaDisponible) = Usr_Id
End With


'Aqui es donde ocurre la magia. tu pones hoja excel .range y entreparentesis la cadena de caracteres
'para jalar los datos del excel, puedes poner un rango o una celdita usando la misma nomenclatura que en excel
'Print HojaExcel.range(Text1.Text)

'esto tiene que ir, asi no lo veas el excel se abre, y si no lo pones quedaria abierto y tocaria cerrarlo desde
'el administrador de tareas con control alt suprimir.
LibroExcel.save
Excel.Quit



'ErrorHandler:
'
'    If MsgBox("la base de datos de participantes no existe desea crearla?", vbYesNo) = 6 Then
'        MsgBox "Listo"
'    Else
'        MsgBox "De acuerdo entonces no se registraran los participantes"
'    End If

End Sub


Public Sub DB_Gen()
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim ruta As String
    
'Crear nuevo archivo de excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'agregar datos a la primer hoja en el excel
Set oSheet = oBook.Worksheets(1)
oSheet.range("A1").Value = "Nombre:"
oSheet.range("B1").Value = "Nombre del paciente"
oSheet.range("C1").Value = "Fecha y hora:"
oSheet.range("D1").Value = (Date & " " & Time)

ruta = App.Path + "\data\Nombre_Paciente.xlsx"

'Guardar el excel y cerrarlo
oBook.SaveAs ruta
oExcel.Quit
End Sub
