VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Geometry Calculator"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Exit Geometry Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   79
      Top             =   6240
      Width           =   8295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Welcome"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label48"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label49"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Area"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Volume"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label42"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label43"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label44"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label45"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label46"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label47"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.Frame Frame6 
         Caption         =   "Cube"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -71160
         TabIndex        =   60
         Top             =   480
         Width           =   3135
         Begin VB.TextBox CubeAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   70
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox CubeHeight 
            Height          =   285
            Left            =   840
            TabIndex        =   66
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox CubeWidth 
            Height          =   285
            Left            =   840
            TabIndex        =   64
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox CubeLength 
            Height          =   285
            Left            =   840
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label41 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "Formula: (Length x Width) x height"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label39 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label38 
            Caption         =   "Width"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label37 
            Caption         =   "Length"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cylinder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   3255
         Begin VB.CommandButton Command5 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox CylAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   58
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox CylHeight 
            Height          =   285
            Left            =   840
            TabIndex        =   52
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox CylRadius 
            Height          =   285
            Left            =   840
            TabIndex        =   51
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "Formula: (R² x Height) x Pi"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label34 
            Caption         =   "Height of cylinder"
            Height          =   255
            Left            =   1800
            TabIndex        =   55
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label33 
            Caption         =   "or Diameter ÷ 2"
            Height          =   255
            Left            =   1800
            TabIndex        =   54
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "Radius"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   375
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Circle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -71160
         TabIndex        =   40
         Top             =   3240
         Width           =   4335
         Begin VB.CommandButton Command4 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox circAnswer 
            Height          =   285
            Left            =   720
            TabIndex        =   46
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox CircRadius 
            Height          =   285
            Left            =   720
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label30 
            Caption         =   "Formula: R² x Pi"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   2400
            Width           =   3975
         End
         Begin VB.Label Label29 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Radius = Diameter Divided by 2"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label27 
            Caption         =   "Diameter ÷ 2 (if radius isnt given)"
            Height          =   255
            Left            =   1800
            TabIndex        =   43
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label26 
            Caption         =   "Radius"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Triangle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   29
         Top             =   3240
         Width           =   3495
         Begin VB.CommandButton Command3 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox TriAnswer 
            Height          =   285
            Left            =   720
            TabIndex        =   38
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox TriAltitude 
            Height          =   285
            Left            =   720
            TabIndex        =   33
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TriBase 
            Height          =   285
            Left            =   720
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Formula: (Base x Altitude) x 1/2"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label23 
            Caption         =   "Heigth of the triangle"
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Base Length"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Altitude"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Base"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rectangle / Square"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -71160
         TabIndex        =   18
         Top             =   480
         Width           =   4335
         Begin VB.CommandButton Command2 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox RectAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox RectBase2 
            Height          =   285
            Left            =   840
            TabIndex        =   22
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox RectBase1 
            Height          =   285
            Left            =   840
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Base2 = Length of another side"
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Base1 = Length of one side"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Formula: Base1 * Base2"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Base 2"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Base 1"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Trapazoid / Parallellagram"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   3495
         Begin VB.CommandButton Command1 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox TrapAnswer 
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox TrapAltitude 
            Height          =   285
            Left            =   840
            TabIndex        =   9
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox TrapBase2 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TrapBase1 
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Altitude = height"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Base2 = another side"
            Height          =   255
            Left            =   1800
            TabIndex        =   16
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Base1 = one side"
            Height          =   255
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Formula: [ ( Base1 + Base2 ) x Altitude ] x ½"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label8 
            Caption         =   "Altitude"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1095
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Base 2"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Base 1"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   375
            Width           =   615
         End
      End
      Begin VB.Label Label49 
         Caption         =   "Do NOT enter the Metric Mesurment: ft. in. cm, mm etc., please add that yourself"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   78
         Top             =   4800
         Width           =   7455
      End
      Begin VB.Label Label48 
         Caption         =   "Note: All answers are in Square format. Like Sq. Ft., Square Inch, and so on."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   77
         Top             =   4440
         Width           =   6855
      End
      Begin VB.Label Label47 
         Caption         =   "R = Symbol for Radius"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "³ = Exponent 3 or Cubed"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label45 
         Caption         =   "² = Exponent 2 or Squared."
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label44 
         Caption         =   "Pi = 3.141592654"
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Radius = From dead center of the circle to the outer edge. Or the Diameter ÷ 2"
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label42 
         Caption         =   "Base = Length of a side of the object"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Area claculator is good for those of us who need quick answers! I hate math, so thats why i made this!! :)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   2760
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Author: Rizky Khapidsyah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Geometry Calculator v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi = 3.141592654

Private Sub Command1_Click()
'calculate area of Trapazoid, Parallelagram
Dim Base1 As Double
Dim Base2 As Double
Dim Altitude As Double
Dim Answer As Double
Dim temp As Double
Base1 = TrapBase1.Text
Base2 = TrapBase2.Text
Altitude = TrapAltitude.Text
temp = Base1 + Base2
temp = temp * Altitude
temp = temp * 0.5
Answer = temp
TrapAnswer.Text = Answer & " ²"
End Sub

Private Sub Command2_Click()
'calculate area of rectangle
Dim Base1 As Double
Dim Base2 As Double
Dim temp As Double
Dim Answer As Double
Base1 = RectBase1.Text
Base2 = RectBase2.Text
temp = Base1 * Base2
Answer = temp
RectAnswer = Answer
End Sub

Private Sub Command3_Click()
'calculate area of triangle
Dim Base As Double
Dim Altitude As Double
Dim Answer As Double
Dim temp As Double
Base = TriBase.Text
Altitude = TriAltitude.Text
temp = Base * Altitude
temp = temp * 0.5
Answer = temp
TriAnswer = Answer & " ²"
End Sub

Private Sub Command4_Click()
'calculate area of a circle
Dim Radius As Double
Dim Answer As Double
Dim temp As Double
Radius = CircRadius.Text
temp = Radius * Radius
temp = temp * Pi
Answer = temp
circAnswer.Text = Answer & " ²"
End Sub

Private Sub Command5_Click()
'calculate volume of a cylinder
Dim Radius As Double
Dim Height As Double
Dim temp As Double
Dim Answer As Double
Radius = CylRadius.Text
Height = CylHeight.Text
temp = Radius * Radius
temp = temp * Height
temp = temp * Pi
Answer = temp
CylAnswer.Text = Answer & " ³"
End Sub

Private Sub Command6_Click()
'find the volume of a cube
Dim length As Double
Dim Width As Double
Dim cHeight As Double
Dim temp As Double
Dim Answer As Double
length = CubeLength.Text
Width = CubeLength.Text
cHeight = CubeHeight.Text
temp = length * Width
temp = temp * cHeight
Answer = temp
CubeAnswer.Text = Answer & " ³"
End Sub

