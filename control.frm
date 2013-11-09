VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "DefconBot Automatic Control Panel"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Logging"
      Height          =   2175
      Left            =   240
      TabIndex        =   68
      Top             =   6960
      Width           =   7035
      Begin VB.TextBox txtCommLog 
         Height          =   1875
         Left            =   4200
         MultiLine       =   -1  'True
         TabIndex        =   70
         Text            =   "control.frx":0000
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtHistory 
         Height          =   1875
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   6915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Current Status"
      Height          =   1995
      Left            =   5760
      TabIndex        =   54
      Top             =   2580
      Width           =   5595
      Begin VB.Timer tmrClock 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3180
         Top             =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "(x,y):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblFiring 
         Caption         =   "FIRING FIRING FIRING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1695
         Left            =   3780
         TabIndex        =   60
         Top             =   180
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblPosition 
         BackStyle       =   0  'Transparent
         Caption         =   "0,0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   59
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblClock 
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1920
         TabIndex        =   58
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label lblNumTargets 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1920
         TabIndex        =   57
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Target:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label21 
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   55
         Top             =   1380
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Behavior"
      Height          =   2475
      Left            =   8400
      TabIndex        =   49
      Top             =   4740
      Width           =   2955
      Begin VB.Timer tmrFireSearch 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   2520
         Top             =   840
      End
      Begin VB.CommandButton cmdCameraOn 
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   660
         Width           =   555
      End
      Begin VB.CommandButton cmdCameraOff 
         Caption         =   "Off"
         Height          =   255
         Left            =   720
         TabIndex        =   65
         Top             =   660
         Width           =   555
      End
      Begin VB.Timer tmrNoTarget 
         Interval        =   30000
         Left            =   3000
         Top             =   1200
      End
      Begin VB.CheckBox chkEnableFire 
         Caption         =   "Enable Firing"
         Height          =   195
         Left            =   180
         TabIndex        =   64
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chkWebcamControl 
         Caption         =   "Enable Webcam Control"
         Height          =   195
         Left            =   180
         TabIndex        =   63
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkNoTarget 
         Caption         =   "Recenter every 30 seconds"
         Height          =   315
         Left            =   180
         TabIndex        =   62
         Top             =   1380
         Width           =   2475
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   2340
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton cmdCommOff 
         Caption         =   "Off"
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCommOn 
         Caption         =   "On"
         Height          =   255
         Left            =   1140
         TabIndex        =   51
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtCommPort 
         Height          =   285
         Left            =   660
         TabIndex        =   50
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Camera"
         Height          =   195
         Left            =   1380
         TabIndex        =   67
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Comm"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Camera"
      Height          =   4215
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   5655
      Begin VB.VScrollBar scrCenterSize 
         Height          =   2655
         LargeChange     =   5
         Left            =   5340
         Max             =   25
         TabIndex        =   71
         Top             =   660
         Value           =   15
         Width           =   195
      End
      Begin VB.Timer tmrGrab 
         Interval        =   500
         Left            =   4380
         Top             =   3300
      End
      Begin VB.HScrollBar scrGrabTimer 
         Height          =   255
         LargeChange     =   100
         Left            =   900
         TabIndex        =   42
         Top             =   3840
         Width           =   4095
      End
      Begin VB.VScrollBar scrTargetColor 
         Height          =   3135
         LargeChange     =   20
         Left            =   4980
         Max             =   255
         TabIndex        =   41
         Top             =   420
         Value           =   15
         Width           =   195
      End
      Begin VB.PictureBox pctPreview 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3555
         Left            =   60
         ScaleHeight     =   233
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   317
         TabIndex        =   39
         Top             =   240
         Width           =   4815
      End
      Begin VB.CheckBox chkTargetsShow 
         Caption         =   "Show"
         Height          =   255
         Left            =   4860
         TabIndex        =   80
         Top             =   3540
         Width           =   735
      End
      Begin VB.Label lblCenterSize 
         BackStyle       =   0  'Transparent
         Caption         =   "ms"
         Height          =   195
         Left            =   5220
         TabIndex        =   73
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Left            =   5280
         TabIndex        =   72
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Grab Timer"
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   3900
         Width           =   915
      End
      Begin VB.Label lblGrabTimer 
         Caption         =   "ms"
         Height          =   195
         Left            =   5040
         TabIndex        =   43
         Top             =   3900
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Color"
         Height          =   195
         Left            =   4920
         TabIndex        =   40
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame x 
      Caption         =   "Raw Information and Status"
      Height          =   2475
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   8055
      Begin VB.TextBox txtCameraOffsetX 
         Height          =   285
         Left            =   2100
         TabIndex        =   76
         Text            =   "0"
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtCameraOffsetY 
         Height          =   285
         Left            =   2940
         TabIndex        =   75
         Text            =   "0"
         Top             =   2100
         Width           =   615
      End
      Begin VB.CommandButton cmdSetCameraOffset 
         Caption         =   "Set"
         Height          =   255
         Left            =   3600
         TabIndex        =   74
         Top             =   2100
         Width           =   495
      End
      Begin VB.Timer tmrGUI 
         Interval        =   500
         Left            =   900
         Top             =   1020
      End
      Begin VB.HScrollBar ScrRawFire 
         Height          =   315
         LargeChange     =   100
         Left            =   1680
         Max             =   5500
         Min             =   500
         TabIndex        =   37
         Top             =   1020
         Value           =   3000
         Width           =   5895
      End
      Begin VB.HScrollBar ScrRawTilt 
         Height          =   315
         LargeChange     =   100
         Left            =   1680
         Max             =   5500
         Min             =   500
         TabIndex        =   36
         Top             =   660
         Value           =   3000
         Width           =   5895
      End
      Begin VB.HScrollBar ScrRawPan 
         Height          =   315
         LargeChange     =   100
         Left            =   1680
         Max             =   5500
         Min             =   500
         TabIndex        =   35
         Top             =   300
         Value           =   3000
         Width           =   5895
      End
      Begin VB.CommandButton cmdSetLimitFireNeutral 
         Caption         =   "Set"
         Height          =   255
         Left            =   7380
         TabIndex        =   34
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtLimitFireNeutral 
         Height          =   285
         Left            =   6720
         TabIndex        =   33
         Text            =   "3000"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdSetLimitFireFast 
         Caption         =   "Set"
         Height          =   255
         Left            =   7380
         TabIndex        =   30
         Top             =   2100
         Width           =   495
      End
      Begin VB.CommandButton cmdSetLimitDown 
         Caption         =   "Set"
         Height          =   255
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdSetLimitRight 
         Caption         =   "Set"
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtLimitFireFast 
         Height          =   285
         Left            =   6720
         TabIndex        =   27
         Text            =   "2000"
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtLimitDown 
         Height          =   285
         Left            =   4260
         TabIndex        =   26
         Text            =   "2000"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtLimitRight 
         Height          =   285
         Left            =   4260
         TabIndex        =   25
         Text            =   "4000"
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmdSetLimitFireSlow 
         Caption         =   "Set"
         Height          =   255
         Left            =   7380
         TabIndex        =   21
         Top             =   1500
         Width           =   495
      End
      Begin VB.CommandButton cmdSetLimitUp 
         Caption         =   "Set"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdSetLimitLeft 
         Caption         =   "Set"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtLimitFireSlow 
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Text            =   "4000"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtLimitUp 
         Height          =   285
         Left            =   2100
         TabIndex        =   17
         Text            =   "4000"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtLimitLeft 
         Height          =   285
         Left            =   2100
         TabIndex        =   16
         Text            =   "2000"
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label lblCameraOffset 
         BackStyle       =   0  'Transparent
         Caption         =   "Now Click Image where the shots are hitting."
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   120
         TabIndex        =   79
         Top             =   1740
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Camera Offset:"
         Height          =   255
         Left            =   900
         TabIndex        =   78
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   ","
         Height          =   195
         Left            =   2700
         TabIndex        =   77
         Top             =   2220
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   "Left"
         Height          =   255
         Left            =   7620
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "Up"
         Height          =   255
         Left            =   7620
         TabIndex        =   47
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Right"
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Down"
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Fire Neutral:"
         Height          =   195
         Left            =   5580
         TabIndex        =   32
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Gun Limits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Fire Fast:"
         Height          =   195
         Left            =   6000
         TabIndex        =   24
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Down:"
         Height          =   195
         Left            =   3660
         TabIndex        =   23
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Right:"
         Height          =   255
         Left            =   3660
         TabIndex        =   22
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Fire Slow:"
         Height          =   195
         Left            =   5820
         TabIndex        =   15
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Up:"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   1860
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Left:"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lblRawFire 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblRawTilt 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   660
         Width           =   675
      End
      Begin VB.Label lblRawPan 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Fire:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Tilt:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Pan:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manual Control"
      Height          =   2475
      Left            =   8160
      TabIndex        =   0
      Top             =   60
      Width           =   3195
      Begin VB.Timer tmrFire2 
         Enabled         =   0   'False
         Left            =   2220
         Top             =   180
      End
      Begin VB.Timer tmrFireTime 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2220
         Top             =   1020
      End
      Begin VB.Timer tmrFire 
         Enabled         =   0   'False
         Left            =   2220
         Top             =   600
      End
      Begin VB.CommandButton cmdCenter 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1500
         TabIndex        =   38
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton cmdFireSlow 
         Caption         =   "Fire Slow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdFireFast 
         Caption         =   "Fire Fast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1695
      End
      Begin VB.HScrollBar scrPan 
         Height          =   375
         LargeChange     =   10
         Left            =   60
         Max             =   100
         Min             =   -100
         TabIndex        =   10
         Top             =   1980
         Width           =   2595
      End
      Begin VB.VScrollBar scrTilt 
         Height          =   2235
         LargeChange     =   10
         Left            =   2700
         Max             =   100
         Min             =   -100
         TabIndex        =   9
         Top             =   120
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const PS_SOLID = 0 ' solid line for lineto

' which channels each servo is on
Const SERVO_TILT = 0
Const SERVO_PAN = 1
Const SERVO_FIRE = 2
Dim camerasRunning As Boolean
Dim processed As Integer
Const CAMERA_FOV = 40 ' degrees
Const DIST_TO_ARENA = 95 ' inches from the camera to the targets, perpendicular


' Make sure that you register CapStill.dll and FSFWrap.dll
'dimensions of the image
Const IMAGE_WIDTH = 150 '250 '160
Const IMAGE_HEIGHT = 150 '50
Const OBJ_WIDTH = 30
Const OBJ_HEIGHT = 30
Const target_color_r = 255
Const target_color_g = 255
Const target_color_b = 255
Dim target_color_slop_r As Integer ' +- out of 255 to be close enough
Dim target_color_slop_g As Integer  ' +- out of 255 to be close enough
Dim target_color_slop_b As Integer  ' +- out of 255 to be close enough
Const MIN_PIXELS1 = 2
Const MIN_PIXELS2 = 1

Dim gGraph As IMediaControl
Dim gRegFilters As Object
Dim gCapStill As VBGrabber
Dim initialized As Boolean

Dim hMemDc As Long

'bitmaps for left and right images
Dim bmp() As Byte
Dim bma As IBitmapAccess

Dim NoOfBoxes As Long
Dim NoOfCorners As Long
Dim NoOfHooks As Long
Dim NoOfHorizontalLines As Long
Dim NoOfVerticalLines As Long
Dim NoOfVerticalSeparators As Long
Dim NoOfHorizontalSeparators As Long
Dim currentCameraIndex As Long

Dim tJunk As POINTAPI

Dim busy As Boolean
Dim center_x As Integer, center_y As Integer ' where the shots are actually hitting
Dim actual_center_x As Integer, actual_center_y As Integer ' center of the screen
Dim target_x As Integer, target_y As Integer ' where the next target is





Private Sub chkWebcamControl_Click()
  If chkWebcamControl.Value = vbChecked Then
    lblClock.Caption = "0.0"
    tmrClock.Enabled = True
  Else
    tmrClock.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  txtCommPort.Text = GetSetting("DefconBot", "Settings", "txtCommPort")
  txtLimitLeft.Text = GetSetting("DefconBot", "Settings", "txtLimitLeft")
  txtLimitRight.Text = GetSetting("DefconBot", "Settings", "txtLimitRight")
  txtLimitUp.Text = GetSetting("DefconBot", "Settings", "txtLimitUp")
  txtLimitDown.Text = GetSetting("DefconBot", "Settings", "txtLimitDown")
  txtLimitFireSlow.Text = GetSetting("DefconBot", "Settings", "txtLimitFireSlow")
  txtLimitFireFast.Text = GetSetting("DefconBot", "Settings", "txtLimitFireFast")
  txtLimitFireNeutral.Text = GetSetting("DefconBot", "Settings", "txtLimitFireNeutral")
  txtCameraOffsetX.Text = GetSetting("DefconBot", "Settings", "txtCameraOffsetX")
  txtCameraOffsetY.Text = GetSetting("DefconBot", "Settings", "txtCameraOffsetY")
  scrPan.Value = Int(GetSetting("DefconBot", "Settings", "scrPan", 0))
  scrTilt.Value = Int(GetSetting("DefconBot", "Settings", "scrTilt", 0))
  ScrRawPan.Value = Int(GetSetting("DefconBot", "Settings", "scrRawPan", 3000))
  ScrRawTilt.Value = Int(GetSetting("DefconBot", "Settings", "scrRawTilt", 3000))
  ScrRawFire.Value = Int(GetSetting("DefconBot", "Settings", "scrRawFire", 3000))
  scrTargetColor.Value = Int(GetSetting("DefconBot", "Settings", "scrtargetcolor", scrTargetColor.Value))
  scrCenterSize.Value = Int(GetSetting("DefconBot", "Settings", "scrCenterSize", scrCenterSize.Value))
  scrGrabTimer.Value = Int(GetSetting("DefconBot", "Settings", "tmrGrab", tmrGrab.Interval))
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSetting "DefconBot", "Settings", "txtCommPort", txtCommPort.Text
   SaveSetting "DefconBot", "Settings", "txtLimitLeft", txtLimitLeft.Text
   SaveSetting "DefconBot", "Settings", "txtLimitRight", txtLimitRight.Text
   SaveSetting "DefconBot", "Settings", "txtLimitUp", txtLimitUp.Text
   SaveSetting "DefconBot", "Settings", "txtLimitDown", txtLimitDown.Text
   SaveSetting "DefconBot", "Settings", "txtLimitFireSlow", txtLimitFireSlow.Text
   SaveSetting "DefconBot", "Settings", "txtLimitFireFast", txtLimitFireFast.Text
   SaveSetting "DefconBot", "Settings", "txtLimitFireNeutral", txtLimitFireNeutral.Text
   SaveSetting "DefconBot", "Settings", "txtCameraOffsetX", txtCameraOffsetX.Text
   SaveSetting "DefconBot", "Settings", "txtCameraOffsetY", txtCameraOffsetY.Text
   SaveSetting "DefconBot", "Settings", "scrPan", scrPan.Value
   SaveSetting "DefconBot", "Settings", "scrTilt", scrTilt.Value
   SaveSetting "DefconBot", "Settings", "scrRawPan", ScrRawPan.Value
   SaveSetting "DefconBot", "Settings", "scrRawTilt", ScrRawTilt.Value
   SaveSetting "DefconBot", "Settings", "scrRawFire", ScrRawFire.Value
   SaveSetting "DefconBot", "Settings", "scrtargetcolor", scrTargetColor.Value
   SaveSetting "DefconBot", "Settings", "scrcentersize", scrCenterSize.Value
   SaveSetting "DefconBot", "Settings", "tmrGrab", scrGrabTimer.Value
End Sub




Private Sub pctPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' offset is +- from the center of the webcam
   Dim dist_x, dist_y
   
   If lblCameraOffset.Visible = True Then
      dist_x = x - actual_center_x
      dist_y = y - actual_center_y
      
      txtCameraOffsetX.Text = dist_x
      txtCameraOffsetY.Text = dist_y
   
      center_x = actual_center_x + txtCameraOffsetX.Text
      center_y = actual_center_y + txtCameraOffsetY.Text
   End If
   lblCameraOffset.Visible = False
End Sub

Private Sub scrCenterSize_Change()
  lblCenterSize.Caption = scrCenterSize.Value
End Sub

Private Sub scrGrabTimer_Change()
  tmrGrab.Interval = scrGrabTimer.Value / scrGrabTimer.max * (1 * 1000)  ' 5 second max
  lblGrabTimer.Caption = tmrGrab.Interval & "ms"
End Sub

Private Sub scrTargetColor_Change()
  target_color_slop_r = scrTargetColor.Value
  target_color_slop_g = scrTargetColor.Value
  target_color_slop_b = scrTargetColor.Value
End Sub









''''''''''''''''''''''''''''''''''''''''''''
'' CAMERA CONTROL
''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCameraOn_Click()
'preview using the appropriate WDM filter name for the cameras
'if in doubt about the filter name have a look through filters.txt
   Call CameraPreview("Logitech QuickCam Zoom", 0)
   actual_center_x = pctPreview.ScaleWidth / 2
   actual_center_y = pctPreview.ScaleHeight / 2
   center_x = actual_center_x + txtCameraOffsetX.Text
   center_y = actual_center_y + txtCameraOffsetY.Text
   camerasRunning = True
End Sub



Private Sub cmdCameraOff_Click()
  If (initialized) Then
    initialized = False
    Call gGraph.Stop
    Set gGraph = Nothing
    Set gRegFilters = Nothing
    Set gCapStill = Nothing
  End If
  camerasRunning = False
End Sub



Private Sub CameraPreview(DriverName As String, cameraIndex As Integer)
  On Error GoTo CameraPreview_err
  
  Dim i As Integer
  Dim index As Integer
  Dim xbar As CrossbarInfo
  Dim pinOut As IPinInfo
  Dim idx As Long
  Dim filter As IRegFilterInfo
  Dim fGrab As IFilterInfo
  Dim fSrc As IFilterInfo
  Dim pin As String
  Dim found As Boolean
  Dim pSC As StreamConfig
  Dim pinIn As IPinInfo
  Dim ppropOut As PinPropInfo
  Dim strFilters As String
  Dim pinErr As Boolean
  
  pinErr = False
  
  'make a new graph
  Set gGraph = Nothing
  Set gCapStill = Nothing
  Set gGraph = New FilgraphManager
  Set gRegFilters = gGraph.RegFilterCollection
    
  
  'add the grabber including vb wrapper and default props
  found = False
  i = 0
  While (i < gRegFilters.Count) And (Not found)
    Call gRegFilters.Item(i, filter)
    If (filter.Name = "SampleGrabber") Then
      filter.filter fGrab
      'wrap this filter in the capstill vb wrapper
      'also sets rgb-24 media type and other properties
      Set gCapStill = New VBGrabber
      gCapStill.FilterInfo = fGrab
      found = True
    End If
    i = i + 1
  Wend
    
  
  Open App.Path & "\filters.txt" For Output As #1
  'strFilters = ""
  i = 0
  While (i < gRegFilters.Count)
    Call gRegFilters.Item(i, filter)
    Print #1, filter.Name
    'strFilters = strFilters & filter.Name & Chr(13)
    i = i + 1
  Wend
  Close #1
  'MsgBox strFilters
  

  'add the selected source filter
  'WDM drivers for the cameras can be identified by the word "QuickCam" in their title
  index = 0
  found = False
  i = 0
  While (i < gRegFilters.Count) And (Not found)
    Call gRegFilters.Item(i, filter)
    
    If (InStr(LCase(filter.Name), LCase(DriverName)) > 0) Then
      If (index = cameraIndex) Then
        filter.filter fSrc
        found = True
      End If
      index = index + 1
    End If
    i = i + 1
  Wend
  
    
  'find first output on src
  found = False
  i = 0
  While (i < fSrc.Pins.Count) And (Not found)
    Call fSrc.Pins.Item(i, pinOut)
    If (pinOut.Direction = 1) Then
      found = True
    End If
    i = i + 1
  Wend
  
    
  'restore specified file before dlg
  Set pSC = New StreamConfig
  pSC.pin = pinOut
  If (pSC.SupportsConfig) Then
    If (Dir$("mtsave.mt") <> "") Then
      'pSC.Restore ("mtsave.mt")
    End If
  End If
    
  'show format of output pin before rendering
  Set ppropOut = New PinPropInfo
  ppropOut.pin = pinOut
  
  
  'find first input on grabber and connect
  found = False
  i = 0
  While (i < fGrab.Pins.Count) And (Not found)
    Call fGrab.Pins.Item(i, pinIn)
    If (pinIn.Direction = 0) Then
      pinErr = True
      pinOut.Connect pinIn
      pinErr = False
      found = True
    End If
    i = i + 1
  Wend
    
  ' find grabber output pin and render
  found = False
  i = 0
  While (i < fGrab.Pins.Count) And (Not found)
    Call fGrab.Pins.Item(i, pinOut)
    If (pinOut.Direction = 1) Then
      pinOut.Render
      found = True
    End If
    i = i + 1
  Wend
       
  ' run graph and we are successfully in preview mode
  Call gGraph.Run
  
  'camera has been initialized
  initialized = True
  
CameraPreview_exit:
  Exit Sub
CameraPreview_err:
  If (pinErr) Then
    Resume CameraPreview_exit
  End If
  
  MsgBox "frmQuickCamStereo/CameraPreview/" & Err & "/" & Error$(Err)
  Resume CameraPreview_exit
End Sub




Private Sub cmdStop_Click()
  Dim i As Integer
  
  If (initialized) Then
    initialized = False
    Call gGraph.Stop
    Set gGraph = Nothing
    Set gRegFilters = Nothing
    Set gCapStill = Nothing
  End If
  camerasRunning = False
End Sub




Private Sub ShowBitmap(bma As IBitmapAccess)
   'set correct size of image and then
   'BitBlt to the picture control's HDC

   Dim hbm As Long
   Dim hOldBM As Long
   Dim initEye As Boolean
   Dim camIndex As Integer
   Static flip As Boolean
   
   camIndex = 0
   If (Not initEye) Then
      initEye = True
      If pctPreview.Width <> bma.Width * Screen.TwipsPerPixelX Then
         pctPreview.Width = bma.Width * Screen.TwipsPerPixelX
         pctPreview.Height = bma.Height * Screen.TwipsPerPixelY
      End If
      hMemDc = CreateCompatibleDC(pctPreview.hdc)
  End If

  hbm = bma.DIBSection

  hOldBM = SelectObject(hMemDc, hbm)
  BitBlt pctPreview.hdc, 0, 0, bma.Width, bma.Height, hMemDc, 0, 0, &HCC0020
  SelectObject hMemDc, hOldBM
  pctPreview.Refresh
End Sub





Private Sub tmrClock_Timer()
  lblClock.Caption = lblClock.Caption + tmrClock.Interval / 1000
End Sub


Private Sub tmrGrab_Timer()
'  On Error GoTo catch_error
On Error Resume Next
   Dim i As Integer
   Dim x(2) As Single
   Dim y(2) As Single
   Dim dist As Single
   
   If (initialized) And (camerasRunning) And (Not busy) Then
      busy = True
      Set bma = gCapStill.CapToMem
      If bma.DIBSection > 0 Then
         Call ShowBitmap(bma)
      
'      Dim bm As Bitmap
'      GetObject PictureBox.image, Len(bm), bm '
'      Dim ImageData() As Byte
'      ReDim ImageData(0 To (bm.bmBitsPixel \ 8) - 1, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
'      GetBitmapBits PictureBox.image, bm.bmWidthBytes * bm.bmHeight, ImageData(0, 0, 0)
'      Call FilterBitmap
         If processed >= 1 Then
           'pctOutput.Cls
           Call DecideWhatToDo
           processed = 0
         End If
         processed = processed + 1
         Call DrawOverlay
         pctPreview.Refresh
      End If
      busy = False
   End If
   
  Exit Sub
catch_error:
  ' nothing, go to next loop
End Sub

Private Function DrawOverlay()
' draws the overlay components onto the video. **MUST COME AFTER CHECKS** since it alters the image
   Dim hpen As Long
   Dim x, y, Red, color
   
   If chkTargetsShow.Value = vbChecked Then
      ' color all the detected targets
      ' decolor all the non-detected
      For x = 0 To pctPreview.ScaleWidth
         For y = 0 To pctPreview.ScaleHeight
            color = GetPixel(pctPreview.hdc, x, y)
            Red = ExtractR(color)
            If Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r Then
               SetPixelV pctPreview.hdc, x, y, RGB(Red, Red, Red)
            Else
               SetPixelV pctPreview.hdc, x, y, RGB(0, 0, 0)
            End If
         Next y
      Next x
      pctPreview.Refresh
'      MsgBox "waiting"
   End If
'  ' crosshairs - actual center
'  hpen = CreatePen(PS_SOLID, 1, RGB(128, 64, 64))
'  hpen = SelectObject(pctPreview.hdc, hpen)
'  ' horizontal crosshair
'  MoveToEx pctPreview.hdc, actual_center_x + 2, actual_center_y, tJunk
'  LineTo pctPreview.hdc, actual_center_x + 5, actual_center_y
'  MoveToEx pctPreview.hdc, actual_center_x - 2, actual_center_y, tJunk
'  LineTo pctPreview.hdc, actual_center_x - 5, actual_center_y
'  ' vertictal crosshair
'  MoveToEx pctPreview.hdc, actual_center_x, actual_center_y + 2, tJunk
'  LineTo pctPreview.hdc, actual_center_x, actual_center_y + 5
'  MoveToEx pctPreview.hdc, actual_center_x, actual_center_y - 2, tJunk
'  LineTo pctPreview.hdc, actual_center_x, actual_center_y - 5
    
  ' crosshairs - center +- offset
  hpen = CreatePen(PS_SOLID, 1, RGB(255, 0, 0))
  hpen = SelectObject(pctPreview.hdc, hpen)
  ' horizontal crosshair
  MoveToEx pctPreview.hdc, center_x + 2, center_y, tJunk
  LineTo pctPreview.hdc, center_x + 5, center_y
  MoveToEx pctPreview.hdc, center_x - 2, center_y, tJunk
  LineTo pctPreview.hdc, center_x - 5, center_y
  ' vertictal crosshair
  MoveToEx pctPreview.hdc, center_x, center_y + 2, tJunk
  LineTo pctPreview.hdc, center_x, center_y + 5
  MoveToEx pctPreview.hdc, center_x, center_y - 2, tJunk
  LineTo pctPreview.hdc, center_x, center_y - 5

' draw the line to the target
  If target_x <> -1 Then ' -1 = no target
    hpen = CreatePen(PS_SOLID, 1, RGB(0, 255, 0))
    hpen = SelectObject(pctPreview.hdc, hpen)
    MoveToEx pctPreview.hdc, center_x, center_y, tJunk
    LineTo pctPreview.hdc, target_x, target_y
  End If
  
  DeleteObject hpen
End Function


Private Sub chkNoTarget_Click()
  If chkWebcamControl.Value = vbChecked Then
   tmrNoTarget.Enabled = False
   If chkNoTarget.Value = vbChecked Then
     tmrNoTarget.Enabled = True ' resets interval to default
   End If
 End If
End Sub










Public Function DecideWhatToDo()
  ' search the image for targets
  ' sweep the area
  ' time out if failed to shoot down
  
  ' tmrFireTime: the maximum amount of time to try to shoot any one target
  ' tmrFireSearch: once firing starts, the changes its search pattern
  
  Dim Width As Integer, Height As Integer
  Dim x As Integer, y As Integer
  ' upper left = 0,0
  
  
  Dim center_white As Integer, center_white2 As Integer
  Dim timer_enabled As Boolean
  timer_enabled = tmrNoTarget.Enabled ' set at the end
  
  
  
  Height = pctPreview.ScaleHeight
  Width = pctPreview.ScaleWidth
    
  
  
  ' check if we're currently looking at a target,  if so then fire
'  center_white = NumTargets(center_y - height / 2 * 0.0169, center_x + width / 2 * 0.0126, center_y + height / 2 * 0.0169, center_x - width / 2 * 0.0126)
'  center_white2 = NumTargets(center_y - Height / 2 * 0.0693, center_x + Width / 2 * 0.0543, center_y + Height / 2 * 0.0593, center_x - Width / 2 * 0.0443)
  center_white2 = NumTargets(center_y - scrCenterSize.Value, center_x + scrCenterSize.Value, center_y + scrCenterSize.Value, center_x - scrCenterSize.Value)
'  lblWhitePixels.Caption = center_white
'  lblWhitePixels2.Caption = center_white
'  If center_white > MIN_PIXELS1 Then
'    ' the center 2.52% grid in the center are a target
'    tmrFireTime.Enabled = True
'    tmrFireSearch.Enabled = True
'    lblFiring.Visible = True
'    SendPololu SERVO_FIRE, txtLimitFireFast.Text
  If center_white2 >= MIN_PIXELS2 Then
    ' the next 8.86% are white
    tmrFireTime.Enabled = True
'    tmrFireSearch.Enabled = True
    If chkWebcamControl.Value = vbChecked And chkEnableFire.Value = vbChecked Then
      cmdFireSlow_Click
    End If
  Else
    If chkWebcamControl.Value = vbChecked Then
      SendPololu SERVO_FIRE, txtLimitFireNeutral.Text ' stop firing
    End If
    tmrFireTime.Enabled = False
    tmrFireSearch.Enabled = False
    lblFiring.Visible = False
  End If
  ' done checking whether to fire or not
  
  
  Dim color As Double
  Dim Red As Integer
  Dim loop_num As Integer
  Dim min_x As Integer, min_y As Integer, max_x As Integer, max_y As Integer
'  If tmrFireTime.Enabled = False Then
    ' if we're not on target, search for the next nearest target
    ' # = top, right, bottom, left
    Dim h As Double, w As Double
    Dim g1 As Integer, g2 As Integer, g3 As Integer, g4 As Integer, g5 As Integer, g6 As Integer, g7 As Integer, g8 As Integer
    Dim side As Integer
    Dim move_x As Double, move_y As Double, dist As Double
      
    ' starting from the center, spiral out looking for the nearest target.
    ' then calculate how much to move to get to that position and move
    x = center_x ' center of the screen
    y = center_y
    ' sides: 0=top, 1=right, 2=bottom, 3=left, always scan clockwise from the top left
    loop_num = 1
    target_x = -1 ' -1 = white not found yet
    Do While loop_num < center_y And target_x = -1 ' number of squares to search
      ' move us to the starting position for this square
      side = 0
      x = center_x - loop_num
      y = center_y - loop_num
      min_x = center_x - loop_num ' size of the box to search
      max_x = center_x + loop_num
      min_y = center_y - loop_num
      max_y = center_y + loop_num
      Do While side <= 3 And target_x = -1
        ' check the current pixel
        color = GetPixel(pctPreview.hdc, x, y)
        Red = ExtractR(color)
        If Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r Then
          
          ' now find the center of that white edge
          target_x = findcenter_x(x, y)
          target_y = findcenter_y(x, y)
          
          move_x = target_x - center_x
          move_y = target_y - center_y
          dist = Abs(Sqr(move_x * move_x + move_y * move_y))
          AddHistory "Moving " & move_x & ", " & move_y
          
          ' now that we know where it is, we can calculate how far to move
          If chkWebcamControl.Value = vbChecked Then
            MoveRelative move_x / (Width * 2), move_y / (Height * 2)
          End If
          AddHistory "Found @" & target_x & ", " & target_y & ", " & dist & " pixels away"
        End If
      
      
        ' move to the next position
        If side = 0 Then
          If x >= max_x Then
            side = side + 1
          Else
            x = x + 1
          End If
        ElseIf side = 1 Then
          If y >= max_y Then
            side = side + 1
          Else
            y = y + 1
          End If
        ElseIf side = 2 Then
          If x <= min_x Then
            side = side + 1
          Else
            x = x - 1
          End If
        ElseIf side = 3 Then
          If y <= min_y Then
            side = side + 1  ' next square
          Else
            y = y - 1
          End If
        Else
          MsgBox "Side error: side=" & side
        End If
      Loop
      loop_num = loop_num + 1
    Loop
'  End If
'  pctOutput.Refresh
    
    
      
'  tmrNoTarget.Enabled = timer_enabled
End Function


Public Function NumTargets(top As Integer, right As Integer, bottom As Integer, left As Integer) As Double
  ' returns the number of pixels that match the target color
  Dim target_sum As Integer
  Dim x As Integer, y As Integer
  Dim color As Double
  Dim Red As Integer, Green As Integer, Blue As Integer
  Dim array_size As Integer
  
  target_sum = 0 ' array of the values, 0=not target, 1=target, take average at
  For y = top To bottom
    For x = left To right
      color = GetPixel(pctPreview.hdc, x, y)
      'SetPixelV pctOutput.hdc, x, y, RGB(255, 255, 255)
      Red = ExtractR(color)
'      green = ExtractG(color)
'      blue = ExtractB(color)
      If Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r Then
        'SetPixelV pctOutput.hdc, x, y, RGB(255, 0, 0)
'        If green > target_color_g - target_color_slop_g And green < target_color_g + target_color_slop_g Then
'          SetPixelV pctPreview.hdc, x, y, RGB(0, 255, 0)
'          If blue > target_color_b - target_color_slop_b And blue < target_color_b + target_color_slop_b Then
'            SetPixelV pctPreview.hdc, x, y, RGB(0, 0, 255)
            ' pixel is a target color
            target_sum = target_sum + 1
'          End If
'        End If
      End If
    Next
  Next
'  pctOutput.Refresh
'  array_size = (top - bottom) * (right - left)
'  If array_size > 0 Then
'    NumTargets = target_sum / array_size ' average
'  Else
'    NumTargets = 0
'  End If
  lblNumTargets.Caption = target_sum
  NumTargets = target_sum
End Function



Private Function findcenter_x(x As Integer, y As Integer) As Integer
' On Error GoTo catch_error
  ' find the x component center of the blob which intersects (x, y)
  
  Dim min, max, color, Red
  min = x
  max = x
  color = GetPixel(pctPreview.hdc, min, y)
  Red = ExtractR(color)
  Do While min > 0 And Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r
    min = min - 1
    color = GetPixel(pctPreview.hdc, min, y)
    Red = ExtractR(color)
  Loop
  color = GetPixel(pctPreview.hdc, max, y)
  Red = ExtractR(color)
  Do While max < pctPreview.ScaleWidth And Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r
    max = max + 1
    color = GetPixel(pctPreview.hdc, max, y)
    Red = ExtractR(color)
  Loop
  
  findcenter_x = min + (max - min) / 2
  
  Exit Function
catch_error:
  MsgBox "findcenter_x: " & Err & "/" & Error$(Err)
End Function

Private Function findcenter_y(x As Integer, y As Integer) As Integer
' On Error GoTo catch_error
  ' find the x component center of the blob which intersects (x, y)
  
  Dim min, max, color, Red
  min = y
  max = y
  color = GetPixel(pctPreview.hdc, x, min)
  Red = ExtractR(color)
  Do While min > 0 And Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r
    min = min - 1
    color = GetPixel(pctPreview.hdc, x, min)
    Red = ExtractR(color)
  Loop
  
  color = GetPixel(pctPreview.hdc, x, max)
  Red = ExtractR(color)
  Do While max < pctPreview.ScaleHeight And Red > target_color_r - target_color_slop_r And Red < target_color_r + target_color_slop_r
    max = max + 1
    color = GetPixel(pctPreview.hdc, x, max)
    Red = ExtractR(color)
  Loop
  
  findcenter_y = min + (max - min) / 2
  
  Exit Function
catch_error:
  MsgBox "findcenter_y: " & Err & "/" & Error$(Err)
End Function




Private Sub tmrFireTime_Timer()
  ' failed to shoot down the target, so abort and move to a new area
   If chkNoTarget.Value = vbChecked Then
      If chkWebcamControl.Value = vbChecked Then
         tmrFireTime.Enabled = False
         scrTilt.Value = rand(scrTilt.min, scrTilt.max)
         scrPan.Value = rand(scrPan.min, scrPan.max)
         lblFiring.Visible = False
         AddHistory "Failed to shoot down target in " & (tmrFireTime.Interval / 1000) & " seconds, randomizing"
      End If
   End If
End Sub

Private Sub tmrNoTarget_Timer()
' recenter every x seconds
    If chkNoTarget.Value = vbChecked Then
      scrTilt.Value = 0
      scrPan.Value = 0
      AddHistory "No target found in " & (tmrNoTarget.Interval / 1000) & " seconds, randomizing"
    End If
  ' Haven't found any targets in the maximum time
'  If chkNoTarget.Value = vbChecked Then
'    scrTilt.Value = rand(scrTilt.min, scrTilt.max)
'    scrPan.Value = rand(scrPan.min, scrPan.max)
'    AddHistory "No target found in " & (tmrNoTarget.Interval / 1000) & " seconds, randomizing"
'  End If
End Sub

'Private Sub tmrFireSearch_Timer()
'  ' firing on the target, do a search pattern
'
'  Dim newval As Integer
'  Dim diff_x As Integer, diff_y As Integer
'  diff_x = 0
'  diff_y = 0
'
'  Select Case tmrFireTime.Tag
'  ' first loop
'  Case 0:
'    diff_x = 1
'    diff_y = 0
'  Case 1:
'    diff_x = 1
'    diff_y = 0
'  Case 2:
'    diff_x = -1
'    diff_y = -1
'  Case 3:
'    diff_x = -1
'    diff_y = -1
'  Case 4:
'    diff_x = -1
'    diff_y = 1
'  Case 5:
'    diff_x = -1
'    diff_y = 1
'  Case 6:
'    diff_x = 1
'    diff_y = 1
'  Case 7:
'    diff_x = 1
'    diff_y = 1
'  Case 8:
'    diff_x = 1
'    diff_y = -1
'  Case 9:
'    diff_x = 1
'    diff_y = -2
'
'  ' second loop
'  Case 10:
'    diff_x = 0
'    diff_y = -1
'  Case 11:
'    diff_x = -1
'    diff_y = 0
'  Case 12:
'    diff_x = -1
'    diff_y = 1
'  Case 13:
'    diff_x = -1
'    diff_y = -1
'  Case 14:
'    diff_x = -1
'    diff_y = 0
'  Case 15:
'    diff_x = 0
'    diff_y = 1
'  Case 16:
'    diff_x = 1
'    diff_y = 1
'  Case 17:
'    diff_x = -1
'    diff_y = 1
'  Case 18:
'    diff_x = 0
'    diff_y = 1
'  Case 19:
'    diff_x = 1
'    diff_y = 0
'  Case 20:
'    diff_x = 1
'    diff_y = -1
'  Case 21:
'    diff_x = 1
'    diff_y = 1
'  Case 22:
'    diff_x = 1
'    diff_y =
'  Case 23:
'    diff_x = 0
'    diff_y = -1
'  Case Else:
'    tmrFireTime.Tag = -1 ' loop
'  End Select
'
'  newval = scrPan.Value + diff_x * 0.0175 * (scrPan.max - scrPan.min)
'  If newval < -100 Then
'    AddHistory "Target Pan < -100: " & newval
'    newval = -100
'  ElseIf newval > 100 Then
'    AddHistory "Target Pan > 100: " & newval
'    newval = 100
'  End If
'  scrPan.Value = newval
'
'  newval = scrTilt.Value + diff_y * 0.0175 * (scrTilt.max - scrTilt.min)
'  If newval < -100 Then
'    AddHistory "Target Tilt < -100: " & newval
'    newval = -100
'  ElseIf newval > 100 Then
'    AddHistory "Target Tilt > 100: " & newval
'    newval = 100
'  End If
'  scrTilt.Value = newval
'
'  tmrFireTime.Enabled = True
'  tmrFireTime.Tag = tmrFireTime.Tag + 1
'End Sub





Function MoveRelative(x_pct As Double, y_pct As Double)
  ' move to a position relative to the current position
  ' x_pct = -0.05 means move 5% left
  ' y_pct = -0.8 means move 80% down

  Dim newval As Integer
  newval = scrPan.Value + x_pct * (scrPan.max - scrPan.min)
  If newval < scrPan.min Then
    AddHistory "MoveRelative: new x < min = " & newval
    newval = scrPan.min
  ElseIf newval > scrPan.max Then
    AddHistory "MoveRelative: new x > max = " & newval
    newval = scrPan.max
  ElseIf newval = 0 And x_pct > 0 Then
    AddHistory "MoveRelative: new x = 0, rounding to 1 = " & x_pct
    newval = 1
  ElseIf newval = 0 And x_pct < 0 Then
    AddHistory "MoveRelative: new x = 0, rounding to -1 = " & x_pct
    newval = -1
  End If
  scrPan.Value = newval
  
  newval = scrTilt.Value + y_pct * (scrTilt.max - scrTilt.min)
  If newval < scrTilt.min Then
    AddHistory "MoveRelative: new y < min = " & newval
    newval = scrTilt.min
  ElseIf newval > scrTilt.max Then
    AddHistory "MoveRelative: new y > may = " & newval
    newval = scrTilt.max
  ElseIf newval = 0 And y_pct > 0 Then
    AddHistory "MoveRelative: new y = 0, rounding to 1 = " & y_pct
    newval = 1
  ElseIf newval = 0 And y_pct < 0 Then
    AddHistory "MoveRelative: new y = 0, rounding to -1 = " & y_pct
    newval = -1
  End If
  scrTilt.Value = newval

  ' update the display
'  Dim x As Double, y As Double
  'x = scrpan.Value / (scrpan.Max - scrpan.min)
  lblPosition = scrPan.Value & ", " & scrTilt.Value
  
  AddHistory "MoveRelative: (" & (x_pct * 100) & "%," & (y_pct * 100) & "%) -> (" & scrPan.Value & ", " & scrTilt.Value & ")"
End Function





'''''''''''''''''''''''''''''''''''''''''''
'' SERVO CONTROL
'''''''''''''''''''''''''''''''''''''''''''


Private Sub cmdCommOn_Click()
  If MSComm1.PortOpen = True Then
    AddCommLog "Comm Off"
    MSComm1.PortOpen = False
  End If
  MSComm1.PortOpen = True
  AddCommLog "Comm On"
End Sub

Private Sub cmdCommOff_Click()
  If MSComm1.PortOpen = True Then
    AddCommLog "Comm Off"
    MSComm1.PortOpen = False
  End If
  MSComm1.CommPort = txtCommPort.Text
End Sub


Private Sub cmdSetLimitLeft_Click()
   txtLimitLeft.Text = ScrRawPan.Value
End Sub
Private Sub cmdSetLimitRight_Click()
   txtLimitRight.Text = ScrRawPan.Value
End Sub
Private Sub cmdSetLimitUp_Click()
   txtLimitUp.Text = ScrRawTilt.Value
End Sub
Private Sub cmdSetLimitDown_Click()
   txtLimitDown.Text = ScrRawTilt.Value
End Sub
Private Sub cmdSetLimitSlow_Click()
   txtLimitFireSlow.Text = ScrRawFire.Value
End Sub
Private Sub cmdSetLimitFireFast_Click()
   txtLimitFireFast.Text = ScrRawFire.Value
End Sub
Private Sub cmdSetLimitFireNeutral_Click()
   txtLimitFireNeutral.Text = ScrRawFire.Value
End Sub

Private Sub cmdSetCameraOffset_Click()
   ' set the camera offset
   ' click this to start, then capture the click on the webcam image and calculate the position
   lblCameraOffset.Visible = True
End Sub




Private Sub ScrRawPan_Change()
   AddCommLog "Pan->" & ScrRawPan.Value
   SendPololu SERVO_PAN, ScrRawPan.Value
End Sub
Private Sub ScrRawTilt_Change()
   AddCommLog "Tilt->" & ScrRawTilt.Value
   SendPololu SERVO_TILT, ScrRawTilt.Value
End Sub
Private Sub ScrRawFire_Change()
   SendPololu SERVO_FIRE, ScrRawFire.Value
End Sub






Private Sub scrPan_Change()
   Dim val, amin, amax, diff, center, cmd As Integer
   val = scrPan.Value ' -100 to 100
   amin = txtLimitLeft.Text
   amax = txtLimitRight.Text
   diff = amax - amin
   center = diff / 2 + amin
   
   ScrRawPan.Value = center + (val / 200) * diff
'   cmd = center + (val / 200) * diff
'   AddCommLog "Pan->" & val & "(" & cmd & ")"
'   SendPololu SERVO_PAN, cmd
End Sub

Private Sub scrTilt_Change()
   Dim val, amin, amax, diff, center, cmd As Integer
   val = -scrTilt.Value ' -100 to 100
   amin = txtLimitDown.Text
   amax = txtLimitUp.Text
   diff = amax - amin
   center = diff / 2 + amin
   ScrRawTilt.Value = center + (val / 200) * diff
   
'   cmd = center + (val / 200) * diff
'   AddCommLog "Tilt->" & val & "(" & cmd & ")"
'   SendPololu SERVO_TILT, cmd
End Sub

Private Sub cmdCenter_Click()
   scrTilt.Value = 1
   scrTilt.Value = 0
   Sleep (200)
   scrPan.Value = 1
   scrPan.Value = 0
   SendPololu SERVO_FIRE, txtLimitFireNeutral.Text
End Sub

Private Sub cmdFireSlow_Click()
   AddCommLog "Slow"
      
   lblFiring.Visible = True ' turn on the warning
   
   If tmrFire.Enabled = True Then ' if already firing, just keep going
      tmrFire.Interval = 350
   Else ' else we need to kick on in high speed, then slow down
      SendPololu SERVO_FIRE, txtLimitFireFast.Text
      tmrFire2.Interval = 150
      tmrFire2.Enabled = True
   End If
End Sub

Private Sub cmdFireFast_Click()
   AddCommLog "FireFast"
   
   tmrFire.Interval = 500
   tmrFire.Enabled = True
   lblFiring.Visible = True
   SendPololu SERVO_FIRE, txtLimitFireFast.Text
End Sub


Private Sub tmrFire_Timer()
   ' firing has ended, so stop the gun and reset its state
   tmrFire.Enabled = False ' no repeat
   
   SendPololu SERVO_FIRE, txtLimitFireNeutral.Text
   
   lblFiring.Visible = False
End Sub


Private Sub tmrFire2_Timer()
   ' slow-fire is a two stage process, this is stage 2
   tmrFire2.Enabled = False ' no repeat
   
   SendPololu SERVO_FIRE, txtLimitFireSlow.Text
   
   tmrFire.Enabled = True
   tmrFire.Interval = 350
End Sub





Private Sub tmrGUI_Timer()
   lblRawPan.Caption = ScrRawPan.Value
   lblRawTilt.Caption = ScrRawTilt.Value
   lblRawFire.Caption = ScrRawFire.Value
End Sub











Private Function SendPololu(num As Integer, Value As Integer)
   ' num is 0-7
   ' value is 500-5500
   ' hex(Asc(Mid(cmd, 1, 1))) & "." & hex(Asc(Mid(cmd, 2, 1))) & "." & hex(Asc(Mid(cmd, 3, 1))) & "." & hex(Asc(Mid(cmd, 4, 1))) & "." & hex(Asc(Mid(cmd, 5, 1))) & "." & hex(Asc(Mid(cmd, 6, 1)))

   Dim cmd As String
   Dim strbinary As String
   Dim lsb As String
   Dim msb As String
   
   If MSComm1.PortOpen = True Then
      cmd = Chr(&H80)      ' start byte   - 0x80 ' always
      cmd = cmd & Chr(&H1) ' Device ID    - 0x01 ' always
      cmd = cmd & Chr(&H4) ' command      - 0x04 ' we want command 4: Set Position, Absolute (2 data bytes)
      cmd = cmd & Chr(num) ' servo num    - 0x.. ' 00-07
      
      ' data1        - 0x.. ' upper 7 bits - range is 500 through 5500
      ' data2        - 0x.. ' lower 7 bits
      strbinary = "0000000000000000" & dec2bin(Value)
      lsb = "0" & Mid(strbinary, Len(strbinary) - 6)
      msb = "0" & Mid(strbinary, Len(strbinary) - 12, 6)
      
      cmd = cmd & Chr(bin2dec(msb))
      cmd = cmd & Chr(bin2dec(lsb))
      
      MSComm1.Output = cmd & vbNewLine
      Sleep 20
      DoEvents
   End If
End Function




Private Function AddCommLog(msg As String)
  txtCommLog.Text = Replace(msg, "|", vbTab) & vbNewLine & Mid(txtCommLog.Text, 1, 2000)
End Function

Private Function AddHistory(msg As String)
  txtHistory.Text = Replace(msg, "|", vbTab) & vbNewLine & Mid(txtHistory.Text, 1, 2000)
End Function

















'''''''''' GENERIC FUNCTIONS
Function GetToken(ByVal strVal As String, intIndex As Integer, strDelimiter As String) As String
'-------------------------------------------------------
' Author  : Troy DeMonbreun (vb@8x.com)
' source  : http://www.freevbcode.com/ShowCode.asp?ID=161
' Revised : 12/22/1998
'-------------------------------------------------------
   Dim strSubString() As String
   Dim intIndex2 As Integer
   Dim i As Integer
   Dim intDelimitLen As Integer
   
   intIndex2 = 1
   i = 0
   intDelimitLen = Len(strDelimiter)
   
   Do While intIndex2 > 0
      ReDim Preserve strSubString(i + 1)
      intIndex2 = InStr(1, strVal, strDelimiter)
      If intIndex2 > 0 Then
         strSubString(i) = Mid(strVal, 1, (intIndex2 - 1))
         strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
      Else
         strSubString(i) = strVal
      End If
      i = i + 1
   Loop
   
   If intIndex > (i + 1) Or intIndex < 1 Then
      GetToken = ""
   Else
      GetToken = strSubString(intIndex - 1)
   End If
End Function

Function substr_count(haystack, needle)
  substr_count = UBound(Split(UCase(haystack), UCase(needle)))
End Function


Public Function dec2bin(mynum As Variant) As String
   ' from http://cuinl.tripod.com/Tips/decimaltobinary.htm
   Dim loopcounter As Integer
   If mynum >= 2 ^ 31 Then
      dec2bin = "Number too big"
      Exit Function
   End If
   Do
      If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
         dec2bin = "1" & dec2bin
      Else
         dec2bin = "0" & dec2bin
      End If
      loopcounter = loopcounter + 1
   Loop Until 2 ^ loopcounter > mynum
End Function
Private Function bin2dec(ByVal BinValue As String) As Long
   ' from http://cuinl.tripod.com/Tips/math5.htm
   Dim lngValue As Long
   Dim x As Long
   Dim k As Long
   k = Len(BinValue) ' will only work with 32 or fewer "bits"
   For x = k To 1 Step -1 ' work backwards down string
      If Mid$(BinValue, x, 1) = "1" Then
         If k - x > 30 Then ' bit 31 is the sign bit
            lngValue = lngValue Or -2147483648# ' avoid overflow error
         Else
            lngValue = lngValue + 2 ^ (k - x)
         End If
      End If
   Next x
   bin2dec = lngValue
End Function



Public Function rand(min As Integer, max As Integer) As Integer
  Randomize
  rand = Int((max - min) * Rnd + min)
End Function
Public Function ExtractR(ByVal CurrentColor As Long) As Byte
  ExtractR = CurrentColor And 255
End Function
Public Function ExtractG(ByVal CurrentColor As Long) As Byte
  ExtractG = (CurrentColor \ 256) And 255
End Function
Public Function ExtractB(ByVal CurrentColor As Long) As Byte
  ExtractB = (CurrentColor \ 65536) And 255
End Function













'GDI functions to draw a DIBSection into a DC
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal width As Long, ByVal height As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mode As Long) As Long
'Private Declare Sub DeleteDC Lib "gdi32" (ByVal hdc As Long)
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal count As Long)

'Private Type Bitmap
'  bmType As Long
'  bmWidth As Long
'  bmHeight As Long
'  bmWidthBytes As Long
'  bmPlanes As Integer
'  bmBitsPixel As Integer
'  bmBits As Long
'End Type
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
'Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
'Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Sub FilterBitmap()
' IR shows up as a specific color
' so, filter for that color, removes the garbage
  
  Dim x As Integer, y As Integer
  Dim color As Long
  Dim Red As Integer, Green As Integer, Blue As Integer
  For x = 0 To pctPreview.ScaleWidth
    For y = 0 To pctPreview.ScaleHeight
      color = GetPixel(pctPreview.hdc, x, y)
      Red = ExtractR(color)
      Green = ExtractG(color)
      Blue = ExtractB(color)
      Green = 0
      Blue = 0
'      SetPixelV pctOutput.hdc, x, y, RGB(Red, Green, Blue)
    Next y
  Next x
End Sub

Private Function SecondsToClock(seconds2 As Double) As String
' On Error GoTo catch_error
  ' takes in a number of seconds and returns a clock format
  ' < 60 seconds: 48.3
  ' else: 4:34
  Dim ret, minutes
  Dim seconds
  seconds = seconds2
  ret = ""
  minutes = 0
  
  If seconds < 1 Then ' add the 10ths
    ret = "0." & Int(seconds * 10)
  ElseIf seconds <= 10 Then ' add the 10ths
    ret = Int(seconds) & "." & Mid(Int(seconds * 10), 2, 1)
  ElseIf seconds <= 60 Then ' add the 10th of seconds
    ret = Int(seconds) & "." & Mid(Int(seconds * 10), 3, 1)
  Else
    minutes = 0
    Do While seconds >= 60
      minutes = minutes + 1
      seconds = seconds - 60
    Loop
    ' minutes now has the minutes
    ' seconds has the remainder
    If seconds < 10 Then
      ret = minutes & ":0" & Int(seconds)
    Else
      ret = minutes & ":" & Int(seconds)
    End If
  End If
    
  SecondsToClock = ret
  
  Exit Function
catch_error:
  MsgBox "SecondsToClock: " & Err & "/" & Error$(Err)
End Function

