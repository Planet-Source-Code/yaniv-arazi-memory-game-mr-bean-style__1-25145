VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@Mr. Bean Memory Game@"
   ClientHeight    =   5415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFF00&
      Caption         =   "?"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox cmdErrors 
      BackColor       =   &H00FFFF00&
      DownPicture     =   "Form1.frx":08CA
      Height          =   615
      Left            =   1200
      Picture         =   "Form1.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Check Erorrs"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdGiveUP 
      BackColor       =   &H00FFFF00&
      Caption         =   "Give Up?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer timFind 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   840
      Top             =   480
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   31
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFF00&
      Caption         =   "&New Game!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer timMain 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   360
      Top             =   480
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   23
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   22
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   21
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   20
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   19
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   18
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   17
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   16
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   15
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   14
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   13
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   12
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   11
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   10
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   9
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   8
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   7
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   5
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   24
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   25
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   26
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   27
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   28
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   29
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   615
      Index           =   30
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lblErrors 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   600
      Left            =   120
      TabIndex        =   68
      ToolTipText     =   "Check Errors"
      Top             =   120
      Visible         =   0   'False
      Width           =   945
      WordWrap        =   -1  'True
   End
   Begin VB.Image PicFind 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   5160
      Picture         =   "Form1.frx":1A5E
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image picFind1 
      Height          =   585
      Left            =   5160
      Picture         =   "Form1.frx":2328
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   12
      Left            =   2760
      Picture         =   "Form1.frx":2D96
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   15
      Left            =   4560
      Picture         =   "Form1.frx":3660
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   14
      Left            =   3960
      Picture         =   "Form1.frx":3F2A
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   13
      Left            =   3360
      Picture         =   "Form1.frx":47F4
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   11
      Left            =   2160
      Picture         =   "Form1.frx":50BE
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   10
      Left            =   1560
      Picture         =   "Form1.frx":5988
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   9
      Left            =   960
      Picture         =   "Form1.frx":6252
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   8
      Left            =   360
      Picture         =   "Form1.frx":6B1C
      Top             =   6600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   7
      Left            =   4560
      Picture         =   "Form1.frx":73E6
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   6
      Left            =   3960
      Picture         =   "Form1.frx":7CB0
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   5
      Left            =   3360
      Picture         =   "Form1.frx":857A
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   4
      Left            =   2760
      Picture         =   "Form1.frx":8E44
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   3
      Left            =   2160
      Picture         =   "Form1.frx":970E
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   2
      Left            =   1560
      Picture         =   "Form1.frx":9FD8
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   1
      Left            =   960
      Picture         =   "Form1.frx":A8A2
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":B16C
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "øîä"
      Visible         =   0   'False
      Begin VB.Menu mnuStart 
         Caption         =   "îúçéì"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAdvanced 
         Caption         =   "îú÷ãí"
      End
      Begin VB.Menu mnuSuperior 
         Caption         =   "çáì""æ"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nArrNums() As Integer
Dim nTemp As Integer
Dim X As Integer
Dim nClicks As Byte
Dim nVisiblePicsTag(0 To 1) As Integer
Dim nVisiblePicsIndex(0 To 1) As Integer
Dim b2Clicks As Boolean, bFound As Boolean
Dim nMinVal As Integer
Dim nErrors As Integer, nFind As Integer
Dim sHiScore As String
Private Sub cmd_Click(Index As Integer)
   If b2Clicks = False And bFound = False Then
        If nClicks = 1 Then
            b2Clicks = True
            nVisiblePicsTag(nClicks) = cmdPic(Index).Tag
            nVisiblePicsIndex(nClicks) = Index
            cmdPic(nVisiblePicsIndex(nClicks)).Visible = True
            CheckPics
        Else
            nVisiblePicsTag(nClicks) = cmdPic(Index).Tag
            nVisiblePicsIndex(nClicks) = Index
    If b2Clicks = False Then cmdPic(nVisiblePicsIndex(nClicks)).Visible = True
nClicks = 1
        End If
    End If
End Sub
Private Sub ShowPics(nPicIndex)
cmdPic(nPicIndex).Visible = True
nVisiblePicsIndex(nClicks) = nPicIndex
End Sub
Private Sub CheckPics()
Dim tmpX As Integer

If nFind = 15 Then
    Beep
    GameOver
    Exit Sub
Else
    If nVisiblePicsTag(0) = nVisiblePicsTag(1) Then
            b2Clicks = False
            nClicks = 0
            nFind = nFind + 1
            bFound = True
            timFind.Enabled = True
    Else
        timMain.Enabled = True
    End If
End If

End Sub

Private Sub cmdAbout_Click()
MsgBox "Memory Game" & vbCrLf & _
"yanivarazi@hotmail.com" & vbCrLf & _
"App was made for MCSD project (Nr.1)", vbInformation, "About"
End Sub

Private Sub cmdErrors_Click()
If cmdErrors.Value = 1 Then
lblErrors.Visible = True
Else
lblErrors.Visible = False
End If
End Sub

Private Sub cmdExit_Click()
Dim nResult As Integer
nResult = MsgBox("Quit Mr. Bean?", vbYesNo + vbQuestion, "Bean")

If nResult = vbYes Then
    Unload Me
End If
End Sub

Private Sub cmdGiveUP_Click()
Dim tmpX As Integer
Dim nResult As Integer

Beep
nResult = MsgBox("Give up?!", vbQuestion + vbYesNo, "Been")

If nResult = vbYes Then

    For tmpX = cmdPic.LBound To cmdPic.UBound
        cmdPic(tmpX).Visible = True
    Next tmpX

End If

End Sub

Private Sub cmdStart_Click()
Form_Load
Dim K As Integer

For K = cmdPic.LBound To cmdPic.UBound
    cmd(K).Visible = True
    cmdPic(K).Visible = False
    cmdPic(K).BackColor = &HC0C000
Next K

X = 0
nFind = 0
nErrors = 0
nClicks = 0
b2Clicks = False
lblErrors.Caption = "You have " & nErrors & " errors"
cmdGiveUP.Enabled = True
Start 0, 15
Start 16, 31
End Sub
Private Sub Form_Load()
On Error Resume Next

Open App.Path & "\HiScore.MrBin" For Input As #1
Input #1, sHiScore
Close #1

For K = cmdPic.LBound To cmdPic.UBound
    cmd(K).Visible = False
Next K

End Sub
Private Function MakeRandomPics()
Randomize Timer
Dim I As Integer
nTemp = (Rnd * 15)

    For I = nMinVal To (X - 1)
        If nTemp = nArrNums(I) Then
                MakeRandomPics
                Exit Function
           
        End If
       
    Next I
    
nArrNums(X) = nTemp
cmdPic(X).Picture = Pic(nArrNums(X)).Picture
cmdPic(X).Tag = nArrNums(X)
'Debug.Print X & ")" & " " & nArrNums(X)
Exit Function

End Function

Private Sub timFind_Timer()

For tmpX = 0 To 1
    cmdPic(nVisiblePicsIndex(tmpX)).BackColor = vbYellow
    cmdPic(nVisiblePicsIndex(tmpX)).Picture = PicFind.Picture
    cmdPic(nVisiblePicsIndex(tmpX)).Visible = True
Next tmpX
timFind.Enabled = False
bFound = False
End Sub

Private Sub timMain_Timer()
cmdPic(nVisiblePicsIndex(0)).Visible = False
cmdPic(nVisiblePicsIndex(1)).Visible = False
b2Clicks = False
nClicks = 0
nErrors = nErrors + 1
lblErrors.Caption = "You have " & nErrors & " errors"
timMain.Enabled = False
End Sub
Private Sub Start(nMin As Integer, nMax As Integer)
ReDim nArrNums(nMin To nMax) As Integer
Randomize Timer
nArrNums(nMin) = (Rnd * 15)
cmdPic(nMin).Picture = Pic(nArrNums(nMin)).Picture
nMinVal = nMin
Do Until X = (nMax + 1)
    MakeRandomPics
    X = X + 1
Loop
End Sub
Private Sub GameOver()
Dim nResult As Integer
Dim sBitHiScore As String

If nErrors < Val(sHiScore) Or sHiScore = "" Then
    Open App.Path & "\HiScore.MrBin" For Output As #1
    Print #1, CStr(nErrors)
    Close #1
    
If nerros < Val(sHiScore) Then sBitHiScore = "You bit hi score!"
    
    nResult = MsgBox("Game Over!" & vbCrLf & vbCrLf & _
    "Your score is - " & nErrors & " errors " & vbclrf & _
    "The hi score is - " & " " & Val(sHiScore) & vbCrLf & _
    sBitHiScore & "  " & "Play another game?", vbQuestion + vbYesNo, "Memory Game")
End If

If nResult = vbYes Then
    Call cmdStart_Click
        Else
    Unload Me
End If

End Sub
