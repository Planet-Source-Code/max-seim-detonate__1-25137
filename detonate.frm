VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00404040&
   Caption         =   "Detonate! ... Cut the correct RED jumpers to diffuse the bomb!!       Only cut the LOGICAL 1 BITS"
   ClientHeight    =   7815
   ClientLeft      =   1695
   ClientTop       =   2295
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10020
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   7140
      Width           =   7125
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CUT"
      Height          =   345
      Index           =   2
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   5790
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CUT"
      Height          =   345
      Index           =   1
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3750
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CUT"
      Height          =   345
      Index           =   0
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1710
      Width           =   525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   4590
      TabIndex        =   84
      Top             =   60
      Width           =   765
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   85
         Top             =   180
         Width           =   555
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   6660
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   83
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   83
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   82
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   82
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   81
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   81
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   80
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   80
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   79
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   79
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   78
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   78
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   77
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   77
      Top             =   6450
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   76
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   76
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   75
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   75
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   74
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   74
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   73
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   73
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   72
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   72
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   71
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   71
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   70
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   70
      Top             =   5940
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   69
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   69
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   68
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   68
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   67
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   67
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   66
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   66
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   65
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   65
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   64
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   64
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   63
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   63
      Top             =   5430
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   62
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   62
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   61
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   61
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   60
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   60
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   59
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   59
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   58
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   58
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   57
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   57
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   56
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   56
      Top             =   4920
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   55
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   55
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   54
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   54
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   53
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   53
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   52
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   52
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   51
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   51
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   50
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   50
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   49
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   49
      Top             =   4410
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   48
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   48
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   47
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   47
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   46
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   46
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   45
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   45
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   44
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   44
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   43
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   43
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   42
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   42
      Top             =   3900
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   41
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   41
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   40
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   40
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   39
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   39
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   38
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   38
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   37
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   37
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   36
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   36
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   35
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   35
      Top             =   3390
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   34
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   34
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   33
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   33
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   32
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   32
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   31
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   31
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   30
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   30
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   29
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   29
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   28
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   28
      Top             =   2880
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   27
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   27
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   26
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   26
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   25
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   25
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   24
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   24
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   23
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   23
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   22
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   22
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   21
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   21
      Top             =   2370
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   20
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   20
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   19
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   19
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   18
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   18
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   17
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   17
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   16
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   16
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   15
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   15
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   14
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   14
      Top             =   1860
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   13
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   13
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   12
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   12
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   11
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   11
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   10
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   10
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   9
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   9
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   8
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   8
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   7
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   7
      Top             =   1350
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   6
      Left            =   7080
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   6
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   5
      Left            =   6060
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   5
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   4
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   3
      Left            =   4020
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   2
      Left            =   3000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   1980
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   1
      Top             =   840
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   0
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Index           =   2
      Left            =   8400
      TabIndex        =   92
      Top             =   5220
      Width           =   285
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Index           =   1
      Left            =   8400
      TabIndex        =   91
      Top             =   3180
      Width           =   285
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Index           =   0
      Left            =   8400
      TabIndex        =   90
      Top             =   1140
      Width           =   285
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   7
      X1              =   450
      X2              =   840
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   6
      X1              =   480
      X2              =   870
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   5
      X1              =   480
      X2              =   870
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   4
      X1              =   480
      X2              =   870
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   3
      X1              =   480
      X2              =   870
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   2
      X1              =   480
      X2              =   870
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Index           =   1
      X1              =   450
      X2              =   840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Index           =   0
      X1              =   480
      X2              =   870
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line15 
      X1              =   720
      X2              =   750
      Y1              =   570
      Y2              =   540
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Index           =   3
      Left            =   240
      Shape           =   2  'Oval
      Top             =   5160
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Index           =   2
      Left            =   240
      Shape           =   2  'Oval
      Top             =   3750
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Index           =   1
      Left            =   240
      Shape           =   2  'Oval
      Top             =   2340
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Index           =   0
      Left            =   240
      Shape           =   2  'Oval
      Top             =   960
      Width           =   405
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   3060
      X2              =   2580
      Y1              =   540
      Y2              =   780
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   4590
      X2              =   3060
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   9240
      X2              =   5340
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   8790
      X2              =   9240
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   8820
      X2              =   9240
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9240
      Y1              =   5400
      Y2              =   540
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   8790
      X2              =   9240
      Y1              =   6180
      Y2              =   5370
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      X1              =   8130
      X2              =   8790
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      X1              =   8130
      X2              =   8790
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      X1              =   8130
      X2              =   8790
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line4 
      BorderWidth     =   10
      X1              =   8130
      X2              =   8130
      Y1              =   810
      Y2              =   7020
   End
   Begin VB.Line Line3 
      BorderWidth     =   10
      X1              =   930
      X2              =   8130
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      X1              =   930
      X2              =   930
      Y1              =   840
      Y2              =   7050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      X1              =   930
      X2              =   8130
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnNewgame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnAnd 
      Caption         =   "AND"
      Begin VB.Menu m1 
         Caption         =   "0 - 0 = 0"
      End
      Begin VB.Menu m2 
         Caption         =   "0 - 1 = 0"
      End
      Begin VB.Menu m3 
         Caption         =   "1 - 0 = 0"
      End
      Begin VB.Menu m4 
         Caption         =   "1 - 1 = 1"
      End
   End
   Begin VB.Menu mnNand 
      Caption         =   "NAND"
      Begin VB.Menu m5 
         Caption         =   "0 - 0 = 1"
      End
      Begin VB.Menu m6 
         Caption         =   "0 - 1 = 1"
      End
      Begin VB.Menu m7 
         Caption         =   "1 - 0 = 1"
      End
      Begin VB.Menu m8 
         Caption         =   "1 - 1 = 0"
      End
   End
   Begin VB.Menu mnOr 
      Caption         =   "OR"
      Begin VB.Menu m9 
         Caption         =   "0 - 0 = 0"
      End
      Begin VB.Menu m10 
         Caption         =   "0 - 1 = 1"
      End
      Begin VB.Menu m11 
         Caption         =   "1 - 0 = 1"
      End
      Begin VB.Menu m12 
         Caption         =   "1 - 1 = 1"
      End
   End
   Begin VB.Menu mnNor 
      Caption         =   "NOR"
      Begin VB.Menu m13 
         Caption         =   "0 - 0 = 1"
      End
      Begin VB.Menu m14 
         Caption         =   "0 - 1 = 0"
      End
      Begin VB.Menu m15 
         Caption         =   "1 - 0 = 0"
      End
      Begin VB.Menu m16 
         Caption         =   "1 - 1 = 0"
      End
   End
   Begin VB.Menu mnXor 
      Caption         =   "XOR"
      Begin VB.Menu m17 
         Caption         =   "0 - 0 = 0"
      End
      Begin VB.Menu m18 
         Caption         =   "0 - 1 = 1"
      End
      Begin VB.Menu m19 
         Caption         =   "1 - 0 = 1"
      End
      Begin VB.Menu m20 
         Caption         =   "1 - 1 = 0"
      End
   End
   Begin VB.Menu mnXnor 
      Caption         =   "XNOR"
      Begin VB.Menu m21 
         Caption         =   "0 - 0 = 1"
      End
      Begin VB.Menu m22 
         Caption         =   "0 - 1 = 0"
      End
      Begin VB.Menu m23 
         Caption         =   "1 - 0 = 0"
      End
      Begin VB.Menu m24 
         Caption         =   "1 - 1 = 1"
      End
   End
   Begin VB.Menu mnBuf 
      Caption         =   "BUF"
      Begin VB.Menu m25 
         Caption         =   "0 = 0  pass through, no invert"
      End
      Begin VB.Menu m26 
         Caption         =   "1 = 1  pass through, no invert"
      End
   End
   Begin VB.Menu mnInv 
      Caption         =   "INV"
      Begin VB.Menu m27 
         Caption         =   "0 = 1  invert"
      End
      Begin VB.Menu m28 
         Caption         =   "1 = 0  invert"
      End
   End
   Begin VB.Menu mnblank 
      Caption         =   "blank"
      Visible         =   0   'False
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
      Begin VB.Menu mnAbout1 
         Caption         =   "By Max Seim   mlseim@mmm.com"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seconds As Integer
Dim game As Integer
Dim temparray(84)
Dim logicarray(84)

Private Sub Command2_Click()
replay
End Sub

Private Sub Form_Load()
main.Top = (Screen.Height - main.Height) / 2 ' center form
main.Left = (Screen.Width - main.Width) / 2 ' center form
For x = 0 To 83 ' set all picture backgrounds to white
Picture1(x).BackColor = vbWhite
Next x
game = 2
seconds = 60
Command2.Visible = False
Label1.Caption = seconds
While logicarray(20) = 0 And logicarray(48) = 0 And logicarray(76) = 0
vcclogic ' load the initial voltage (logic levels)
startlogic ' load the first column of logic gates
connect1 ' connect the logic gate outputs
logic2 ' load the second column of logic gates
connect2 ' connect the logic gate outputs
logic3 ' load the third column of logic gates
logic4 ' load the last column of non-inverting buffers or inverting buffers
Wend
'  Output logic array for troubleshooting
'-------------------------------------------
'c = 0
'Open "c:\test.txt" For Output As 1
'For x = 0 To 83
'Print #1, x; "  "; logicarray(x)
'Next x
'Close
Label5(0).Caption = ""
Label5(1).Caption = ""
Label5(2).Caption = ""
For x = 0 To 83
temparray(x) = logicarray(x)
Next x
End Sub
Private Sub replay()
Erase logicarray
Erase temparray
Label5(0).Caption = ""
Label5(1).Caption = ""
Label5(2).Caption = ""
For x = 0 To 83 ' set all picture backgrounds to white
Picture1(x).BackColor = vbWhite
Next x
game = 2
seconds = 60
For x = 0 To 2
Command1(x).Visible = True
Next x
Command2.Visible = False
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Label1.Caption = seconds
While logicarray(20) = 0 And logicarray(48) = 0 And logicarray(76) = 0
vcclogic ' load the initial voltage (logic levels)
startlogic ' load the first column of logic gates
connect1 ' connect the logic gate outputs
logic2 ' load the second column of logic gates
connect2 ' connect the logic gate outputs
logic3 ' load the third column of logic gates
logic4 ' load the last column of non-inverting buffers or inverting buffers
Wend
For x = 0 To 83
temparray(x) = logicarray(x)
Next x
Timer1.Enabled = True
End Sub
Private Sub vcclogic()
y = 0
For x = 0 To 11
Randomize
r = Int(Rnd(1) * 3 + 1)
If r = 1 Then
Picture1(x * 7).Picture = LoadPicture("vcc1.bmp")
logicarray(x * 7) = 1
End If
If r = 2 Then
Picture1(x * 7).Picture = LoadPicture("vcc2.bmp")
logicarray(x * 7) = 2
End If
If r = 3 Then
Picture1(x * 7).Picture = LoadPicture("vcc3.bmp")
logicarray(x * 7) = 3
End If
Next x
End Sub
Private Sub startlogic()
y = 1
For x = 0 To 11
Randomize
r = Int(Rnd(1) * 6 + 1)
If r = 1 Then
Picture1(x * 7 + y).Picture = LoadPicture("and.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 0
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 0
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 1
         End If
End If
If r = 2 Then
Picture1(x * 7 + y).Picture = LoadPicture("nand.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 1
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 1
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 0
         End If
End If
If r = 3 Then
Picture1(x * 7 + y).Picture = LoadPicture("nor.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 1
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 0
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 0
         End If
End If
If r = 4 Then
Picture1(x * 7 + y).Picture = LoadPicture("or.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 0
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 1
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 1
         End If
End If
If r = 5 Then
Picture1(x * 7 + y).Picture = LoadPicture("xor.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 0
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 1
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 0
         End If
End If
If r = 6 Then
Picture1(x * 7 + y).Picture = LoadPicture("xnor.bmp")
   If logicarray(x * 7) = 1 Then
   logicarray(x * 7 + y) = 1
   End If
      If logicarray(x * 7) = 2 Then
      logicarray(x * 7 + y) = 0
      End If
         If logicarray(x * 7) = 3 Then
         logicarray(x * 7 + y) = 1
         End If
End If
Next x
End Sub
Private Sub connect1()
y = 2
For x = 0 To 11 Step 4
Picture1(x * 7 + y).Picture = LoadPicture("c1.bmp")
Picture1((x + 1) * 7 + y).Picture = LoadPicture("c2.bmp")
Picture1((x + 2) * 7 + y).Picture = LoadPicture("c3.bmp")
Picture1((x + 3) * 7 + y).Picture = LoadPicture("c4.bmp")
Next x
End Sub
Private Sub logic2()
y = 3
For x = 1 To 9 Step 4
  For t = 0 To 1
Randomize
r = Int(Rnd(1) * 6 + 1)
If r = 1 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("and.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
If r = 2 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("nand.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 3 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("nor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 4 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("or.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
If r = 5 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("xor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 6 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("xnor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 1) = 0 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 1) = 1 And logicarray((x - 1) * 7 + (t * 14) + 8) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
  Next t
Next x
End Sub
Private Sub connect2()
y = 4
For x = 1 To 9 Step 4
Picture1((x) * 7 + y).Picture = LoadPicture("c1.bmp")
Picture1((x + 1) * 7 + y).Picture = LoadPicture("c2.bmp")
Next x
End Sub
Private Sub logic3()
y = 5
For x = 2 To 10 Step 4
t = 0
Randomize
r = Int(Rnd(1) * 6 + 1)
If r = 1 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("and.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
If r = 2 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("nand.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 3 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("nor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 4 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("or.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
If r = 5 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("xor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 1
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 1
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 0
            End If
End If
If r = 6 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("xnor.bmp")
   If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 0 Then
      logicarray((x + t) * 7 + y) = 0
      End If
         If logicarray((x - 1) * 7 + (t * 14) + 3) = 0 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
         logicarray((x + t) * 7 + y) = 0
         End If
            If logicarray((x - 1) * 7 + (t * 14) + 3) = 1 And logicarray((x - 1) * 7 + (t * 14) + 10) = 1 Then
            logicarray((x + t) * 7 + y) = 1
            End If
End If
Next x
End Sub
Private Sub logic4()
y = 6
For x = 2 To 10 Step 4
t = 0
Randomize
r = Int(Rnd(1) * 2 + 1)
If r = 1 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("inv.bmp")
   If logicarray((x + t) * 7 + y - 1) = 0 Then
   logicarray((x + t) * 7 + y) = 1
   End If
      If logicarray((x + t) * 7 + y - 1) = 1 Then
      logicarray((x + t) * 7 + y) = 0
      End If
End If
If r = 2 Then
Picture1((x + t) * 7 + y).Picture = LoadPicture("buf.bmp")
   If logicarray((x + t) * 7 + y - 1) = 0 Then
   logicarray((x + t) * 7 + y) = 0
   End If
      If logicarray((x + t) * 7 + y - 1) = 1 Then
      logicarray((x + t) * 7 + y) = 1
      End If
End If
Next x
End Sub

Private Sub mnExit_Click()
Erase logicarray
Unload Me
End Sub
Private Sub Command1_Click(Index As Integer)
'
If Index = 0 And logicarray(20) = 0 Then
game = 0 ' player loses
End If
   If Index = 1 And logicarray(48) = 0 Then
   game = 0 ' player loses
   End If
      If Index = 2 And logicarray(76) = 0 Then
      game = 0 ' player loses
      End If
If Index = 0 And logicarray(20) = 1 Then
Line5.Visible = False
Command1(0).Visible = False
temparray(20) = 0
End If
   If Index = 1 And logicarray(48) = 1 Then
   Line6.Visible = False
   Command1(1).Visible = False
   temparray(48) = 0
   End If
      If Index = 2 And logicarray(76) = 1 Then
      Line7.Visible = False
      Command1(2).Visible = False
      temparray(76) = 0
      End If
         If temparray(20) = 0 And temparray(48) = 0 And temparray(76) = 0 Then
         game = 1 ' player wins!
         End If
End Sub

Private Sub mnNewgame_Click()
replay
End Sub

Private Sub Timer1_Timer()
seconds = seconds - 1
Label1.Caption = "  "
Label1.Caption = seconds
If seconds = 0 Or game = 0 Then
  For x = 0 To 83
  Picture1(x).BackColor = vbRed
  Next x
  Timer1.Enabled = False
Command2.Caption = "KABOOM!!! You LOSE!!!  Click Me"
Command2.Visible = True
Label5(0).Caption = logicarray(20)
Label5(1).Caption = logicarray(48)
Label5(2).Caption = logicarray(76)
End If

If game = 1 Then
  Timer1.Enabled = False
Command2.Caption = "You Defused the Bomb!!   Click Me"
Command2.Visible = True
Label5(0).Caption = logicarray(20)
Label5(1).Caption = logicarray(48)
Label5(2).Caption = logicarray(76)
End If
End Sub
