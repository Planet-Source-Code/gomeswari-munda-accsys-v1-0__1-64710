VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBalanceSheet 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   Icon            =   "frmBalanceSheet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   8805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   360
      TabIndex        =   61
      Top             =   240
      Width           =   11895
      Begin VB.CommandButton cmdNxtIncIE 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   242
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrevIncIE 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   241
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdNxtExpIE 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   240
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrevExpIE 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   239
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   6855
         Left            =   11040
         TabIndex        =   135
         Top             =   -3840
         Width           =   11895
         Begin VB.TextBox Text25 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10800
            TabIndex        =   238
            Top             =   6240
            Width           =   870
         End
         Begin VB.TextBox Text24 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7560
            TabIndex        =   236
            Top             =   6240
            Width           =   2295
         End
         Begin VB.TextBox Text23 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   234
            Top             =   6240
            Width           =   870
         End
         Begin VB.TextBox Text22 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   232
            Top             =   6240
            Width           =   2310
         End
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   10800
            TabIndex        =   230
            Top             =   5520
            Width           =   870
         End
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   10800
            TabIndex        =   229
            Top             =   4920
            Width           =   870
         End
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   10800
            TabIndex        =   228
            Top             =   4320
            Width           =   870
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   7560
            TabIndex        =   224
            Top             =   5520
            Width           =   2295
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   7560
            TabIndex        =   223
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7560
            TabIndex        =   222
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   10800
            TabIndex        =   218
            Top             =   3120
            Width           =   870
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   10800
            TabIndex        =   217
            Top             =   2520
            Width           =   870
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   10800
            TabIndex        =   216
            Top             =   1920
            Width           =   870
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   7560
            TabIndex        =   212
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   7560
            TabIndex        =   211
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7560
            TabIndex        =   210
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   10800
            TabIndex        =   209
            Top             =   3720
            Width           =   870
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   7560
            TabIndex        =   207
            Top             =   3720
            Width           =   2295
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   10800
            TabIndex        =   202
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   7560
            TabIndex        =   200
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4800
            TabIndex        =   198
            Top             =   5520
            Width           =   855
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4800
            TabIndex        =   197
            Top             =   4920
            Width           =   855
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4800
            TabIndex        =   196
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   192
            Top             =   5520
            Width           =   2295
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   191
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   190
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   186
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   185
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   184
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4800
            TabIndex        =   183
            Top             =   3720
            Width           =   855
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   181
            Top             =   3720
            Width           =   2295
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4800
            TabIndex        =   179
            Top             =   3120
            Width           =   870
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4800
            TabIndex        =   178
            Top             =   2520
            Width           =   870
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4800
            TabIndex        =   177
            Top             =   1920
            Width           =   870
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4800
            TabIndex        =   170
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   168
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10800
            TabIndex        =   166
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7560
            TabIndex        =   164
            Text            =   "--------------Select one------------"
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   162
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   160
            Text            =   "--------------Select one------------"
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9960
            TabIndex        =   237
            Top             =   6240
            Width           =   810
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6240
            TabIndex        =   235
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   233
            Top             =   6240
            Width           =   810
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   231
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   9960
            TabIndex        =   227
            Top             =   5520
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   9960
            TabIndex        =   226
            Top             =   4920
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   9960
            TabIndex        =   225
            Top             =   4320
            Width           =   810
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   7080
            TabIndex        =   221
            Top             =   5520
            Width           =   420
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   7080
            TabIndex        =   220
            Top             =   4920
            Width           =   420
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   7080
            TabIndex        =   219
            Top             =   4320
            Width           =   420
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   9960
            TabIndex        =   215
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   9960
            TabIndex        =   214
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   9960
            TabIndex        =   213
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   9960
            TabIndex        =   208
            Top             =   3720
            Width           =   810
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   7080
            TabIndex        =   206
            Top             =   3720
            Width           =   420
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   7080
            TabIndex        =   205
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   7080
            TabIndex        =   204
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   7080
            TabIndex        =   203
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   9960
            TabIndex        =   201
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   7080
            TabIndex        =   199
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3960
            TabIndex        =   195
            Top             =   5520
            Width           =   810
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   3960
            TabIndex        =   194
            Top             =   4920
            Width           =   810
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3960
            TabIndex        =   193
            Top             =   4320
            Width           =   810
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   189
            Top             =   5520
            Width           =   420
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   188
            Top             =   4920
            Width           =   420
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   187
            Top             =   4320
            Width           =   420
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3960
            TabIndex        =   182
            Top             =   3720
            Width           =   810
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Less:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   180
            Top             =   3720
            Width           =   420
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3960
            TabIndex        =   176
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   3960
            TabIndex        =   175
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3960
            TabIndex        =   174
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   173
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   172
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   171
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3960
            TabIndex        =   169
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   167
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9960
            TabIndex        =   165
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6240
            TabIndex        =   163
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Amt.(Rs.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   161
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   159
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label36 
            Height          =   480
            Index           =   19
            Left            =   6120
            TabIndex        =   158
            Top             =   6120
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   18
            Left            =   6120
            TabIndex        =   157
            Top             =   5400
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   17
            Left            =   6120
            TabIndex        =   156
            Top             =   4800
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   16
            Left            =   6120
            TabIndex        =   155
            Top             =   4200
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   15
            Left            =   6120
            TabIndex        =   154
            Top             =   3600
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   14
            Left            =   6120
            TabIndex        =   153
            Top             =   3000
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   13
            Left            =   6120
            TabIndex        =   152
            Top             =   2400
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   12
            Left            =   6120
            TabIndex        =   151
            Top             =   1800
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   11
            Left            =   6120
            TabIndex        =   150
            Top             =   1200
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00808080&
            Height          =   480
            Index           =   10
            Left            =   6120
            TabIndex        =   149
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label36 
            Height          =   480
            Index           =   9
            Left            =   120
            TabIndex        =   148
            Top             =   6120
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   8
            Left            =   120
            TabIndex        =   147
            Top             =   5400
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   7
            Left            =   120
            TabIndex        =   146
            Top             =   4800
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   6
            Left            =   120
            TabIndex        =   145
            Top             =   4200
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   5
            Left            =   120
            TabIndex        =   144
            Top             =   3600
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   4
            Left            =   120
            TabIndex        =   143
            Top             =   3000
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   3
            Left            =   120
            TabIndex        =   142
            Top             =   2400
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   2
            Left            =   120
            TabIndex        =   141
            Top             =   1800
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00E0E0E0&
            Height          =   480
            Index           =   1
            Left            =   120
            TabIndex        =   140
            Top             =   1200
            Width           =   5655
         End
         Begin VB.Label Label36 
            BackColor       =   &H00808080&
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   139
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Assets"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8520
            TabIndex        =   138
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Liabilities"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2040
            TabIndex        =   137
            Top             =   120
            Width           =   885
         End
         Begin VB.Line Line1 
            BorderWidth     =   5
            Index           =   1
            X1              =   6000
            X2              =   6000
            Y1              =   120
            Y2              =   6240
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0C000&
            Height          =   495
            Left            =   0
            TabIndex        =   136
            Top             =   0
            Width           =   11895
         End
      End
      Begin VB.ComboBox cmbAcNameExp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   134
         Text            =   "--------------Select one------------"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAdAmtIn 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   6120
         Width           =   855
      End
      Begin VB.TextBox txtANameInc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtAdAmtExp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   6120
         Width           =   855
      End
      Begin VB.TextBox txtANameExp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   6120
         Width           =   2295
      End
      Begin VB.ComboBox cmbAcNameInc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   121
         Text            =   "--------------Select one------------"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Click for:      Income and Expenditure A/c"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7080
         Width           =   11655
      End
      Begin VB.CommandButton cmdEdInIE 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSvInIE 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalInIE 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdEdExpIE 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSvExpIE 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCalExpIE 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6600
         Width           =   1050
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   10800
         TabIndex        =   34
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   10800
         TabIndex        =   32
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   10800
         TabIndex        =   30
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   10800
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7560
         TabIndex        =   33
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   31
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7560
         TabIndex        =   29
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   27
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   10800
         TabIndex        =   26
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   10800
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   10800
         TabIndex        =   22
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   10800
         TabIndex        =   20
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7560
         TabIndex        =   25
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   23
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7560
         TabIndex        =   21
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtAcNameInc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   15
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   13
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   9
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   14
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   7
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   5
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   3
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtAcNameExp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   600
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   150
         Left            =   11640
         Picture         =   "frmBalanceSheet.frx":030A
         Top             =   5880
         Width           =   120
      End
      Begin VB.Image Image1 
         Height          =   150
         Left            =   5640
         Picture         =   "frmBalanceSheet.frx":036B
         Top             =   5760
         Width           =   120
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adjusted Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   9960
         TabIndex        =   133
         Top             =   5760
         Width           =   1650
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9960
         TabIndex        =   131
         Top             =   6240
         Width           =   810
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   129
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   127
         Top             =   6240
         Width           =   810
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adjusted Total Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   3960
         TabIndex        =   126
         Top             =   5760
         Width           =   1650
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   124
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label27 
         Height          =   480
         Index           =   1
         Left            =   6120
         TabIndex        =   123
         Top             =   6000
         Width           =   5655
      End
      Begin VB.Label Label27 
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   122
         Top             =   6000
         Width           =   5655
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   9960
         TabIndex        =   120
         Top             =   5520
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   9960
         TabIndex        =   119
         Top             =   4920
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   9960
         TabIndex        =   118
         Top             =   4320
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   9960
         TabIndex        =   117
         Top             =   3720
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   9960
         TabIndex        =   116
         Top             =   3120
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   9960
         TabIndex        =   115
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   9960
         TabIndex        =   114
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   9960
         TabIndex        =   113
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   7080
         TabIndex        =   112
         Top             =   5520
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   7080
         TabIndex        =   111
         Top             =   4920
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   7080
         TabIndex        =   110
         Top             =   4320
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   7080
         TabIndex        =   109
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9960
         TabIndex        =   107
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   106
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   105
         Top             =   5520
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   104
         Top             =   4920
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   103
         Top             =   4320
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3960
         TabIndex        =   102
         Top             =   3720
         Width           =   810
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   101
         Top             =   5520
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   100
         Top             =   4920
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   99
         Top             =   4320
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   98
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   97
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   96
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   95
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   94
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   7080
         TabIndex        =   93
         Top             =   3120
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   7080
         TabIndex        =   92
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   7080
         TabIndex        =   91
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   7080
         TabIndex        =   90
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   89
         Top             =   3120
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   88
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   87
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   86
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   17
         Left            =   6120
         TabIndex        =   84
         Top             =   5280
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   16
         Left            =   6120
         TabIndex        =   83
         Top             =   4680
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   15
         Left            =   6120
         TabIndex        =   82
         Top             =   4080
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   14
         Left            =   6120
         TabIndex        =   81
         Top             =   3480
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   13
         Left            =   6120
         TabIndex        =   80
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   12
         Left            =   6120
         TabIndex        =   79
         Top             =   2280
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   11
         Left            =   6120
         TabIndex        =   78
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   10
         Left            =   6120
         TabIndex        =   77
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00808080&
         Height          =   480
         Index           =   9
         Left            =   6120
         TabIndex        =   76
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   8
         Left            =   120
         TabIndex        =   75
         Top             =   5280
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   7
         Left            =   120
         TabIndex        =   74
         Top             =   4680
         Width           =   5655
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Amt.(Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   73
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   4080
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   5
         Left            =   120
         TabIndex        =   70
         Top             =   3480
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   4
         Left            =   120
         TabIndex        =   69
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   3
         Left            =   120
         TabIndex        =   68
         Top             =   2280
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   67
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00808080&
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   480
         Width           =   5655
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   0
         X1              =   6000
         X2              =   6000
         Y1              =   0
         Y2              =   6120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Income"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8520
         TabIndex        =   64
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Expenditure"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   63
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C000&
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Frame FrameBS 
      BackColor       =   &H00E0E0E0&
      Height          =   6495
      Left            =   9000
      TabIndex        =   45
      Top             =   6720
      Width           =   11415
      Begin VB.PictureBox PictureBS 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   11385
         TabIndex        =   46
         Top             =   0
         Width           =   11415
         Begin VB.Label Label12 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4800
            TabIndex        =   49
            Top             =   120
            Width           =   5535
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Balance Sheet of"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2880
            TabIndex        =   48
            Top             =   120
            Width           =   1785
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Balance Sheet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   240
            TabIndex        =   47
            Top             =   120
            Width           =   1530
         End
      End
   End
   Begin VB.Frame FrameIE 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7695
      Left            =   360
      TabIndex        =   40
      Top             =   240
      Width           =   11415
      Begin VB.PictureBox PictureIE 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   11385
         TabIndex        =   42
         Top             =   0
         Width           =   11415
         Begin VB.Label Label9 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3720
            TabIndex        =   44
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Income and Expenditure A/c"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   2775
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridIE 
         Height          =   5535
         Left            =   360
         TabIndex        =   41
         Top             =   1680
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   $"frmBalanceSheet.frx":03CC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   243
         Top             =   1320
         Width           =   10575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   7695
      Left            =   360
      TabIndex        =   39
      Top             =   240
      Width           =   11415
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   4080
         TabIndex        =   60
         Top             =   6240
         Width           =   1575
      End
      Begin VB.TextBox txtPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtRecp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   5400
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Return to main"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   6360
         Width           =   1455
      End
      Begin VB.PictureBox PictureRP 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   11385
         TabIndex        =   51
         Top             =   0
         Width           =   11415
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "Receipts and Payment A/c"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label lblRP 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   52
            Top             =   120
            Width           =   4815
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridRP 
         Height          =   3735
         Left            =   1080
         TabIndex        =   50
         Top             =   1560
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   4
         Cols            =   4
         FixedCols       =   0
         BackColor       =   14209995
         ForeColor       =   0
         BackColorFixed  =   -2147483638
         BackColorSel    =   0
         BackColorBkg    =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL DIFFERENCE (Rs.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1080
         TabIndex        =   59
         Top             =   6360
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL PAYMENTS->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5880
         TabIndex        =   58
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL RECEIPTS->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1320
         TabIndex        =   57
         Top             =   5400
         Width           =   2655
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "OUTSTANDING calculation form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   5160
      TabIndex        =   244
      Top             =   0
      Width           =   2910
   End
End
Attribute VB_Name = "frmBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------mundSoft Technologies product----------------------------
'Programmer-@gomesh
'Year-2005
'E-mail:(i)  g_munda@rediffmail.com
'       (ii) gomesh_p@yahoo.co.in
'Copyright (c) mundSoft Technologies -- All Rights Reserved
'-----------------------------------------------------------------------
Option Explicit
Dim ADr1                As Double
Dim ACr1                As Double
Dim ADr2                As Double
Dim ACr2                As Double
Dim DifDrCr1            As Double
Dim DifDrCr2            As Double
Dim str                 As Integer
Dim res                 As ADODB.Recordset
Dim resVch              As ADODB.Recordset
Dim resBS               As ADODB.Recordset
Dim resIE               As ADODB.Recordset
Dim resIE1              As ADODB.Recordset
Dim resIE2              As ADODB.Recordset
Dim recNo               As Integer
Dim i                   As Integer
Dim j                   As Integer
Dim st                  As String
Dim picPath             As String

Private Sub cmbAcNameExp_Click()
    Call CalIE1
    txtAcNameExp.Text = DifDrCr1
End Sub

Private Sub cmbAcNameInc_Click()
    Call CalIE2
    txtAcNameInc.Text = DifDrCr2
End Sub

Private Sub cmdCalExpIE_Click()
Dim v1, v2, v3, v4, v5, v6, v7, v8, v9, v10 As Double
    v1 = Val(Trim(txtAcNameExp.Text))
    v2 = Val(Trim(Text5(0).Text))
    v3 = Val(Trim(Text5(1).Text))
    v4 = Val(Trim(Text5(2).Text))
    v5 = Val(Trim(Text5(3).Text))
    v6 = Val(Trim(Text7(0).Text))
    v7 = Val(Trim(Text7(1).Text))
    v8 = Val(Trim(Text7(2).Text))
    v9 = Val(Trim(Text7(3).Text))
    v10 = CDbl(v1 + v2 + v3 + v4 + v5 - v6 - v7 - v8 - v9)
    txtAdAmtExp.Text = v10
    txtANameExp.Text = cmbAcNameExp.Text
    
    cmdSvExpIE.Enabled = True
    cmdCalExpIE.Enabled = False
End Sub

Private Sub cmdCalInIE_Click()
Dim u1, u2, u3, u4, u5, u6, u7, u8, u9, u10 As Double
    u1 = Val(Trim(txtAcNameInc.Text))
    u2 = Val(Trim(Text11(0).Text))
    u3 = Val(Trim(Text11(1).Text))
    u4 = Val(Trim(Text11(2).Text))
    u5 = Val(Trim(Text11(3).Text))
    u6 = Val(Trim(Text13(0).Text))
    u7 = Val(Trim(Text13(1).Text))
    u8 = Val(Trim(Text13(2).Text))
    u9 = Val(Trim(Text13(3).Text))
    u10 = CDbl(u1 + u2 + u3 + u4 + u5 - u6 - u7 - u8 - u9)
    txtAdAmtIn.Text = u10
    txtANameInc.Text = cmbAcNameInc.Text
    
    cmdSvInIE.Enabled = True
    cmdCalInIE.Enabled = False
End Sub

Private Sub cmdEdExpIE_Click()
    cmdPrevExpIE.Enabled = True
    cmdNxtExpIE.Enabled = True
End Sub

Private Sub cmdEdInIE_Click()
    cmdPrevIncIE.Enabled = True
    cmdNxtIncIE.Enabled = True
End Sub

Private Sub cmdNxtIncIE_Click()
    If resIE2.RecordCount = 0 Then
       MsgBox "No Records !"
    Exit Sub
    End If
    If Not resIE2.EOF Then
        DisplayincomeIE
       resIE2.MoveNext
    End If
    If resIE2.EOF Then
        resIE2.MoveLast
    End If
        DisplayincomeIE
End Sub

Private Sub cmdPrevIncIE_Click()
    If resIE2.RecordCount = 0 Then
       MsgBox "No Records !"
    Exit Sub
    End If
    If Not resIE2.BOF Then
        DisplayincomeIE
       resIE2.MovePrevious
    End If
    If resIE2.BOF Then
       resIE2.MoveFirst
    End If
        DisplayincomeIE
End Sub

Private Sub cmdSvExpIE_Click()
On Error Resume Next
    resIE1.AddNew
    resIE1!AName = cmbAcNameExp.Text
    resIE1!OrAmt = txtAcNameExp.Text
    resIE1!Add1 = Text4(0).Text
    resIE1!Amt1 = Text5(0).Text
    resIE1!Add2 = Text4(1).Text
    resIE1!Amt2 = Text5(1).Text
    resIE1!Add3 = Text4(2).Text
    resIE1!Amt3 = Text5(2).Text
    resIE1!Add4 = Text4(3).Text
    resIE1!Amt4 = Text5(3).Text
    resIE1!Less1 = Text6(0).Text
    resIE1!LAmt1 = Text7(0).Text
    resIE1!Less2 = Text6(1).Text
    resIE1!LAmt2 = Text7(1).Text
    resIE1!Less3 = Text6(2).Text
    resIE1!LAmt3 = Text7(2).Text
    resIE1!Less4 = Text6(3).Text
    resIE1!LAmt4 = Text7(3).Text
    resIE1!AdjAmt = txtAdAmtExp.Text
    resIE1.Update
    MsgBox "Data saved successfully!"
    
    cmdCalExpIE.Enabled = True
    cmdSvExpIE.Enabled = False
End Sub

Private Sub cmdSvInIE_Click()
On Error Resume Next
    resIE2.AddNew
    resIE2!AName = cmbAcNameInc.Text
    resIE2!OrAmt = txtAcNameInc.Text
    resIE2!Add1 = Text10(0).Text
    resIE2!Amt1 = Text11(0).Text
    resIE2!Add2 = Text10(1).Text
    resIE2!Amt2 = Text11(1).Text
    resIE2!Add3 = Text10(2).Text
    resIE2!Amt3 = Text11(2).Text
    resIE2!Add4 = Text10(3).Text
    resIE2!Amt4 = Text11(3).Text
    resIE2!Less1 = Text12(0).Text
    resIE2!LAmt1 = Text13(0).Text
    resIE2!Less2 = Text12(1).Text
    resIE2!LAmt2 = Text13(1).Text
    resIE2!Less3 = Text12(2).Text
    resIE2!LAmt3 = Text13(2).Text
    resIE2!Less4 = Text12(3).Text
    resIE2!LAmt4 = Text13(3).Text
    resIE2!AdjAmt = txtAdAmtIn.Text
    resIE2.Update
    MsgBox "Data saved successfully!"
    
    cmdCalInIE.Enabled = True
    cmdSvInIE.Enabled = False
End Sub

Private Sub Command1_Click()
    frmStarting.Show
    Unload Me
End Sub
'MSFlexGridIE.CellBackColor = &HFFFFC0
Private Sub Command12_Click() 'Caption="Click for:      Income and Expenditure A/c"
Dim ColIE, RowIE As Integer
    For ColIE = 0 To 1 Step 1
        resIE1.MoveFirst
        RowIE = 1
        '-------------------Expenditure Column------------------------------------
        If Not resIE1.EOF And ColIE = 0 And Not resIE1!AName = "NULL" Then
            While Not resIE1.EOF And ColIE = 0
            MSFlexGridIE.Row = RowIE
            MSFlexGridIE.Col = ColIE
            picPath = App.Path
            'Set MSFlexGridIE.CellPicture = .i16X16
            
            Set MSFlexGridIE.CellPicture = LoadPicture(App.Path + "\Images\post.gif")
                MSFlexGridIE.CellBackColor = &HFFC0C0
                If Not resIE1!AName = "" Then
                    MSFlexGridIE.Text = "    " + resIE1!AName
                    RowIE = RowIE + 1
                End If
                

                If Not resIE1!Add1 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Add1
                    RowIE = RowIE + 1
                End If
                
                
                If Not resIE1!Add2 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Add2
                    RowIE = RowIE + 1
                End If
                
                
                If Not resIE1!Add3 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Add3
                    RowIE = RowIE + 1
                End If
                
                If Not resIE1!Add4 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Add4
                    RowIE = RowIE + 1
                End If
                
                
                If Not resIE1!Less1 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Less1
                    RowIE = RowIE + 1
                End If
                

                If Not resIE1!Less2 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Less2
                    RowIE = RowIE + 1
                End If
                

                If Not resIE1!Less3 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Less3
                    RowIE = RowIE + 1
                End If
                

                If Not resIE1!Less4 = "" Then
                    MSFlexGridIE.TextMatrix(RowIE, 0) = "        " + resIE1!Less4
                    RowIE = RowIE + 1
                End If
                

                resIE1.MoveNext
            Wend
                MSFlexGridIE.Refresh
       '----------------(Adjusted)Amount Column-----------------------------------
        ElseIf Not resIE1.EOF And ColIE = 1 Then
            While Not resIE1.EOF And ColIE = 1
            MSFlexGridIE.Row = RowIE
            MSFlexGridIE.Col = ColIE
                If Not resIE1!Amt1 = "" Then
                    MSFlexGridIE.Text = resIE1!Amt1
                    RowIE = RowIE + 1
                End If
                
                If Not resIE1!Amt2 = "" Then
                    MSFlexGridIE.Text = resIE1!Amt2
                    RowIE = RowIE + 1
                End If
                
                If Not resIE1!Amt3 = "" Then
                    MSFlexGridIE.Text = resIE1!Amt3
                    RowIE = RowIE + 1
                End If
                
                If Not resIE1!Amt4 = "" Then
                    MSFlexGridIE.Text = resIE1!Amt4
                    RowIE = RowIE + 1
                End If
                resIE1.MoveNext
            Wend
                MSFlexGridIE.Refresh
        '----------------(Adjusted +/- Original)Amount Column-----------------------------------
        ElseIf Not resIE1.EOF And ColIE = 2 Then
            While Not resIE1.EOF And ColIE = 2
            MSFlexGridIE.Row = RowIE
            MSFlexGridIE.Col = ColIE
                If Not resIE1!AdjAmt = "" Then
                    MSFlexGridIE.Text = resIE1!AdjAmt
                    RowIE = RowIE + 1
                End If
                resIE1.MoveNext
            Wend
                MSFlexGridIE.Refresh
       End If
      ' Next RowIE
    Next ColIE
    Frame1.Visible = True
    Frame2.Visible = False
    Label3.Visible = False
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
    If keyascii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set resVch = New ADODB.Recordset
        resVch.Open "SELECT * from Voucher", con, adOpenKeyset, adLockOptimistic
    Set resIE1 = New ADODB.Recordset
        resIE1.Open "Select * from IEexpenditure", con, adOpenKeyset, adLockOptimistic
    Set resIE2 = New ADODB.Recordset
        resIE2.Open "Select * from IEincome", con, adOpenKeyset, adLockOptimistic
    Set resIE = New ADODB.Recordset
        resIE.Open "SELECT * from Ledger", con, adOpenKeyset, adLockOptimistic

    If resIE.RecordCount > 0 Then
        resIE.MoveLast
        recNo = resIE.RecordCount

        With MSFlexGridIE
            
            .Rows = recNo + 100
        
            .ColWidth(0) = 3000: .ColWidth(1) = 1080: .ColWidth(2) = 1080
            .ColWidth(3) = 3000: .ColWidth(4) = 1080: .ColWidth(5) = 1080
          
            .Row = 0: .Col = 0: .Text = " Expenditure"
            .Row = 0: .Col = 1: .Text = "Amount(Rs.)"
            .Row = 0: .Col = 2: .Text = "Amount(Rs.)"
            .Row = 0: .Col = 3: .Text = " Income"
            .Row = 0: .Col = 4: .Text = "Amount(Rs.)"
            .Row = 0: .Col = 5: .Text = "Amount(Rs.)"
        End With
    End If
    MSFlexGridIE.Clear
    cmdSvExpIE.Enabled = False
    'cmdEdExpIE.Enabled = False
    cmdPrevExpIE.Enabled = False
    cmdNxtExpIE.Enabled = False
    
    cmdSvInIE.Enabled = False
    'cmdEdInIE.Enabled = False
    cmdPrevIncIE.Enabled = False
    cmdNxtIncIE.Enabled = False
End Sub

Private Sub MSFlexGridRP_KeyPress(keyascii As Integer)
    If keyascii = 27 Then
        Unload Me
    End If
End Sub

Public Sub CalIE1()
Set resVch = New ADODB.Recordset
    resVch.Open "Select sum([Amt]) as TotalDr from Voucher where LdgName='" & cmbAcNameExp.Text & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
        If Not resVch!totalDr = "Null" Then
            ADr1 = resVch!totalDr
        Else
            ADr1 = 0
        End If
        resVch.MoveNext
      Loop
Set resVch = Nothing
Set resVch = New ADODB.Recordset
    resVch.Open "Select sum ([Amt]) as TotalCr from Voucher where LdgName='" & cmbAcNameExp.Text & "' and Type='Cr.'", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
        If Not resVch!totalCr = "NULL" Then
            ACr1 = resVch!totalCr
        Else
            ACr1 = 0
        End If
        resVch.MoveNext
    Loop
Set resVch = Nothing
DifDrCr1 = ADr1 - ACr1
End Sub

Public Sub CalIE2()

Set resVch = New ADODB.Recordset
    resVch.Open "Select sum([Amt]) as TotalDr from Voucher where LdgName='" & cmbAcNameInc.Text & "'  and Type='Dr.' ", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
        If Not resVch!totalDr = "Null" Then
            ADr2 = resVch!totalDr
        Else
            ADr2 = 0
        End If
        resVch.MoveNext
    Loop

Set resVch = Nothing

Set resVch = New ADODB.Recordset
    resVch.Open "Select sum ([Amt]) as TotalCr from Voucher where LdgName='" & cmbAcNameInc.Text & "' and Type='Cr.'", con, adOpenKeyset, adLockOptimistic
    Do While Not resVch.EOF
        If Not resVch!totalCr = "NULL" Then
            ACr2 = resVch!totalCr
        Else
            ACr2 = 0
        End If
        resVch.MoveNext
    Loop

Set resVch = Nothing
        
        DifDrCr2 = ADr2 - ACr2
End Sub


Public Sub DisplayincomeIE()
     cmbAcNameInc.Text = resIE2!AName
     txtAcNameInc.Text = resIE2!OrAmt
     Text10(0).Text = resIE2!Add1
     Text11(0).Text = resIE2!Amt1
     Text10(1).Text = resIE2!Add2
     Text11(1).Text = resIE2!Amt2
     Text10(2).Text = resIE2!Add3
     Text11(2).Text = resIE2!Amt3
     Text10(3).Text = resIE2!Add4
     Text11(3).Text = resIE2!Amt4
     Text12(0).Text = resIE2!Less1
     Text13(0).Text = resIE2!LAmt1
     Text12(1).Text = resIE2!Less2
     Text13(1).Text = resIE2!LAmt2
     Text12(2).Text = resIE2!Less3
     Text13(2).Text = resIE2!LAmt3
     Text12(3).Text = resIE2!Less4
     Text13(3).Text = resIE2!LAmt4
     txtANameInc.Text = resIE2!AName
     txtAdAmtIn.Text = resIE2!AdjAmt
End Sub
