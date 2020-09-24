VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 0"
      Height          =   3210
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   630
      Width           =   5490
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1845
         TabIndex        =   10
         Text            =   "Hello World!"
         Top             =   900
         Width           =   1560
      End
   End
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 2"
      Height          =   3210
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Top             =   630
      Width           =   5490
      Begin VB.CommandButton Command1 
         Caption         =   "HELLO WORLD"
         Height          =   330
         Left            =   1500
         TabIndex        =   9
         Top             =   930
         Width           =   1605
      End
   End
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 1"
      Height          =   3210
      Index           =   1
      Left            =   315
      TabIndex        =   3
      Top             =   630
      Width           =   5490
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "HELLO WORLD!!! #1"
         Height          =   195
         Left            =   1470
         TabIndex        =   8
         Top             =   1080
         Width           =   1560
      End
   End
   Begin VB.PictureBox Cover1 
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   285
      ScaleHeight     =   165
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   435
      Width           =   1275
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "General Info"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton CmdPlat 
      Enabled         =   0   'False
      Height          =   3540
      Left            =   195
      TabIndex        =   0
      Top             =   435
      Width           =   5730
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "Colors/Fonts"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "CSS"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   2895
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################################
'#  ATTENTION:  If you like this code PLEASE VOTE       #
'#  FOR ME!  If you have any feedback, we  apreciate it.#
'#  E-mail us at developers@yarinteractive.com          #
'#                                                      #
'#  yar-interactive software is happy to                #
'#  share our knowledge of Visual Basic with you.       #
'#  please visit our website at                         #
'#  http://www.yarinteractive.com                       #
'#                                                      #
'#  There, you can find more great VB resources and     #
'#  great software.                                     #
'#                      - Thanks.                       #
'########################################################

'HOW TO USE THIS...
'First of all, you should run this code to get a sence of
'what it does, it is meant to be an alternative to the
'tabed dialog control, for developers that prefer to keep
'there applications as self inclosed as possible
'
'PUTTING CONTROLS ON THE TAB-DIALOG...
'Obviously, you can place controls in the frame that
'you can see, but how do you place controls in
'another tab?...
'To place a control on the tab dialog in this example
'select OwnerFrame(2) from the PROPERTIES WINDOW.
'Move into design mode and you should see 8 highlighted boxes
'along the edges of a frame, Right-Click on one of these
'boxes, and select 'Bring to Front'. The frame that will appear
'when the button Captioned 'Tab3'. NOTE that you can change
'the caption of this frame, ect. by selecting it from the
'properties window without having to bring it to
'the front.

Public Sub CmdTab_Click(Index As Integer)

'Brings the platform to the front to cover edges of
'unselected buttons. (CmdPlat is a button behind all the
'frames, and tabs to make a 'raised platform effect'
CmdPlat.ZOrder 0

'Bring the frame associated with the pressed button
'to the front.
OwnerFrame(Index).ZOrder 0

'Bring the Cover over the bottom edge of the
'selected button, the make it look like
'it is 'attached' to the rest of the
' "raised platform"
Cover1.Left = CmdTab(Index).Left

'Adjust the size of the cover to cover the buttons bottom
Cover1.Width = CmdTab(Index).Width - 30

'The cover that hides the edge of the selected
'button got hidden by the big "Raised platform"
'we need it to be on top, so bring it to the fron.
Cover1.ZOrder 0

'GO TO http://www.yarinteractive.com for more great VB codes
'and resources

'Get rid of the focus box that is displayed on the button
CmdTab(0).Enabled = False
CmdTab(0).Enabled = True
CmdTab(1).Enabled = False
CmdTab(1).Enabled = True
CmdTab(2).Enabled = False
CmdTab(2).Enabled = True

'Make Sure all the buttons have Bold set to false
CmdTab(0).FontBold = False
CmdTab(1).FontBold = False
CmdTab(2).FontBold = False

'Make the clicked button BOLD
CmdTab(Index).FontBold = True
End Sub

Private Sub Form_Load()
'Ensures that no matter what OwnerFrame you had at front
'In design time, OwnerFrame(0) is the starting frame,
'because CmdTab(0) is the starting Tab.
OwnerFrame(0).ZOrder 0
Cover1.ZOrder 0
Cover1.Left = CmdTab(0).Left

Call CmdTab_Click(0)
End Sub
