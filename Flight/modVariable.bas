Attribute VB_Name = "modVariable"
'For windows sound
Declare Function PlayWave Lib "winmm" Alias "sndPlaySoundA" _
(ByVal FileName As String, ByVal flags As Long) As Long

Public Turn As Integer
Public DieNumber As Integer
Public Steps As Integer
Public PlayerColour As Integer     'For One Player Mode
Public PlayerName As String        'For One Player Mode
Public GameMode         As String  'A-One Player B-Four Players S-Simulation
Public blnRandomTurn    As Boolean 'Random turn

Public gintStartColour As Integer   'Start Colour
Public gblnUseName As Boolean       'Turn Use Name?
Public gstrObject As String         'Object Name
Public gstrDefault As String        'Default Player Colour

'Use Constant for Colour Index
'Warning: DO NOT CHANGE THESE VALUES!
Public Const BLUE      As Integer = 1
Public Const GREEN     As Integer = 2
Public Const RED       As Integer = 3
Public Const YELLOW    As Integer = 4

Public Player(1 To 4) As String 'Player Customed Name
Public blnAI(1 To 4) As Boolean 'Is Player an AI?
Public AINum(1 To 4) As Integer 'Assign AI Num to AI Player

'Game Rules
Public blnKickPlane As Boolean
Public blnEnableJump As Boolean
Public blnSixOnlyStart As Boolean
Public blnRollAgain As Boolean
Public blnFirstRound As Boolean 'First round when game starts
