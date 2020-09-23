VERSION 5.00
Begin VB.Form frmRTTTL 
   Caption         =   "RTTTL Player"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmRTTTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private HalfTones As Collection
Private Scales As Collection
Private strRTTTL As String

Public Sub PlayRTTTL(sRTTTL As String)
    
    strRTTTL = Replace(sRTTTL, " ", "")
    strRTTTL = Replace(strRTTTL, vbCrLf, "")

    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    
    Dim aParts() As String
    aParts = Split(strRTTTL, ":")

    If UBound(aParts) <> 2 Then
        MsgBox "Invalid ringtone"
        Exit Sub
    End If
    
    aParts(2) = Replace(aParts(2), "d", "z", 1, -1, vbTextCompare)
    aParts(2) = Replace(aParts(2), "e", "y", 1, -1, vbTextCompare)
    
    Dim Duration As Integer
    Dim BPM As Integer
    Dim Scle As Integer
    Dim BeatBase As Long
    
    Duration = 4 ' default
    Scle = 6 ' default
    BPM = 62 ' default

    Dim Ctrl, Ctrls() As String, Parts() As String
    Ctrls() = Split(aParts(1), ",")

    For Each Ctrl In Ctrls
        Parts = Split(Ctrl, "=")
        Select Case UCase(Parts(0))
        Case "O"
            Scle = Parts(1)
        Case "B"
            BPM = Parts(1)
        Case "D"
            Duration = Parts(1)
        End Select
    Next

    Dim dDur As Double, sNote As String, iScle As Integer, Special As String
    
    Dim Notes() As String, Note, Tone
    Notes = Split(aParts(2), ",")

    BeatBase = 240000 / BPM

    For Each Note In Notes
        If Note <> "" Then
            
            dDur = Val(Note)
            If dDur = 0 Then
                dDur = Duration
            Else
                Note = Mid(Note, Len(CStr(dDur)) + 1)
            End If
        
            Tone = Left(Note, 1)
        
            If Mid(Note, 2, 1) = "#" Then
                Tone = Tone & "#"
                
                If IsNumeric(Mid(Note, 3, 1)) Then
                    iScle = Mid(Note, 3, 1)
                    Special = Mid(Note, 4, 1)
                Else
                    iScle = Scle
                    Special = Mid(Note, 3, 1)
                End If
            Else
                If IsNumeric(Mid(Note, 2, 1)) Then
                    iScle = Mid(Note, 2, 1)
                    Special = Mid(Note, 3, 1)
                Else
                    iScle = Scle
                    Special = Mid(Note, 2, 1)
                End If
            End If
        
            If Special = "." Then dDur = dDur * 1.5
            
            Debug.Print "D/T/S : " & dDur & "/" & Tone & "/" & iScle & "/" & Special
        
            Dim Freq As Double
            
            If UCase(Tone) = "P" Then
                Freq = 0
            Else
                Freq = Scales(CStr(iScle)) * ((2 ^ (1 / 12)) ^ HalfTones(UCase(Tone)))
            End If
        
            Debug.Print "Freq: " & Freq
        
            dDur = BeatBase / dDur
            
            If UCase(Tone) = "P" Then
                Sleep dDur
            Else
                Beep Freq, dDur
            End If
        
        End If
    Next

End Sub

Private Sub Form_Load()

    Set HalfTones = New Collection
    Set Scales = New Collection

    HalfTones.Add 0, "A"
    HalfTones.Add 1, "A#"
    HalfTones.Add 2, "B"
    HalfTones.Add 3, "C"
    HalfTones.Add 4, "C#"
    HalfTones.Add 5, "Z"  ' D
    HalfTones.Add 6, "Z#" ' D#
    HalfTones.Add 7, "Y"  ' E
    HalfTones.Add 8, "F"
    HalfTones.Add 9, "F#"
    HalfTones.Add 10, "G"
    HalfTones.Add 11, "G#"

    Scales.Add 440, "4"
    Scales.Add 880, "5"
    Scales.Add 1760, "6"
    Scales.Add 3520, "7"

End Sub

