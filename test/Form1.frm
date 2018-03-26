VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6732
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   6732
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   516
      Left            =   4536
      TabIndex        =   1
      Top             =   1176
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   516
      Left            =   4536
      TabIndex        =   0
      Top             =   504
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' Runtime Tiny C Compiler for VB6
'
' Copyright (c) 2018 by wqweto@gmail.com
'
' Thunks based on Obfuscated Tiny C Compiler
' Copyright (C) 2001-2003 Fabrice Bellard
'
'=========================================================================
Option Explicit
DefObj A-Z

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, Optional ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long

Private m_ctx                   As UcsRtccContextType

Private Sub Command1_Click()
    Dim pfn         As Long
    
    pfn = RtccCompile(m_ctx, _
        "main(n, t, wParam, lParam) {" & vbCrLf & _
        "   int v; " & vbCrLf & _
        "   v = GlobalAlloc(0x40, 148); " & vbCrLf & _
        "   *(int *)v = 148; " & vbCrLf & _
        "   GetVersionExA(v); " & vbCrLf & _
        "   return *(int *)(v + 4) * 100 + *(int *)(v + 8); " & vbCrLf & _
        "}")
    Print CallWindowProc(pfn), "&H" & Hex(pfn)
    
    pfn = RtccCompile(m_ctx, _
        "add(a, b, wParam, lParam) {" & vbCrLf & _
        "   return a*b;" & vbCrLf & _
        "}")
    Print CallWindowProc(pfn, 13, 20), "&H" & Hex(pfn)
End Sub

Private Sub Command2_Click()
    RtccFree m_ctx
End Sub
