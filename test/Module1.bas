Attribute VB_Name = "Module1"
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

Public Function CallMul(ByVal pfn As Long, ByVal lFirst As Long, ByVal lSecond As Long) As Long
    '--- on first call will self-patch to call `pfn` with args `lFirst` and `lSecond` directly
    RtccPatchProto AddressOf Module1.CallMul
    CallMul = CallMul(pfn, lFirst, lSecond)
End Function

