## VbRtcc
Runtime Tiny C Compilier for VB6

### Description

VbRtcc is a fork of OTCC by Fabrice Bellard with additional tweaks.

### Sample usage

```
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
    
    RtccFree m_ctx
End Sub
```
All functions are compiled with stdcall calling convention and unresolved externals like `GlobalAlloc` and `GetVersionExA` are searched in currenty loaded DLLs.

Context `m_ctx` is reusable for additional invokations of `RtccCompile` with all previous symbols in scope, so previously compiled functions can be called from subsequent compilations.

Returned `pfn` are usable until `m_ctx` is not destroyed with `RtccFree`.

### ToDo

 - Support `short` data type for Unicode strings
