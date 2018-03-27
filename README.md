## VbRtcc
Runtime Tiny C Compilier for VB6

### Description

VbRtcc is a fork of OTCC by Fabrice Bellard with additional tweaks (relocatable code, stdcall calling convention, short pointers, emit keyword).

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
All functions are compiled with stdcall calling convention and unresolved externals like `GlobalAlloc` and `GetVersionExA` are searched in currently loaded DLLs.

Context `m_ctx` is reusable for additional invokations of `RtccCompile` with all previous symbols in scope, so already compiled functions can be called from subsequent compilations.

Returned `pfn` are usable until `m_ctx` is not destroyed with `RtccFree`.

### Allowed syntax

 - All local variables are `signed int`, including pointers.
 - Array indexing is not supported. Cannot declare local arrays too.
 - Three types of dereferencing are allowed - `*(char *)p` for byte, `*(short *)p` for 16-bit word and `*(int *)p` for 32-bit dword.
 - Pointer arithmetic uses byte offset, e.g. use `*(int *)(p + 4)` to get second dword.
 - Shifts are signed e.g. `i = i >> 8` extends the sign bit.
 - Function retval/parameters are always `int` (no types allowed) e.g. `main(n, t)` returns `int` and has two `int` parameters.
 - Char literals are cast to dword e.g. `v = 'a'` is sign-extended
 - Char/strings escape allows hex values e.g. `"\x12\x5f"` and `\r`, `\n` and `\t`
 - Unicode char/strings allowed with `L"string"` and `L'a'` and hex word escapes with `L"\x1234\xabff"`
 - Keyword `emit` allows dword values e.g. `emit(0x12 0x34 0x5678)` with commas being optional
 - Supported statements: `if`/`else` `while` `for` `break` `return`
 - Supported operators: `+` `-` `*` `/` `%` `!` `~` `>>` `<<` `|` `&` `^` `++` `--` `&&` `||` `==` `!=` `<` `>` `<=` `>=`
 - Not supported operators: `+=` (and family) `?:` (ternary)
 - Block comments only (no `//` line comments)
 - `#define` supports constants only (no macro expansion)

### ToDo

 - [x] Support `short` data type for Unicode strings
 - [x] Support `emit`
 - [ ] Impl get symbol address after compile
