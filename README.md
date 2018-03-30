## VbRtcc
Runtime Tiny C Compiler for VB6

### Description

VbRtcc is a fork of [OTCC by Fabrice Bellard](https://bellard.org/otcc/) with some additional tweaks -- relocatable code, stdcall calling convention, `short *` access, `_asm` for inline assembly, local and global arrays.

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

Returned `pfn` are usable until `m_ctx` is destroyed with `RtccFree`.

### API

 - `RtccCompile([in] ctx, [in] source, [in, optional] allocsize)`

   If `ctx` is new initializes internal buffers according to the optional `allocsize`. Then compiles inline C `source` to `ctx`, appending generated code to previously compiled code in `ctx` and returns a callable `pfn` pointer to the entry point of the first function in `source`.
   
 - `RtccFree([in] ctx)`
 
   Deallocates `ctx` internal buffers, rendering previously compiled `pfn`s in `ctx` invalid.
   
 - `RtccGetSymbol([in] ctx, [in] name)`
 
   Searches function `name` in `ctx` and returns a callable `pfn` pointer to its entry point or `0` (`NULL`) if not found. `RtccGetSymbol` can be used when compiling sources with multiple functions that call each other because OTCC can compile calls to forward references e.g. `f` calling `g` and `g` calling `f` in a single source code. 
   
 - `RtccPatchProto([in] addressof)`
 
   Helper function that patches a native VB6 function to become a `pfn` calling trampoline. For instance a VB6 function with declaration `Function MyProto(byval pfn As Long, a1 As Long, a2 As Long, ...) As Long` can be patched with a call to `RtccPatchProto AddressOf MyProto` to jump at run-time to `pfn` with arguments `a1, a2, ...`.
   
   This allows calling `RtccCompile`d `pfn` without resorting to `CallWindowProc` or similar work-arounds that are usually slower that implementing simple trampolines with `RtccPatchProto`.

### Allowed C syntax

See `README` in `lib/ottc` for original Bellard's comments.

 - All local variables are `signed int`, including pointers.
 - Array indexing is not supported but global and local arrays declaration is possible.
 - Three types of dereferencing are allowed - `*(char *)p` for byte, `*(short *)p` for 16-bit word and `*(int *)p` for 32-bit dword.
 - Pointer arithmetic uses byte offset, e.g. use `*(int *)(p + 4)` to get second dword.
 - Shifts are signed e.g. `i = i >> 8` extends the sign bit.
 - Function retval/parameters are always `int` (no types allowed) e.g. `main(n, t)` returns `int` and has two `int` parameters.
 - Char literals are cast to dword e.g. `v = 'a'` is sign-extended
 - Char/strings escape allows hex values e.g. `"\x12\x5f"` and `\r`, `\n` and `\t`
 - Unicode char/strings allowed with `L"string"` and `L'a'` and hex word escapes with `L"\x1234\xabff"`
 - Supported statements: `if`/`else` `while` `for` `break` `return`
 - Supported operators: `+` `-` `*` `/` `%` `!` `~` `>>` `<<` `|` `&` `^` `++` `--` `&&` `||` `==` `!=` `<` `>` `<=` `>=`
 - Not supported operators: `+=` (and family) `?:` (ternary) `,` (comma in for loops too)
 - Block comments only (no `//` line comments)
 - `#define` supports constants only (no macro expansion)
 - Function `alloca` can be used for stack allocation of local arrays
 - Only global arrays declaration is supported w/ size in bytes e.g. `int a[100]` reserves 100 bytes (from `ctx->glob`)
 - `_asm _emit <num>` allows encoding custom instructions in codegen
 - `_asm mov eax, <expr>` supported only, where `<expr>` can be const/var or complex expression
 
### ToDo

 - [x] Support `short` data type (for Unicode strings)
 - [x] Support inline `_asm _emit`
 - [x] Implement get symbol address after compile
