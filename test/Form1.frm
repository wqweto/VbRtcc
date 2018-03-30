VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2592
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6732
   LinkTopic       =   "Form1"
   ScaleHeight     =   2592
   ScaleWidth      =   6732
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   516
      Left            =   4536
      TabIndex        =   2
      Top             =   1848
      Width           =   1860
   End
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
    Debug.Assert RtccGetSymbol(m_ctx, "main") = pfn
    
    Print CallWindowProc(pfn), "&H" & Hex(pfn)
    
    pfn = RtccCompile(m_ctx, _
        "mul(a, b) {" & vbCrLf & _
        "   return a*b;" & vbCrLf & _
        "}")
    Debug.Assert RtccGetSymbol(m_ctx, "mul") = pfn
        
    Print CallMul(pfn, 13, 20), "&H" & Hex(pfn)
End Sub

Private Sub Command2_Click()
    RtccFree m_ctx
End Sub

Private Sub Command3_Click()
    Dim lIdx            As Long
    Dim sExpr           As String
    Dim dblTimer        As Double
    Dim dblResult       As Double
    
    sExpr = "(3.5) + 2.9*(2+(1+2))"
    dblTimer = Timer
    For lIdx = 1 To 1000000
        dblResult = SimpleEval(sExpr)
    Next
    Print "SimpleEval: " & dblResult, Format(Timer - dblTimer, "0.000")
End Sub

Private Function SimpleEval(sText As String) As Double
    Dim src         As String
    
    If m_ctx.m_state(31) = 0 Then
        src = src & _
            "#define TOK_FINAL     0" & vbCrLf & _
            "#define TOK_RPAREN    1" & vbCrLf & _
            "#define TOK_ADD       2" & vbCrLf & _
            "#define TOK_MOD       3" & vbCrLf & _
            "#define TOK_IDIV      4" & vbCrLf & _
            "#define TOK_MUL       5" & vbCrLf & _
            "#define TOK_UNARY     6" & vbCrLf & _
            "#define TOK_POWER     7" & vbCrLf & _
            "#define TOK_LPAREN    8" & vbCrLf & _
            "#define TOK_NUM       9" & vbCrLf & _
            "#define TOK_WHITE     10" & vbCrLf & _
            "" & vbCrLf & _
            "int lookup[256];" & vbCrLf & _
            "" & vbCrLf & _
            "simple_eval(s, pdbl, wParam, lParam)" & vbCrLf & _
            "{" & vbCrLf & _
            "    int i, p, l, ch, prec, prev_pr;" & vbCrLf & _
            "    int op_stack, op_idx;" & vbCrLf & _
            "    int val_stack, val_idx;" & vbCrLf
        src = src & _
            "    int num_size;" & vbCrLf & _
            "" & vbCrLf & _
            "    op_idx = op_stack = alloca(4000);" & vbCrLf & _
            "    val_idx = val_stack = alloca(8000);" & vbCrLf & _
            "    l = &lookup;" & vbCrLf & _
            "    if (*(char *)(l + 32) == 0) {" & vbCrLf & _
            "        p = l;" & vbCrLf & _
            "        i = 0;" & vbCrLf & _
            "        while (i < 256) {" & vbCrLf & _
            "            *(char *)p++ = TOK_WHITE;" & vbCrLf & _
            "            i++;" & vbCrLf & _
            "        }" & vbCrLf & _
            "        *(char *)(l + '(') = TOK_LPAREN;" & vbCrLf & _
            "        *(char *)(l + ')') = TOK_RPAREN;" & vbCrLf & _
            "        *(char *)(l + '+') = TOK_ADD;" & vbCrLf & _
            "        *(char *)(l + '-') = TOK_ADD;" & vbCrLf & _
            "        *(char *)(l + '*') = TOK_MUL;" & vbCrLf & _
            "        *(char *)(l + '/') = TOK_MUL;" & vbCrLf & _
            "        *(char *)(l + '^') = TOK_POWER;" & vbCrLf & _
            "        *(char *)(l + '\\') = TOK_IDIV;" & vbCrLf & _
            "        *(char *)(l + '%') = TOK_MOD;" & vbCrLf & _
            "        *(char *)(l + '.') = TOK_NUM;" & vbCrLf
        src = src & _
            "        p = l + '0';" & vbCrLf & _
            "        i = '0';" & vbCrLf & _
            "        while (i <= '9') {" & vbCrLf & _
            "            *(char *)p++ = TOK_NUM;" & vbCrLf & _
            "            i++;" & vbCrLf & _
            "        }" & vbCrLf & _
            "    }" & vbCrLf & _
            "    prev_pr = 0;" & vbCrLf & _
            "    p = s;" & vbCrLf & _
            "    while ((ch = *(short *)p)) {" & vbCrLf & _
            "        if (!(ch >> 8)) {" & vbCrLf & _
            "            prec = *(char *)(l + ch);" & vbCrLf & _
            "            if (prec != TOK_WHITE) {" & vbCrLf & _
            "                if (prec == TOK_NUM) {" & vbCrLf & _
            "                    val_idx = val_idx + 8;" & vbCrLf & _
            "                    parse_num(p, val_idx, &num_size);" & vbCrLf & _
            "                    p = p + ((num_size-1) << 1);" & vbCrLf & _
            "                } else if (prec == TOK_ADD) {" & vbCrLf & _
            "                    if (prev_pr >= TOK_ADD && prev_pr < TOK_NUM)" & vbCrLf & _
            "                        prec = TOK_UNARY;" & vbCrLf
        src = src & _
            "                }" & vbCrLf & _
            "                if (prec >= TOK_ADD && prec < TOK_NUM) {" & vbCrLf & _
            "                    if(prec != TOK_UNARY)" & vbCrLf & _
            "                        eval_stack(prec, op_stack, &op_idx, val_stack, &val_idx);" & vbCrLf & _
            "                    op_idx = op_idx + 4;" & vbCrLf & _
            "                    *(int *)op_idx = (prec << 16) + ch;" & vbCrLf & _
            "                }" & vbCrLf & _
            "                prev_pr = prec;" & vbCrLf & _
            "            }" & vbCrLf & _
            "        }" & vbCrLf & _
            "        p++; p++;" & vbCrLf & _
            "    }" & vbCrLf & _
            "    eval_stack(TOK_FINAL, op_stack, &op_idx, val_stack, &val_idx);" & vbCrLf & _
            "    *(int *)pdbl = *(int *)val_idx;" & vbCrLf & _
            "    *(int *)(pdbl + 4) = *(int *)(val_idx + 4);" & vbCrLf & _
            "}" & vbCrLf & _
            "" & vbCrLf
        src = src & _
            "#define ASM_MOV_EAX_    _asm mov eax," & vbCrLf & _
            "#define ASM_ADD_EAX_    _asm _emit 0x83 _asm _emit 0xc0 _asm _emit" & vbCrLf & _
            "#define ASM_SUB_EAX_    _asm _emit 0x83 _asm _emit 0xe8 _asm _emit" & vbCrLf & _
            "#define ASM_FSTP_EAX    _asm _emit 0xdd _asm _emit 0x18" & vbCrLf & _
            "#define ASM_FLD_EAX     _asm _emit 0xdd _asm _emit 0x00" & vbCrLf & _
            "#define ASM_FLD_EAX_    _asm _emit 0xdd _asm _emit 0x40 _asm _emit" & vbCrLf & _
            "#define ASM_FADD_EAX_   _asm _emit 0xdc _asm _emit 0x40 _asm _emit" & vbCrLf & _
            "#define ASM_FSUB_EAX_   _asm _emit 0xdc _asm _emit 0x60 _asm _emit" & vbCrLf & _
            "#define ASM_FMUL_EAX_   _asm _emit 0xdc _asm _emit 0x48 _asm _emit" & vbCrLf & _
            "#define ASM_FDIV_EAX_   _asm _emit 0xdc _asm _emit 0x70 _asm _emit" & vbCrLf & _
            "#define ASM_FCHS        _asm _emit 0xd9 _asm _emit 0xe0" & vbCrLf & _
            "#define ASM_FILD_EAX    _asm _emit 0xdb _asm _emit 0x00" & vbCrLf & _
            "#define ASM_FISTP_EAX   _asm _emit 0xdb _asm _emit 0x18" & vbCrLf & _
            "#define ASM_FYL2X       _asm _emit 0xd9 _asm _emit 0xf1" & vbCrLf & _
            "#define ASM_FLD1        _asm _emit 0xd9 _asm _emit 0xe8" & vbCrLf & _
            "#define ASM_FLD_ST1     _asm _emit 0xd9 _asm _emit 0xc1" & vbCrLf & _
            "#define ASM_FPREM       _asm _emit 0xd9 _asm _emit 0xf8" & vbCrLf & _
            "#define ASM_F2XM1       _asm _emit 0xd9 _asm _emit 0xf0" & vbCrLf & _
            "#define ASM_FADDP_ST1   _asm _emit 0xde _asm _emit 0xc1" & vbCrLf & _
            "#define ASM_FSCALE      _asm _emit 0xd9 _asm _emit 0xfd" & vbCrLf
        src = src & _
            "" & vbCrLf & _
            "eval_stack(prec, op_stack, pop_idx, val_stack, pval_idx)" & vbCrLf & _
            "{" & vbCrLf & _
            "    int op_idx, val_idx, op, t1, pt1, t2, pt2;" & vbCrLf & _
            "" & vbCrLf & _
            "    op_idx = *(int *)pop_idx;" & vbCrLf & _
            "    val_idx = *(int *)pval_idx;" & vbCrLf & _
            "    while (op_idx > op_stack) {" & vbCrLf & _
            "        if (*(int *)(op_idx) < (prec << 16))" & vbCrLf & _
            "            break;" & vbCrLf & _
            "        val_idx = val_idx - 8;" & vbCrLf & _
            "        op = *(short *)op_idx;" & vbCrLf & _
            "        if (op == '+') {" & vbCrLf & _
            "            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {" & vbCrLf & _
            "                val_idx = val_idx + 8;" & vbCrLf & _
            "            } else {" & vbCrLf & _
            "                /* *(double *)val_idx = *(double *)val_idx + *(double *)(val_idx + 8); */" & vbCrLf & _
            "                ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "                ASM_FLD_EAX;" & vbCrLf & _
            "                ASM_FADD_EAX_(8);" & vbCrLf & _
            "                ASM_FSTP_EAX;" & vbCrLf & _
            "            }" & vbCrLf
        src = src & _
            "        } else if (op == '-') {" & vbCrLf & _
            "            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {" & vbCrLf & _
            "                val_idx = val_idx + 8;" & vbCrLf & _
            "                /* *(double *)val_idx = -*(double *)val_idx; */" & vbCrLf & _
            "                ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "                ASM_FLD_EAX;" & vbCrLf & _
            "                ASM_FCHS;" & vbCrLf & _
            "                ASM_FSTP_EAX;" & vbCrLf & _
            "            } else {" & vbCrLf & _
            "                /* *(double *)val_idx = *(double *)val_idx - *(double *)(val_idx + 8); */" & vbCrLf & _
            "                ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "                ASM_FLD_EAX;" & vbCrLf & _
            "                ASM_FSUB_EAX_(8);" & vbCrLf & _
            "                ASM_FSTP_EAX;" & vbCrLf & _
            "            }" & vbCrLf & _
            "        } else if (op == '*') {" & vbCrLf & _
            "            /* *(double *)val_idx = *(double *)val_idx * *(double *)(val_idx + 8); */" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf
        src = src & _
            "            ASM_FMUL_EAX_(8);" & vbCrLf & _
            "            ASM_FSTP_EAX;" & vbCrLf & _
            "        } else if (op == '/') {" & vbCrLf & _
            "            /* *(double *)val_idx = *(double *)val_idx / *(double *)(val_idx + 8); */" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_FDIV_EAX_(8);" & vbCrLf & _
            "            ASM_FSTP_EAX;" & vbCrLf & _
            "        } else if (op == '^') {" & vbCrLf & _
            "            /* *(double *)val_idx = pow(*(double *)val_idx, *(double *)(val_idx + 8)); */" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_ADD_EAX_(8);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_SUB_EAX_(8);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_FYL2X;" & vbCrLf & _
            "            ASM_FLD1;" & vbCrLf & _
            "            ASM_FLD_ST1;" & vbCrLf & _
            "            ASM_FPREM;" & vbCrLf & _
            "            ASM_F2XM1;" & vbCrLf & _
            "            ASM_FADDP_ST1;" & vbCrLf & _
            "            ASM_FSCALE;" & vbCrLf & _
            "            ASM_FSTP_EAX;" & vbCrLf
        src = src & _
            "        } else if (op == '\\') {" & vbCrLf & _
            "            pt1 = &t1;" & vbCrLf & _
            "            /* *(double *)val_idx = (int)(*(double *)val_idx / *(double *)(val_idx + 8)); */" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_FDIV_EAX_(8);" & vbCrLf & _
            "            ASM_MOV_EAX_(pt1);" & vbCrLf & _
            "            ASM_FISTP_EAX;" & vbCrLf & _
            "            ASM_FILD_EAX;" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FSTP_EAX;" & vbCrLf & _
            "        } else if (op == '%') {" & vbCrLf & _
            "            pt1 = &t1;" & vbCrLf & _
            "            pt2 = &t2;" & vbCrLf & _
            "            /* *(double *)val_idx = (int)*(double *)val_idx % (int)*(double *)(val_idx + 8); */" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_MOV_EAX_(pt1);" & vbCrLf & _
            "            ASM_FISTP_EAX;" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_ADD_EAX_(8);" & vbCrLf
        src = src & _
            "            ASM_FLD_EAX;" & vbCrLf & _
            "            ASM_MOV_EAX_(pt2);" & vbCrLf & _
            "            ASM_FISTP_EAX;" & vbCrLf & _
            "            t1 = t1 % t2;" & vbCrLf & _
            "            ASM_MOV_EAX_(pt1);" & vbCrLf & _
            "            ASM_FILD_EAX;" & vbCrLf & _
            "            ASM_MOV_EAX_(val_idx);" & vbCrLf & _
            "            ASM_FSTP_EAX;" & vbCrLf & _
            "        } else if (op == '(') {" & vbCrLf & _
            "            val_idx = val_idx + 8;" & vbCrLf & _
            "            if (prec == TOK_RPAREN) {" & vbCrLf & _
            "                op_idx = op_idx - 4;" & vbCrLf & _
            "                break;" & vbCrLf & _
            "            } else if (prec > TOK_RPAREN)" & vbCrLf & _
            "                break;" & vbCrLf & _
            "        }" & vbCrLf & _
            "        op_idx = op_idx - 4;" & vbCrLf & _
            "    }" & vbCrLf & _
            "    *(int *)pval_idx = val_idx;" & vbCrLf & _
            "    *(int *)pop_idx = op_idx;" & vbCrLf & _
            "}" & vbCrLf & _
            "" & vbCrLf
        src = src & _
            "#define PARSE_FLAGS_DEFAULT 0xB14" & vbCrLf & _
            "#define VTBIT_R8 0x20" & vbCrLf & _
            "" & vbCrLf & _
            "parse_num(s, pdbl, psize)" & vbCrLf & _
            "{" & vbCrLf & _
            "    int numparse, dig, variant_res;" & vbCrLf & _
            "" & vbCrLf & _
            "    numparse = alloca(24);" & vbCrLf & _
            "    dig = alloca(30);" & vbCrLf & _
            "    variant_res = alloca(16);" & vbCrLf & _
            "    *(int *)numparse = 30;" & vbCrLf & _
            "    *(int *)(numparse + 4) = PARSE_FLAGS_DEFAULT;" & vbCrLf & _
            "    if (!VarParseNumFromStr(s, 0, 0, numparse, dig)) {" & vbCrLf & _
            "        if (!VarNumFromParseNum(numparse, dig, VTBIT_R8, variant_res)) {" & vbCrLf & _
            "            *(int *)pdbl = *(int *)(variant_res + 8);" & vbCrLf & _
            "            *(int *)(pdbl + 4) = *(int *)(variant_res + 12);" & vbCrLf & _
            "            *(int *)psize = *(int *)(numparse + 12); /* cchUsed */" & vbCrLf & _
            "            return;" & vbCrLf & _
            "        }" & vbCrLf & _
            "    }" & vbCrLf & _
            "    *(int *)pdbl = 0;" & vbCrLf & _
            "    *(int *)(pdbl + 4) = 0;" & vbCrLf & _
            "    *(int *)psize = 1;" & vbCrLf & _
            "}"
        m_ctx.m_state(31) = RtccCompile(m_ctx, src)
    End If
    CallWindowProc m_ctx.m_state(31), StrPtr(sText), VarPtr(SimpleEval)
End Function

