#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <psapi.h>
#include <malloc.h>
#include <math.h>

char m_sym_stk[ALLOC_SIZE] = { 0 };
char m_glo[ALLOC_SIZE] = { 0 };
char m_vars[ALLOC_SIZE] = { 0 };
char m_mods[1000] = { 0 };

#pragma comment(lib, "ole32")
int __stdcall CoInitialize(int);
#pragma comment(lib, "oleaut32")
int __stdcall VarParseNumFromStr(int, int, int, int, int);
int __stdcall VarNumFromParseNum(int, int, int, int);

test(n, t) {
    struct ctx_t _ctx = { 0 }, *ctx = &_ctx;
    int pinp, pfn;

    EnumProcessModules(GetCurrentProcess(), m_mods, sizeof(m_mods) / sizeof(int), 0);
    ctx->prog = VirtualAlloc(0, ALLOC_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    ctx->sym_stk = m_sym_stk;
    ctx->glo = m_glo;
    ctx->vars = m_vars;
    ctx->mods = m_mods;
    //pinp = L"main2(n, t) { int v, c; c = L\"\\x12\\x34AAtest\"; v = GlobalAlloc(0x40, 148); *(int *)v = 148; GetVersionExA(v); return *(int *)(v + 4) * 100 + *(int *)(v + 8); }";
    //pfn = compile(&ctx, pinp, 0, 0);
    //pinp = L"main(n, t) { int v; v = GlobalAlloc(0x40, 148); *(int *)v = 148; GetVersionExA(v); return *(int *)(v + 4) * 100 + *(int *)(v + 8); }";
    //pfn = compile(&ctx, pinp, 0, 0);
    //pinp = L"int a, b[100]; main(n, t) { int v, c; v = alloca(550); emit 0x90 0x90; *(int *)v = 55; c = L\"\\x12AF\\x34AAtest\"; eax = *(short *)c; if((eax & 2) != 2) eax = eax + eax; return eax; }";
    pinp = 
L"#define TOK_FINAL     0\n"
L"#define TOK_RPAREN    1\n"
L"#define TOK_ADD       2\n"
L"#define TOK_MOD       3\n"
L"#define TOK_IDIV      4\n"
L"#define TOK_MUL       5\n"
L"#define TOK_UNARY     6\n"
L"#define TOK_POWER     7\n"
L"#define TOK_LPAREN    8\n"
L"#define TOK_NUM       9\n"
L"#define TOK_WHITE     10\n"
L"\n"
L"int lookup[256];\n"
L"\n"
L"simple_eval(s, pdbl)\n"
L"{\n"
L"    int i, p, l, ch, prec, prev_pr;\n"
L"    int op_stack, op_idx;\n"
L"    int val_stack, val_idx;\n"
L"\n"
L"    op_idx = op_stack = alloca(4000);\n"
L"    val_idx = val_stack = alloca(8000);\n"
L"    l = &lookup;\n"
L"    if (*(char *)(l + 32) == 0) {\n"
L"        p = l;\n"
L"        i = 0;\n"
L"        while (i < 256) {\n"
L"            *(char *)p++ = TOK_WHITE;\n"
L"            i++;\n"
L"        }\n"
L"        *(char *)(l + '(') = TOK_LPAREN;\n"
L"        *(char *)(l + ')') = TOK_RPAREN;\n"
L"        *(char *)(l + '+') = TOK_ADD;\n"
L"        *(char *)(l + '-') = TOK_ADD;\n"
L"        *(char *)(l + '*') = TOK_MUL;\n"
L"        *(char *)(l + '/') = TOK_MUL;\n"
L"        *(char *)(l + '^') = TOK_POWER;\n"
L"        *(char *)(l + '\\\\') = TOK_IDIV;\n"
L"        *(char *)(l + '%') = TOK_MOD;\n"
L"        *(char *)(l + '.') = TOK_NUM;\n"
L"        p = l + '0';\n"
L"        i = '0';\n"
L"        while (i <= '9') {\n"
L"            *(char *)p++ = TOK_NUM;\n"
L"            i++;\n"
L"        }\n"
L"    }\n"
L"    prev_pr = 0;\n"
L"    p = s;\n"
L"    while ((ch = *(short *)p)) {\n"
L"        if (!(ch >> 8)) {\n"
L"            prec = *(char *)(l + ch);\n"
L"            if (prec != TOK_WHITE) {\n"
L"                if (prec == TOK_NUM) {\n"
L"                    val_idx = val_idx + 8;\n"
L"                    p = fast_val(p, val_idx);\n"
L"                } else if (prec == TOK_ADD) {\n"
L"                    if (prev_pr >= TOK_ADD && prev_pr < TOK_NUM)\n"
L"                        prec = TOK_UNARY;\n"
L"                }\n"
L"                if (prec >= TOK_ADD && prec < TOK_NUM) {\n"
L"                    if(prec != TOK_UNARY)\n"
L"                        eval_stack(prec, op_stack, &op_idx, val_stack, &val_idx);\n"
L"                    op_idx = op_idx + 4;\n"
L"                    *(int *)op_idx = (prec << 16) + ch;\n"
L"                }\n"
L"                prev_pr = prec;\n"
L"            }\n"
L"        }\n"
L"        p++; p++;\n"
L"    }\n"
L"    eval_stack(TOK_FINAL, op_stack, &op_idx, val_stack, &val_idx);\n"
L"    *(int *)pdbl = *(int *)val_idx;\n"
L"    *(int *)(pdbl + 4) = *(int *)(val_idx + 4);\n"
L"}\n"
L"\n"
L"#define ASM_MOV_EAX_    _asm mov eax,\n"
L"#define ASM_ADD_EAX_    _asm _emit 0x83 _asm _emit 0xc0 _asm _emit\n"
L"#define ASM_SUB_EAX_    _asm _emit 0x83 _asm _emit 0xe8 _asm _emit\n"
L"#define ASM_FSTP_EAX    _asm _emit 0xdd _asm _emit 0x18\n"
L"#define ASM_FLD_EAX     _asm _emit 0xdd _asm _emit 0x00\n"
L"#define ASM_FLD_EAX_    _asm _emit 0xdd _asm _emit 0x40 _asm _emit\n"
L"#define ASM_FADD_EAX    _asm _emit 0xdc _asm _emit 0x00\n"
L"#define ASM_FADD_EAX_   _asm _emit 0xdc _asm _emit 0x40 _asm _emit\n"
L"#define ASM_FADDP_ST1   _asm _emit 0xde _asm _emit 0xc1\n"
L"#define ASM_FSUB_EAX_   _asm _emit 0xdc _asm _emit 0x60 _asm _emit\n"
L"#define ASM_FMUL_EAX    _asm _emit 0xdc _asm _emit 0x08\n"
L"#define ASM_FMUL_EAX_   _asm _emit 0xdc _asm _emit 0x48 _asm _emit\n"
L"#define ASM_FMULP_ST1   _asm _emit 0xde _asm _emit 0xc9\n"
L"#define ASM_FDIV_EAX    _asm _emit 0xdc _asm _emit 0x30\n"
L"#define ASM_FDIV_EAX_   _asm _emit 0xdc _asm _emit 0x70 _asm _emit\n"
L"#define ASM_FCHS        _asm _emit 0xd9 _asm _emit 0xe0\n"
L"#define ASM_FILD_EAX    _asm _emit 0xdb _asm _emit 0x00\n"
L"#define ASM_FISTP_EAX   _asm _emit 0xdb _asm _emit 0x18\n"
L"#define ASM_FYL2X       _asm _emit 0xd9 _asm _emit 0xf1\n"
L"#define ASM_FLD1        _asm _emit 0xd9 _asm _emit 0xe8\n"
L"#define ASM_FLD_ST1     _asm _emit 0xd9 _asm _emit 0xc1\n"
L"#define ASM_FPREM       _asm _emit 0xd9 _asm _emit 0xf8\n"
L"#define ASM_F2XM1       _asm _emit 0xd9 _asm _emit 0xf0\n"
L"#define ASM_FSCALE      _asm _emit 0xd9 _asm _emit 0xfd\n"
L"#define ASM_FLDZ        _asm _emit 0xd9 _asm _emit 0xee\n"
L"\n"
L"eval_stack(prec, op_stack, pop_idx, val_stack, pval_idx)\n"
L"{\n"
L"    int op_idx, val_idx, op, t1, pt1, t2, pt2;\n"
L"\n"
L"    op_idx = *(int *)pop_idx;\n"
L"    val_idx = *(int *)pval_idx;\n"
L"    while (op_idx > op_stack) {\n"
L"        if (*(int *)(op_idx) < (prec << 16))\n"
L"            break;\n"
L"        val_idx = val_idx - 8;\n"
L"        op = *(short *)op_idx;\n"
L"        if (op == '+') {\n"
L"            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {\n"
L"                val_idx = val_idx + 8;\n"
L"            } else {\n"
L"                /* *(double *)val_idx = *(double *)val_idx + *(double *)(val_idx + 8); */\n"
L"                ASM_MOV_EAX_(val_idx);\n"
L"                ASM_FLD_EAX;\n"
L"                ASM_FADD_EAX_(8);\n"
L"                ASM_FSTP_EAX;\n"
L"            }\n"
L"        } else if (op == '-') {\n"
L"            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {\n"
L"                val_idx = val_idx + 8;\n"
L"                /* *(double *)val_idx = -*(double *)val_idx; */\n"
L"                ASM_MOV_EAX_(val_idx);\n"
L"                ASM_FLD_EAX;\n"
L"                ASM_FCHS;\n"
L"                ASM_FSTP_EAX;\n"
L"            } else {\n"
L"                /* *(double *)val_idx = *(double *)val_idx - *(double *)(val_idx + 8); */\n"
L"                ASM_MOV_EAX_(val_idx);\n"
L"                ASM_FLD_EAX;\n"
L"                ASM_FSUB_EAX_(8);\n"
L"                ASM_FSTP_EAX;\n"
L"            }\n"
L"        } else if (op == '*') {\n"
L"            /* *(double *)val_idx = *(double *)val_idx * *(double *)(val_idx + 8); */\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_FMUL_EAX_(8);\n"
L"            ASM_FSTP_EAX;\n"
L"        } else if (op == '/') {\n"
L"            /* *(double *)val_idx = *(double *)val_idx / *(double *)(val_idx + 8); */\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_FDIV_EAX_(8);\n"
L"            ASM_FSTP_EAX;\n"
L"        } else if (op == '^') {\n"
L"            /* *(double *)val_idx = pow(*(double *)val_idx, *(double *)(val_idx + 8)); */\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_ADD_EAX_(8);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_SUB_EAX_(8);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_FYL2X;\n"
L"            ASM_FLD1;\n"
L"            ASM_FLD_ST1;\n"
L"            ASM_FPREM;\n"
L"            ASM_F2XM1;\n"
L"            ASM_FADDP_ST1;\n"
L"            ASM_FSCALE;\n"
L"            ASM_FSTP_EAX;\n"
L"        } else if (op == '\\\\') {\n"
L"            pt1 = &t1;\n"
L"            /* *(double *)val_idx = (int)(*(double *)val_idx / *(double *)(val_idx + 8)); */\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_FDIV_EAX_(8);\n"
L"            ASM_MOV_EAX_(pt1);\n"
L"            ASM_FISTP_EAX;\n"
L"            ASM_FILD_EAX;\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FSTP_EAX;\n"
L"        } else if (op == '%') {\n"
L"            pt1 = &t1;\n"
L"            pt2 = &t2;\n"
L"            /* *(double *)val_idx = (int)*(double *)val_idx % (int)*(double *)(val_idx + 8); */\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_MOV_EAX_(pt1);\n"
L"            ASM_FISTP_EAX;\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_ADD_EAX_(8);\n"
L"            ASM_FLD_EAX;\n"
L"            ASM_MOV_EAX_(pt2);\n"
L"            ASM_FISTP_EAX;\n"
L"            t1 = t1 % t2;\n"
L"            ASM_MOV_EAX_(pt1);\n"
L"            ASM_FILD_EAX;\n"
L"            ASM_MOV_EAX_(val_idx);\n"
L"            ASM_FSTP_EAX;\n"
L"        } else if (op == '(') {\n"
L"            val_idx = val_idx + 8;\n"
L"            if (prec == TOK_RPAREN) {\n"
L"                op_idx = op_idx - 4;\n"
L"                break;\n"
L"            } else if (prec > TOK_RPAREN)\n"
L"                break;\n"
L"        }\n"
L"        op_idx = op_idx - 4;\n"
L"    }\n"
L"    *(int *)pval_idx = val_idx;\n"
L"    *(int *)pop_idx = op_idx;\n"
L"}\n"
L"\n"
L"fast_val(p, pdbl)\n"
L"{\n"
L"    int ch, addr;\n"
L"    int newval, esgn, eint, hasfrac;\n"
L"    int intpart, fracpart, fracdiv, dbl10; /* doubles */\n"
L"\n"
L"    intpart = alloca(8);\n"
L"    fracpart = alloca(8);\n"
L"    fracdiv = alloca(8);\n"
L"    dbl10 = alloca(8);\n"
L"    newval = esgn = hasfrac = 0;\n"
L"    /* *(double *)intpart = *(double *)fracpart = 0 */\n"
L"    ASM_FLDZ;\n"
L"    ASM_MOV_EAX_(intpart);\n"
L"    ASM_FSTP_EAX;\n"
L"    ASM_FLDZ;\n"
L"    ASM_MOV_EAX_(fracpart);\n"
L"    ASM_FSTP_EAX;\n"
L"    /* *(double *)fracdiv = 1 */\n"
L"    ASM_FLD1;\n"
L"    ASM_MOV_EAX_(fracdiv);\n"
L"    ASM_FSTP_EAX;\n"
L"    /* *(double *)dbl10 = 10 */\n"
L"    ch = 10;\n"
L"    addr = &ch;\n"
L"    ASM_MOV_EAX_(addr);\n"
L"    ASM_FILD_EAX;\n"
L"    ASM_MOV_EAX_(dbl10);\n"
L"    ASM_FSTP_EAX;\n"
L"    while ((ch = *(short *)p)) {\n"
L"        if (ch >= '0' && ch <= '9') {\n"
L"            newval = 1;\n"
L"            if (esgn) {\n"
L"                eint = eint * 10 + ch - '0';\n"
L"            } else {\n"
L"                ch = ch - '0';\n"
L"                if(!hasfrac) {\n"
L"                    /* *(double *)intpart = *(double *)intpart * *(double*)dbl10 + ch; */\n"
L"                    ASM_MOV_EAX_(intpart);\n"
L"                    ASM_FLD_EAX;\n"
L"                    ASM_MOV_EAX_(dbl10);\n"
L"                    ASM_FMUL_EAX;\n"
L"                    addr = &ch;\n"
L"                    ASM_MOV_EAX_(addr);\n"
L"                    ASM_FILD_EAX;\n"
L"                    ASM_FADDP_ST1;\n"
L"                    ASM_MOV_EAX_(intpart);\n"
L"                    ASM_FSTP_EAX\n"
L"                } else {\n"
L"                    /* *(double *)fracpart = *(double *)fracpart * *(double*)dbl10 + ch; */\n"
L"                    ASM_MOV_EAX_(fracpart);\n"
L"                    ASM_FLD_EAX;\n"
L"                    ASM_MOV_EAX_(dbl10);\n"
L"                    ASM_FMUL_EAX;\n"
L"                    addr = &ch;\n"
L"                    ASM_MOV_EAX_(addr);\n"
L"                    ASM_FILD_EAX;\n"
L"                    ASM_FADDP_ST1;\n"
L"                    ASM_MOV_EAX_(fracpart);\n"
L"                    ASM_FSTP_EAX\n"
L"                    /* *(double *)fracdiv = *(double *)fracdiv * *(double*)dbl10; */\n"
L"                    ASM_MOV_EAX_(fracdiv);\n"
L"                    ASM_FLD_EAX;\n"
L"                    ASM_MOV_EAX_(dbl10);\n"
L"                    ASM_FMUL_EAX;\n"
L"                    ASM_MOV_EAX_(fracdiv);\n"
L"                    ASM_FSTP_EAX\n"
L"                }\n"
L"            }\n"
L"        } else if (ch == '.') {\n"
L"            if (hasfrac)\n"
L"                break;\n"
L"            newval = 1;\n"
L"            hasfrac = 1;\n"
L"        } else if (ch == 'e' || ch == 'E') {\n"
L"            if (esgn)\n"
L"                break;\n"
L"            esgn = newval;\n"
L"            eint = 0;\n"
L"        } else if (ch == '-') {\n"
L"            if (esgn > 0)\n"
L"                esgn = -1;\n"
L"            else\n"
L"                break;\n"
L"        } else\n"
L"            break;\n"
L"        p++; p++;\n"
L"    }\n"
L"    /* *(double *)pdbl = newval * *(double *)intpart; */\n"
L"    addr = &newval;\n"
L"    ASM_MOV_EAX_(addr);\n"
L"    ASM_FILD_EAX;\n"
L"    ASM_MOV_EAX_(intpart);\n"
L"    ASM_FMUL_EAX;\n"
L"    ASM_MOV_EAX_(pdbl);\n"
L"    ASM_FSTP_EAX\n"
L"    if (hasfrac) {\n"
L"        /* *(double *)pdbl = *(double *)pdbl + newval * *(double *)fracpart / *(double *)fracdiv; */\n"
L"        addr = &newval;\n"
L"        ASM_MOV_EAX_(addr);\n"
L"        ASM_FILD_EAX;\n"
L"        ASM_MOV_EAX_(fracpart);\n"
L"        ASM_FMUL_EAX;\n"
L"        ASM_MOV_EAX_(fracdiv);\n"
L"        ASM_FDIV_EAX;\n"
L"        ASM_MOV_EAX_(pdbl);\n"
L"        ASM_FADD_EAX;\n"
L"        ASM_FSTP_EAX\n"
L"    }\n"
L"    if (esgn) {\n"
L"        /* *(double *)pdbl = *(double *)pdbl * pow(dbl10, esgn * eint); */\n"
L"        addr = &esgn;\n"
L"        ASM_MOV_EAX_(addr);\n"
L"        ASM_FILD_EAX;\n"
L"        addr = &eint;\n"
L"        ASM_MOV_EAX_(addr);\n"
L"        ASM_FILD_EAX;\n"
L"        ASM_FMULP_ST1;\n"
L"        ASM_MOV_EAX_(dbl10);\n"
L"        ASM_FLD_EAX;\n"
L"        ASM_FYL2X;\n"
L"        ASM_FLD1;\n"
L"        ASM_FLD_ST1;\n"
L"        ASM_FPREM;\n"
L"        ASM_F2XM1;\n"
L"        ASM_FADDP_ST1;\n"
L"        ASM_FSCALE;\n"
L"        ASM_MOV_EAX_(pdbl);\n"
L"        ASM_FMUL_EAX;\n"
L"        ASM_FSTP_EAX;\n"
L"    }\n"
L"    return p;\n"
L"}";
    pfn = compile(ctx, pinp, 0, 0);

    double dbl1, dbl2;
    wchar_t *expr;

    CoInitialize(0);
    expr = L"(3.5) + 2.9 * (2 + -(1 + 2))";
    simple_eval(expr, &dbl1);
    (*(int (__stdcall *)())pfn)(expr, &dbl2);
    assert(dbl1 == dbl2);
}


#define TOK_FINAL     0
#define TOK_RPAREN    1
#define TOK_ADD       2
#define TOK_MOD       3
#define TOK_IDIV      4
#define TOK_MUL       5
#define TOK_UNARY     6
#define TOK_POWER     7
#define TOK_LPAREN    8
#define TOK_NUM       9
#define TOK_WHITE     10

int lookup[256];

simple_eval(s, pdbl)
{
    int i, p, l, ch, prec, prev_pr;
    int op_stack, op_idx;
    int val_stack, val_idx;

    op_idx = op_stack = alloca(4000);
    val_idx = val_stack = alloca(8000);
    l = &lookup;
    if (*(char *)(l + 32) == 0) {
        p = l;
        i = 0;
        while (i < 256) {
            *(char *)p++ = TOK_WHITE;
            i++;
        }
        *(char *)(l + '(') = TOK_LPAREN;
        *(char *)(l + ')') = TOK_RPAREN;
        *(char *)(l + '+') = TOK_ADD;
        *(char *)(l + '-') = TOK_ADD;
        *(char *)(l + '*') = TOK_MUL;
        *(char *)(l + '/') = TOK_MUL;
        *(char *)(l + '^') = TOK_POWER;
        *(char *)(l + '\\') = TOK_IDIV;
        *(char *)(l + '%') = TOK_MOD;
        *(char *)(l + '.') = TOK_NUM;
        p = l + '0';
        i = '0';
        while (i <= '9') {
            *(char *)p++ = TOK_NUM;
            i++;
        }
    }
    prev_pr = 0;
    p = s;
    while ((ch = *(short *)p)) {
        if (!(ch >> 8)) {
            prec = *(char *)(l + ch);
            if (prec != TOK_WHITE) {
                if (prec == TOK_NUM) {
                    val_idx = val_idx + 8;
                    p = fast_val(p, val_idx);
                } else if (prec == TOK_ADD) {
                    if (prev_pr >= TOK_ADD && prev_pr < TOK_NUM)
                        prec = TOK_UNARY;
                }
                if (prec >= TOK_ADD && prec < TOK_NUM) {
                    if(prec != TOK_UNARY)
                        eval_stack(prec, op_stack, &op_idx, val_stack, &val_idx);
                    op_idx = op_idx + 4;
                    *(int *)op_idx = (prec << 16) + ch;
                }
                prev_pr = prec;
            }
        }
        p++; p++;
    }
    eval_stack(TOK_FINAL, op_stack, &op_idx, val_stack, &val_idx);
    *(int *)pdbl = *(int *)val_idx;
    *(int *)(pdbl + 4) = *(int *)(val_idx + 4);
}

#define ASM_MOV_EAX_    _asm mov eax,
#define ASM_ADD_EAX_    _asm _emit 0x83 _asm _emit 0xc0 _asm _emit
#define ASM_SUB_EAX_    _asm _emit 0x83 _asm _emit 0xe8 _asm _emit
#define ASM_FSTP_EAX    _asm _emit 0xdd _asm _emit 0x18
#define ASM_FLD_EAX     _asm _emit 0xdd _asm _emit 0x00
#define ASM_FLD_EAX_    _asm _emit 0xdd _asm _emit 0x40 _asm _emit
#define ASM_FADD_EAX    _asm _emit 0xdc _asm _emit 0x00
#define ASM_FADD_EAX_   _asm _emit 0xdc _asm _emit 0x40 _asm _emit
#define ASM_FADDP_ST1   _asm _emit 0xde _asm _emit 0xc1
#define ASM_FSUB_EAX_   _asm _emit 0xdc _asm _emit 0x60 _asm _emit
#define ASM_FMUL_EAX    _asm _emit 0xdc _asm _emit 0x08
#define ASM_FMUL_EAX_   _asm _emit 0xdc _asm _emit 0x48 _asm _emit
#define ASM_FMULP_ST1   _asm _emit 0xde _asm _emit 0xc9
#define ASM_FDIV_EAX    _asm _emit 0xdc _asm _emit 0x30
#define ASM_FDIV_EAX_   _asm _emit 0xdc _asm _emit 0x70 _asm _emit
#define ASM_FCHS        _asm _emit 0xd9 _asm _emit 0xe0
#define ASM_FILD_EAX    _asm _emit 0xdb _asm _emit 0x00
#define ASM_FISTP_EAX   _asm _emit 0xdb _asm _emit 0x18
#define ASM_FYL2X       _asm _emit 0xd9 _asm _emit 0xf1
#define ASM_FLD1        _asm _emit 0xd9 _asm _emit 0xe8
#define ASM_FLD_ST1     _asm _emit 0xd9 _asm _emit 0xc1
#define ASM_FPREM       _asm _emit 0xd9 _asm _emit 0xf8
#define ASM_F2XM1       _asm _emit 0xd9 _asm _emit 0xf0
#define ASM_FSCALE      _asm _emit 0xd9 _asm _emit 0xfd
#define ASM_FLDZ        _asm _emit 0xd9 _asm _emit 0xee

eval_stack(prec, op_stack, pop_idx, val_stack, pval_idx)
{
    int op_idx, val_idx, op, t1, pt1, t2, pt2;

    op_idx = *(int *)pop_idx;
    val_idx = *(int *)pval_idx;
    while (op_idx > op_stack) {
        if (*(int *)(op_idx) < (prec << 16))
            break;
        val_idx = val_idx - 8;
        op = *(short *)op_idx;
        if (op == '+') {
            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {
                val_idx = val_idx + 8;
            } else {
                /* *(double *)val_idx = *(double *)val_idx + *(double *)(val_idx + 8); */
                ASM_MOV_EAX_(val_idx);
                ASM_FLD_EAX;
                ASM_FADD_EAX_(8);
                ASM_FSTP_EAX;
            }
        } else if (op == '-') {
            if (*(int *)(op_idx) > (TOK_UNARY << 16)) {
                val_idx = val_idx + 8;
                /* *(double *)val_idx = -*(double *)val_idx; */
                ASM_MOV_EAX_(val_idx);
                ASM_FLD_EAX;
                ASM_FCHS;
                ASM_FSTP_EAX;
            } else {
                /* *(double *)val_idx = *(double *)val_idx - *(double *)(val_idx + 8); */
                ASM_MOV_EAX_(val_idx);
                ASM_FLD_EAX;
                ASM_FSUB_EAX_(8);
                ASM_FSTP_EAX;
            }
        } else if (op == '*') {
            /* *(double *)val_idx = *(double *)val_idx * *(double *)(val_idx + 8); */
            ASM_MOV_EAX_(val_idx);
            ASM_FLD_EAX;
            ASM_FMUL_EAX_(8);
            ASM_FSTP_EAX;
        } else if (op == '/') {
            /* *(double *)val_idx = *(double *)val_idx / *(double *)(val_idx + 8); */
            ASM_MOV_EAX_(val_idx);
            ASM_FLD_EAX;
            ASM_FDIV_EAX_(8);
            ASM_FSTP_EAX;
        } else if (op == '^') {
            /* *(double *)val_idx = pow(*(double *)val_idx, *(double *)(val_idx + 8)); */
            ASM_MOV_EAX_(val_idx);
            ASM_ADD_EAX_(8);
            ASM_FLD_EAX;
            ASM_SUB_EAX_(8);
            ASM_FLD_EAX;
            ASM_FYL2X;
            ASM_FLD1;
            ASM_FLD_ST1;
            ASM_FPREM;
            ASM_F2XM1;
            ASM_FADDP_ST1;
            ASM_FSCALE;
            ASM_FSTP_EAX;
        } else if (op == '\\') {
            pt1 = &t1;
            /* *(double *)val_idx = (int)(*(double *)val_idx / *(double *)(val_idx + 8)); */
            ASM_MOV_EAX_(val_idx);
            ASM_FLD_EAX;
            ASM_FDIV_EAX_(8);
            ASM_MOV_EAX_(pt1);
            ASM_FISTP_EAX;
            ASM_FILD_EAX;
            ASM_MOV_EAX_(val_idx);
            ASM_FSTP_EAX;
        } else if (op == '%') {
            pt1 = &t1;
            pt2 = &t2;
            /* *(double *)val_idx = (int)*(double *)val_idx % (int)*(double *)(val_idx + 8); */
            ASM_MOV_EAX_(val_idx);
            ASM_FLD_EAX;
            ASM_MOV_EAX_(pt1);
            ASM_FISTP_EAX;
            ASM_MOV_EAX_(val_idx);
            ASM_ADD_EAX_(8);
            ASM_FLD_EAX;
            ASM_MOV_EAX_(pt2);
            ASM_FISTP_EAX;
            t1 = t1 % t2;
            ASM_MOV_EAX_(pt1);
            ASM_FILD_EAX;
            ASM_MOV_EAX_(val_idx);
            ASM_FSTP_EAX;
        } else if (op == '(') {
            val_idx = val_idx + 8;
            if (prec == TOK_RPAREN) {
                op_idx = op_idx - 4;
                break;
            } else if (prec > TOK_RPAREN)
                break;
        }
        op_idx = op_idx - 4;
    }
    *(int *)pval_idx = val_idx;
    *(int *)pop_idx = op_idx;
}

fast_val(p, pdbl)
{
    int ch, addr;
    int newval, esgn, eint, hasfrac;
    int intpart, fracpart, fracdiv, dbl10; /* doubles */

    intpart = alloca(8);
    fracpart = alloca(8);
    fracdiv = alloca(8);
    dbl10 = alloca(8);
    newval = esgn = hasfrac = 0;
    /* *(double *)intpart = *(double *)fracpart = 0 */
    ASM_FLDZ;
    ASM_MOV_EAX_(intpart);
    ASM_FSTP_EAX;
    ASM_FLDZ;
    ASM_MOV_EAX_(fracpart);
    ASM_FSTP_EAX;
    /* *(double *)fracdiv = 1 */
    ASM_FLD1;
    ASM_MOV_EAX_(fracdiv);
    ASM_FSTP_EAX;
    /* *(double *)dbl10 = 10 */
    ch = 10;
    addr = &ch;
    ASM_MOV_EAX_(addr);
    ASM_FILD_EAX;
    ASM_MOV_EAX_(dbl10);
    ASM_FSTP_EAX;
    while ((ch = *(short *)p)) {
        if (ch >= '0' && ch <= '9') {
            newval = 1;
            if (esgn) {
                eint = eint * 10 + ch - '0';
            } else {
                ch = ch - '0';
                if(!hasfrac) {
                    /* *(double *)intpart = *(double *)intpart * *(double*)dbl10 + ch; */
                    ASM_MOV_EAX_(intpart);
                    ASM_FLD_EAX;
                    ASM_MOV_EAX_(dbl10);
                    ASM_FMUL_EAX;
                    addr = &ch;
                    ASM_MOV_EAX_(addr);
                    ASM_FILD_EAX;
                    ASM_FADDP_ST1;
                    ASM_MOV_EAX_(intpart);
                    ASM_FSTP_EAX
                } else {
                    /* *(double *)fracpart = *(double *)fracpart * *(double*)dbl10 + ch; */
                    ASM_MOV_EAX_(fracpart);
                    ASM_FLD_EAX;
                    ASM_MOV_EAX_(dbl10);
                    ASM_FMUL_EAX;
                    addr = &ch;
                    ASM_MOV_EAX_(addr);
                    ASM_FILD_EAX;
                    ASM_FADDP_ST1;
                    ASM_MOV_EAX_(fracpart);
                    ASM_FSTP_EAX
                    /* *(double *)fracdiv = *(double *)fracdiv * *(double*)dbl10; */
                    ASM_MOV_EAX_(fracdiv);
                    ASM_FLD_EAX;
                    ASM_MOV_EAX_(dbl10);
                    ASM_FMUL_EAX;
                    ASM_MOV_EAX_(fracdiv);
                    ASM_FSTP_EAX
                }
            }
        } else if (ch == '.') {
            if (hasfrac)
                break;
            newval = 1;
            hasfrac = 1;
        } else if (ch == 'e' || ch == 'E') {
            if (esgn)
                break;
            esgn = newval;
            eint = 0;
        } else if (ch == '-') {
            if (esgn > 0)
                esgn = -1;
            else
                break;
        } else
            break;
        p++; p++;
    }
    /* *(double *)pdbl = newval * *(double *)intpart; */
    addr = &newval;
    ASM_MOV_EAX_(addr);
    ASM_FILD_EAX;
    ASM_MOV_EAX_(intpart);
    ASM_FMUL_EAX;
    ASM_MOV_EAX_(pdbl);
    ASM_FSTP_EAX
    if (hasfrac) {
        /* *(double *)pdbl = *(double *)pdbl + newval * *(double *)fracpart / *(double *)fracdiv; */
        addr = &newval;
        ASM_MOV_EAX_(addr);
        ASM_FILD_EAX;
        ASM_MOV_EAX_(fracpart);
        ASM_FMUL_EAX;
        ASM_MOV_EAX_(fracdiv);
        ASM_FDIV_EAX;
        ASM_MOV_EAX_(pdbl);
        ASM_FADD_EAX;
        ASM_FSTP_EAX
    }
    if (esgn) {
        /* *(double *)pdbl = *(double *)pdbl * pow(dbl10, esgn * eint); */
        addr = &esgn;
        ASM_MOV_EAX_(addr);
        ASM_FILD_EAX;
        addr = &eint;
        ASM_MOV_EAX_(addr);
        ASM_FILD_EAX;
        ASM_FMULP_ST1;
        ASM_MOV_EAX_(dbl10);
        ASM_FLD_EAX;
        ASM_FYL2X;
        ASM_FLD1;
        ASM_FLD_ST1;
        ASM_FPREM;
        ASM_F2XM1;
        ASM_FADDP_ST1;
        ASM_FSCALE;
        ASM_MOV_EAX_(pdbl);
        ASM_FMUL_EAX;
        ASM_FSTP_EAX;
    }
    return p;
}

#if 0
#define PARSE_FLAGS_DEFAULT 0xB14
#define VTBIT_R8 0x20

parse_num(s, pdbl, psize)
{
    int numparse, dig, variant_res;

    numparse = alloca(24);
    dig = alloca(30);
    variant_res = alloca(16);
    *(int *)numparse = 30;
    *(int *)(numparse + 4) = PARSE_FLAGS_DEFAULT;
    if (!VarParseNumFromStr(s, 0, 0, numparse, dig)) {
        if (!VarNumFromParseNum(numparse, dig, VTBIT_R8, variant_res)) {
            *(int *)pdbl = *(int *)(variant_res + 8);
            *(int *)(pdbl + 4) = *(int *)(variant_res + 12);
            *(int *)psize = *(int *)(numparse + 12); /* cchUsed */
            return;
        }
    }
    *(int *)pdbl = 0;
    *(int *)(pdbl + 4) = 0;
    *(int *)psize = 1;
}
#endif
