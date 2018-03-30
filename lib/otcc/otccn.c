/*
  Obfuscated Tiny C Compiler

  Copyright (C) 2001-2003 Fabrice Bellard

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product and its documentation
     *is* required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.
*/
#ifndef TINY
#include <stdarg.h>
#endif
#include <stdio.h>

#ifdef _DEBUG
void error(char *fmt,...);
#define assert(expr) { if (!(expr)) error("assert failed: %s", #expr); }
//#define tok_error(tok, tokn, str) error("%s defined 0x%x (should be 0x%x)", tokn, tok, c);
#define tok_assert(tok, str) { int c = (mystrstr(ctx->sym_stk, (str)) - ctx->sym_stk) * 8 + TOK_IDENT; if((tok) != c) error("%s defined 0x%x (should be 0x%x)", #tok, (tok), c);  }
#else
#define assert(expr)
#define tok_assert(tok, str)
#endif

void __cdecl exit(int);
void * __cdecl memcpy(void *d, const void *s, int n);
void * __cdecl memset(void *p, int v, int n);
int __cdecl strcmp(const char *s1, const char *s2);
int __cdecl strlen(const char *s);
#pragma intrinsic(memcpy, memset, strcmp, strlen)


/* vars: value of variables
   loc : local variable index
   glo : global variable index
   ind : output code ptr
   rsym: return symbol
   prog: output code
   dstk: define stack
   dptr, dch: macro state
*/
typedef struct ctx_t {
    // input params
    int prog, sym_stk, mods;
    // in/out params
    int glo, vars;
    // preserved state
    int ind, dstk, t0;
    // local state (discarded after compile)
    int tok, tokc, tokl, tokw, ch, rsym, loc, dptr, dch, last_id;
    int pinp0, pinp;
} *pctx_t;

#define ALLOC_SIZE 99999

/* depends on the init string */
#define TOK_STR_SIZE 83
#define TOK_IDENT    0x100
#define TOK_INT      0x100
#define TOK_IF       0x120
#define TOK_ELSE     0x138
#define TOK_WHILE    0x160
#define TOK_BREAK    0x190
#define TOK_RETURN   0x1c0
#define TOK_FOR      0x1f8
#define TOK_SHORT    0x218
#define TOK_ALLOCA   0x248
#define TOK_ASM      0x280
#define TOK_ASM_EMIT 0x2b0
#define TOK_ASM_MOV  0x2e8
#define TOK_ASM_EAX  0x310
#define TOK_DEFINE   0x330
#define TOK_MAIN     0x368
#define IDX_ASM_EMIT 54
#define IDX_ASM_MOV  61
#define IDX_ASM_EAX  66

#define TOK_DUMMY   1
#define TOK_NUM     2

#define LOCAL   0x200

#define SYM_FORWARD 0
#define SYM_DEFINE  1

/* tokens in string heap */
#define TAG_TOK    ' '
#define TAG_MACRO  2

#define CONST_T0_SIZE 126


compile(pctx_t ctx, int pinp, int wParam, int lParam)
{
    int r;

    if (!ctx->ind) {
        char t0[] = { '+', '+', '#', 'm', '-', '-', '%', 'a', 'm', '*', '@', 'R', '<', '^', '1', 'c', '/', '@', '%', '[', '_', '[', 'H', '3',
            'c', '%', '@', '%', '[', '_', '[', 'H', '3', 'c', '+', '@', '.', 'B', '#', 'd', '-', '@', '%', ':', '_', '^', 'B', 'K', 'd', '<',
            '<', 'Z', '/', '0', '3', 'e', '>', '>', '`', '/', '0', '3', 'e', '<', '=', '0', 'f', '>', '=', '/', 'f', '<', '@', '.', 'f', '>',
            '@', '1', 'f', '=', '=', '&', 'g', '!', '=', '\'', 'g', '&', '&', 'k', '|', '|', '#', 'l', '&', '@', '.', 'B', 'C', 'h', '^',
            '@', '.', 'B', 'S', 'i', '|', '@', '.', 'B', '+', 'j', '~', '@', '/', '%', 'Y', 'd', '!', '@', '&', 'd', '*', '@', 'b', 0 };
        ctx->t0 = ctx->glo;
        memcpy(ctx->t0, t0, CONST_T0_SIZE);
        ctx->glo = ctx->t0 + CONST_T0_SIZE;
        char stk0[] = { ' ', 'i', 'n', 't', ' ', 'i', 'f', ' ', 'e', 'l', 's', 'e', ' ', 'w', 'h', 'i', 'l', 'e', ' ', 'b', 'r', 'e', 'a', 'k', ' ',
            'r', 'e', 't', 'u', 'r', 'n', ' ', 'f', 'o', 'r', ' ', 's', 'h', 'o', 'r', 't', ' ', 'a', 'l', 'l', 'o', 'c', 'a', ' ', 
            '_', 'a', 's', 'm', ' ', '!', '_', 'e', 'm', 'i', 't', ' ', '!', 'm', 'o', 'v', ' ', '!', 'e', 'a', 'x', ' ',
            'd', 'e', 'f', 'i', 'n', 'e', ' ', 'm', 'a', 'i', 'n', ' ' };
        assert(TOK_STR_SIZE == sizeof(stk0));
        memcpy(ctx->sym_stk, stk0, TOK_STR_SIZE);
        tok_assert(TOK_INT, " int ");
        tok_assert(TOK_IF, " if ");
        tok_assert(TOK_ELSE, " else ");
        tok_assert(TOK_WHILE, " while ");
        tok_assert(TOK_BREAK, " break ");
        tok_assert(TOK_RETURN, " return ");
        tok_assert(TOK_FOR, " for ");
        tok_assert(TOK_SHORT, " short ");
        tok_assert(TOK_ASM, " _asm ");
        tok_assert(TOK_ASM_EMIT, "!_emit ");
        tok_assert(TOK_ASM_MOV, "!mov ");
        tok_assert(TOK_ASM_EAX, "!eax ");
        tok_assert(TOK_ALLOCA, " alloca ");
        tok_assert(TOK_DEFINE, " define ");
        tok_assert(TOK_MAIN, " main ");
        assert(*(char *)(ctx->sym_stk + IDX_ASM_EMIT) == '!');
        assert(*(char *)(ctx->sym_stk + IDX_ASM_MOV) == '!');
        assert(*(char *)(ctx->sym_stk + IDX_ASM_EAX) == '!');
        ctx->dstk = ctx->sym_stk + TOK_STR_SIZE;
        ctx->ind = ctx->prog;
    }
    ctx->pinp0 = ctx->pinp = pinp;
    r = ctx->ind;
    inpu(ctx);
    next(ctx);
    decl(ctx, 0);
    return r;
}

pdef(pctx_t ctx, int t)
{
    *(char *)ctx->dstk++ = t;
}

inpu(pctx_t ctx)
{
    if (ctx->dptr) {
        ctx->ch = *(char *)ctx->dptr++;
        if (ctx->ch == TAG_MACRO) {
            ctx->dptr = 0;
            ctx->ch = ctx->dch;
        }
    } else {
        ctx->ch = *(char *)ctx->pinp;
        ctx->ch = ctx->ch + (*(char *)(ctx->pinp + 1) << 8);
        if (ctx->ch != 0)
            ctx->pinp = ctx->pinp + 2;
        else
            ctx->ch = -1;
    }
    /*    printf("ch=%c 0x%x\n", ch, ch); */
    return ctx->ch;
}

peek(pctx_t ctx)
{
    return *(char *)ctx->pinp + (*(char *)(ctx->pinp + 1) << 8);
}

isid(ch)
{
    return myisalnum(ch) | ch == '_';
}

/* read a character constant */
getq(pctx_t ctx)
{
    int c, v;
    if (ctx->ch == '\\') {
        inpu(ctx);
        if (ctx->ch == 'n')
            ctx->ch = '\n';
        else if (ctx->ch == 'r')
            ctx->ch = '\r';
        else if (ctx->ch == 't')
            ctx->ch = '\t';
        else if (ctx->ch == '\\')
            ctx->ch = '\\';
        else if (ctx->ch == 'x') {
            if ((c = inpu(ctx)) > '9')
                c = (c | 0x20) - 39;
            v = (c - '0');
            if ((c = inpu(ctx)) > '9')
                c = (c | 0x20) - 39;
            v = v * 16 + (c - '0');
            if (ctx->tokw && ((c = peek(ctx)) >= '0' && c <= '9' || (c | 0x20) >= 'a' && (c | 0x20) <= 'f'))
            {
                if ((c = inpu(ctx)) > '9')
                    c = (c | 0x20) - 39;
                v = v * 16 + (c - '0');
                if ((c = inpu(ctx)) > '9')
                    c = (c | 0x20) - 39;
                v = v * 16 + (c - '0');
            }
            ctx->ch = v;
        }
    }
}

next(pctx_t ctx)
{
    int t, l, a;

    while (myisspace(ctx->ch) | ctx->ch == '#') {
        if (ctx->ch == '#') {
            inpu(ctx);
            next(ctx);
            if (ctx->tok == TOK_DEFINE) {
                next(ctx);
                pdef(ctx, TAG_TOK); /* fill last ident tag */
                *(int *)ctx->tok = SYM_DEFINE;
                *(int *)(ctx->tok + 4) = ctx->dstk; /* define stack */
            }
            /* well we always save the values ! */
            while (ctx->ch != '\n') {
                pdef(ctx, ctx->ch);
                inpu(ctx);
            }
            pdef(ctx, ctx->ch);
            pdef(ctx, TAG_MACRO);
        }
        inpu(ctx);
    }
    ctx->tokw = (ctx->ch == 'L' && (peek(ctx) == '\"' || peek(ctx) == '\''));
    if (ctx->tokw)
        inpu(ctx);
    ctx->tokl = 0;
    ctx->tok = ctx->ch;
    /* encode identifiers & numbers */
    if (isid(ctx->ch)) {
        pdef(ctx, TAG_TOK);
        ctx->last_id = ctx->dstk;
        while (isid(ctx->ch)) {
            pdef(ctx, ctx->ch);
            inpu(ctx);
        }
        if (myisdigit(ctx->tok)) {
            ctx->tokc = mystrtol(ctx->last_id);
            ctx->tok = TOK_NUM;
        } else {
            *(char *)ctx->dstk = TAG_TOK; /* no need to mark end of string (we
                                        suppose data is initied to zero */
            ctx->tok = mystrstr(ctx->sym_stk, ctx->last_id - 1) - ctx->sym_stk;
            *(char *)ctx->dstk = 0;   /* mark real end of ident for dlsym() */
            ctx->tok = ctx->tok * 8 + TOK_IDENT;
            if (ctx->tok > TOK_DEFINE) {
                ctx->tok = ctx->vars + ctx->tok;
                /*        printf("tok=%s %x\n", last_id, tok); */
                /* define handling */
                if (*(int *)ctx->tok == SYM_DEFINE) {
                    ctx->dptr = *(int *)(ctx->tok + 4);
                    ctx->dch = ctx->ch;
                    inpu(ctx);
                    next(ctx);
                }
            }
        }
    } else {
        inpu(ctx);
        if (ctx->tok == '\'') {
            ctx->tok = TOK_NUM;
            getq(ctx);
            ctx->tokc = ctx->ch;
            inpu(ctx);
            inpu(ctx);
        } else if (ctx->tok == '/' & ctx->ch == '*') {
            inpu(ctx);
            while (ctx->ch) {
                while (ctx->ch != '*')
                    inpu(ctx);
                inpu(ctx);
                if (ctx->ch == '/')
                    ctx->ch = 0;
            }
            inpu(ctx);
            next(ctx);
        } else
        {
            t = ctx->t0;
            while (l = *(char *)t++) {
                a = *(char *)t++;
                ctx->tokc = 0;
                while ((ctx->tokl = *(char *)t++ - 'b') < 0)
                    ctx->tokc = ctx->tokc * 64 + ctx->tokl + 64;
                if (l == ctx->tok & (a == ctx->ch | a == '@')) {
#if 0
                    printf("%c%c -> tokl=%d tokc=0x%x\n",
                           l, a, tokl, tokc);
#endif
                    if (a == ctx->ch) {
                        inpu(ctx);
                        ctx->tok = TOK_DUMMY; /* dummy token for double tokens */
                    }
                    break;
                }
            }
        }
    }
#if 0
    {
        int p;

        printf("tok=0x%x ", tok);
        if (tok >= TOK_IDENT) {
            printf("'");
            if (tok > TOK_DEFINE)
                p = sym_stk + 1 + (tok - vars - TOK_IDENT) / 8;
            else
                p = sym_stk + 1 + (tok - TOK_IDENT) / 8;
            while (*(char *)p != TAG_TOK && *(char *)p)
                printf("%c", *(char *)p++);
            printf("'\n");
        } else if (tok == TOK_NUM) {
            printf("%d\n", tokc);
        } else {
            printf("'%c'\n", tok);
        }
    }
#endif
}

#ifdef TINY
#define skip(ctx, c) next(ctx)
#else

void error(char *fmt,...)
{
    va_list ap;

    va_start(ap, fmt);
    vfprintf(stderr, fmt, ap);
    fprintf(stderr, "\n");
    getchar();
    exit(1);
    va_end(ap);
}

void skip(pctx_t ctx, int c)
{
    if (ctx->tok != c) {
        error("'%c' expected (found '%c') at position %d\n%.40ws", c, ctx->tok, 
            (ctx->pinp - ctx->pinp0 - 4) / 2, ctx->pinp - 4);
    }
    next(ctx);
}

#endif

o(pctx_t ctx, int n)
{
    /* cannot use unsigned, so we must do a hack */
    while (n && n != -1) {
        *(char *)ctx->ind++ = n;
        n = n >> 8;
    }
}

/* output a symbol and patch all calls to it */
gsym(pctx_t ctx, int t)
{
    int n;
    while (t) {
        n = *(int *)t; /* next value */
        *(int *)t = ctx->ind - t - 4;
        t = n;
    }
}

/* psym is used to put an instruction with a data field which is a
   reference to a symbol. It is in fact the same as oad ! */
#define psym oad

/* instruction + address */
oad(pctx_t ctx, int n, int t)
{
    o(ctx, n);
    *(int *)ctx->ind = t;
    t = ctx->ind;
    ctx->ind = ctx->ind + 4;
    return t;
}

/* load immediate value */
li(pctx_t ctx, int t)
{
    oad(ctx, 0xb8, t); /* mov $xx, %eax */
}

gjmp(pctx_t ctx, int t)
{
    return psym(ctx, 0xe9, t);
}

/* l = 0: je, l == 1: jne */
gtst(pctx_t ctx, int l, int t)
{
    o(ctx, 0x0fc085); /* test %eax, %eax, je/jne xxx */
    return psym(ctx, 0x84 + l, t);
}

gcmp(pctx_t ctx, int t)
{
    o(ctx, 0xc139); /* cmp %eax,%ecx */
    li(ctx, 0);
    o(ctx, 0x0f); /* setxx %al */
    o(ctx, t + 0x90);
    o(ctx, 0xc0);
}

gmov(pctx_t ctx, int l, int t)
{
    o(ctx, l + 0x83);
    oad(ctx, (t < LOCAL) << 7 | 5, t);
}

/* l is one if '=' parsing wanted (quick hack) */
unary(register pctx_t ctx, int l)
{
    int n, t, a, c;

    n = 1; /* type of expression 0 = forward, 1 = value, other =
              lvalue */
    if (ctx->tok == '\"') {
        li(ctx, ctx->glo);
        while (ctx->ch != '\"') {
            getq(ctx);
            *(char *)ctx->glo++ = ctx->ch;
            if (ctx->tokw)
                *(char *)ctx->glo++ = ctx->ch >> 8;
            inpu(ctx);
        }
        if (ctx->tokw)
            *(char *)ctx->glo++ = 0;
        *(char *)ctx->glo = 0;
        ctx->glo = ctx->glo + 3 & -4; /* align heap */
        inpu(ctx);
        next(ctx);
    } else {
        c = ctx->tokl;
        a = ctx->tokc;
        t = ctx->tok;
        next(ctx);
        if (t == TOK_NUM) {
            li(ctx, a);
        } else if (t == TOK_ALLOCA) {
            expr(ctx);
            o(ctx, 0x03c083);                                       // add 3, %eax
            o(ctx, 0xfce083);                                       // and -4, %eax
            o(ctx, 0xc429);                                         // sub %eax, %esp
            o(ctx, 0xe089);                                         // mov %esp, %eax
        } else if (c == 2) {
            /* -, +, !, ~ */
            unary(ctx, 0);
            oad(ctx, 0xb9, 0); /* movl $0, %ecx */
            if (t == '!')
                gcmp(ctx, a);
            else
                o(ctx, a);
        } else if (t == '(') {
            expr(ctx);
            skip(ctx, ')');
        } else if (t == '*') {
            /* parse cast */
            skip(ctx, '(');
            t = ctx->tok; /* get type */
            next(ctx); /* skip int/char/void */
            next(ctx); /* skip '*' or '(' */
            if (ctx->tok == '*') {
                /* function type */
                skip(ctx, '*');
                skip(ctx, ')');
                skip(ctx, '(');
                skip(ctx, ')');
                t = 0;
            }
            skip(ctx, ')');
            unary(ctx, 0);
            if (ctx->tok == '=') {
                next(ctx);
                o(ctx, 0x50); /* push %eax */
                expr(ctx);
                o(ctx, 0x59); /* pop %ecx */
                if (t == TOK_SHORT)
                    o(ctx, 0x018966);                           // movw %ax, (%ecx)
                else
                    o(ctx, 0x0188 + (t == TOK_INT)); /* movl %eax/%al, (%ecx) */
            } else if (t) {
                if (t == TOK_INT)
                    o(ctx, 0x8b); /* mov (%eax), %eax */
                else
                    o(ctx, 0xbe0f + ((t == TOK_SHORT)<<8));     // movswl/movsbl (%eax), %eax
                ctx->ind++; /* add zero in code */
            }
        } else if (t == '&') {
            gmov(ctx, 10, *(int *)ctx->tok); /* leal EA, %eax */
            next(ctx);
        } else {
            if (t == TOK_ASM_EAX) {
                n = 0;
            } else {
                n = *(int *)t;
                /* forward reference: try dlsym */
                if (!n)
                    n = dlsym(ctx, 0, ctx->last_id);
            }
            if (ctx->tok == '=' & l) {
                /* assignment */
                next(ctx);
                expr(ctx);
                if (n)
                    gmov(ctx, 6, n); /* mov %eax, EA */
            } else if (ctx->tok != '(') {
                if (n)  /* variable */
                    gmov(ctx, 8, n); /* mov EA, %eax */
                if (ctx->tokl == 11) {                          // `++` -> ctx->tokc=1, `--` -> ctx->tokc=0xFF
                    if (n) {
                        gmov(ctx, 0, n);                        // add $ctx->tokc, EA
                    } else {
                        o(ctx, 0xc083);                         // add $ctx->tokc, %eax
                    }
                    o(ctx, ctx->tokc);
                    next(ctx);
                }
            }
        }
    }

    /* function call */
    if (ctx->tok == '(') {
        if (n == 1)
            o(ctx, 0x50); /* push %eax */

        /* push args and invert order */
        a = oad(ctx, 0xec81, 0); /* sub $xxx, %esp */
        next(ctx);
        l = 0;
        while(ctx->tok != ')') {
            expr(ctx);
            oad(ctx, 0x248489, l); /* movl %eax, xxx(%esp) */
            if (ctx->tok == ',')
                next(ctx);
            l = l + 4;
        }
        *(int *)a = l;
        next(ctx);
        if (!n) {
            /* forward reference */
            t = t + 4;
            *(int *)t = psym(ctx, 0xe8, *(int *)t);
        } else if (n == 1) {
            oad(ctx, 0x2494ff, l); /* call *xxx(%esp) */
            oad(ctx, 0xc481, 4); /* add $xxx, %esp */
        } else {
            oad(ctx, 0xe8, n - ctx->ind - 5); /* call xxx */
        }
    }
}

sum(pctx_t ctx, int l)
{
    int t, n, a;

    if (l-- == 1)
        unary(ctx, 1);
    else {
        sum(ctx, l);
        a = 0;
        while (l == ctx->tokl) {
            n = ctx->tok;
            t = ctx->tokc;
            next(ctx);

            if (l > 8) {
                a = gtst(ctx, t, a); /* && and || output code generation */
                sum(ctx, l);
            } else {
                o(ctx, 0x50); /* push %eax */
                sum(ctx, l);
                o(ctx, 0x59); /* pop %ecx */
                if (l == 4 | l == 5) {
                    gcmp(ctx, t);
                } else {
                    o(ctx, t);
                    if (n == '%')
                        o(ctx, 0x92); /* xchg %edx, %eax */
                }
            }
        }
        /* && and || output code generation */
        if (a && l > 8) {
            a = gtst(ctx, t, a);
            li(ctx, t ^ 1);
            gjmp(ctx, 5); /* jmp $ + 5 */
            gsym(ctx, a);
            li(ctx, t);
        }
    }
}

expr(pctx_t ctx)
{
    sum(ctx, 11);
}


test_expr(pctx_t ctx)
{
    expr(ctx);
    return gtst(ctx, 0, 0);
}

block(pctx_t ctx, int l)
{
    int a, n, t;

    if (ctx->tok == TOK_IF) {
        next(ctx);
        skip(ctx, '(');
        a = test_expr(ctx);
        skip(ctx, ')');
        block(ctx, l);
        if (ctx->tok == TOK_ELSE) {
            next(ctx);
            n = gjmp(ctx, 0); /* jmp */
            gsym(ctx, a);
            block(ctx, l);
            gsym(ctx, n); /* patch else jmp */
        } else {
            gsym(ctx, a); /* patch if test */
        }
    } else if (ctx->tok == TOK_WHILE | ctx->tok == TOK_FOR) {
        t = ctx->tok;
        next(ctx);
        skip(ctx, '(');
        if (t == TOK_WHILE) {
            n = ctx->ind;
            a = test_expr(ctx);
        } else {
            if (ctx->tok != ';')
                expr(ctx);
            skip(ctx, ';');
            n = ctx->ind;
            a = 0;
            if (ctx->tok != ';')
                a = test_expr(ctx);
            skip(ctx, ';');
            if (ctx->tok != ')') {
                t = gjmp(ctx, 0);
                expr(ctx);
                gjmp(ctx, n - ctx->ind - 5);
                gsym(ctx, t);
                n = t + 4;
            }
        }
        skip(ctx, ')');
        block(ctx, &a);
        gjmp(ctx, n - ctx->ind - 5); /* jmp */
        gsym(ctx, a);
    } else if (ctx->tok == TOK_ASM) {
        *(char *)(ctx->sym_stk + IDX_ASM_EMIT) = ' ';
        *(char *)(ctx->sym_stk + IDX_ASM_MOV) = ' ';
        *(char *)(ctx->sym_stk + IDX_ASM_EAX) = ' ';
        while (ctx->tok == TOK_ASM) {
            next(ctx);
            if (ctx->tok == TOK_ASM_EMIT) {
                next(ctx);
                if (ctx->tok == '(')
                    next(ctx);
                if (ctx->tok == TOK_NUM) {
                    if(ctx->tokc == 0)
                        ctx->ind++;
                    else
                        o(ctx, ctx->tokc);
                    next(ctx);
                }
                if (ctx->tok == ')')
                    next(ctx);
            } else if (ctx->tok == TOK_ASM_MOV) {
                next(ctx);
                skip(ctx, TOK_ASM_EAX);
                skip(ctx, ',');
                expr(ctx);
            }
        }
        *(char *)(ctx->sym_stk + IDX_ASM_EMIT) = '!';
        *(char *)(ctx->sym_stk + IDX_ASM_MOV) = '!';
        *(char *)(ctx->sym_stk + IDX_ASM_EAX) = '!';
    } else if (ctx->tok == '{') {
        next(ctx);
        /* declarations */
        decl(ctx, 1);
        while(ctx->tok != '}')
            block(ctx, l);
        next(ctx);
    } else {
        if (ctx->tok == TOK_RETURN) {
            next(ctx);
            if (ctx->tok != ';')
                expr(ctx);
            ctx->rsym = gjmp(ctx, ctx->rsym); /* jmp */
        } else if (ctx->tok == TOK_BREAK) {
            next(ctx);
            *(int *)l = gjmp(ctx, *(int *)l);
        } else if (ctx->tok != ';')
            expr(ctx);
        skip(ctx, ';');
    }
}

/* 'l' is true if local declarations */
decl(pctx_t ctx, int l)
{
    int a, sa;

    while (ctx->tok == TOK_INT | ctx->tok != -1 & !l) {
        if (ctx->tok == TOK_INT) {
            next(ctx);
            while (ctx->tok != ';') {
                if (l) {
                    ctx->loc = ctx->loc + 4;
                    *(int *)ctx->tok = -ctx->loc;
                    next(ctx);
                } else {
                    *(int *)ctx->tok = ctx->glo;
                    next(ctx);
                    if (ctx->tok == '[') {
                        next(ctx);
                        ctx->glo = ctx->glo + ctx->tokc + 3 & -4; /* align heap */
                        next(ctx);
                        skip(ctx, ']');
                    }
                    else {
                        ctx->glo = ctx->glo + 4;
                    }
                }
                if (ctx->tok == ',')
                    next(ctx);
            }
            skip(ctx, ';');
        } else {
            /* patch forward references (XXX: do not work for function
               pointers) */
            gsym(ctx, *(int *)(ctx->tok + 4));
            /* put function address */
            *(int *)ctx->tok = ctx->ind;
            next(ctx);
            skip(ctx, '(');
            a = 8;
            while (ctx->tok != ')') {
                /* read param name and compute offset */
                *(int *)ctx->tok = a;
                a = a + 4;
                next(ctx);
                if (ctx->tok == ',')
                    next(ctx);
            }
            next(ctx); /* skip ')' */
            ctx->rsym = ctx->loc = 0;
            o(ctx, 0xe58955); /* push   %ebp, mov %esp, %ebp */
            sa = oad(ctx, 0xec81, 0); /* sub $xxx, %esp */
            block(ctx, 0);
            gsym(ctx, ctx->rsym);
            o(ctx, 0xc9);                                       // leave
            oad(ctx, 0xc2, a - 8);                              // ret $xx
            ctx->ind = ctx->ind - 2;
            *(int *)sa = ctx->loc; /* save local variables */
        }
    }
}


findsym(module, name)
{
    char *nt_headers;
    char *export_dir;
    unsigned long *names, *funcs;
    unsigned short *nameords;
    int i, n;

    nt_headers = (char *)module + *(int *)((char *)module + 60);
    export_dir = (char *)module + *(int *)(nt_headers + 120);
    if (*(int *)((char *)export_dir + 32) == 0)
        return 0;
    names = (unsigned long *)((char *)module + *(int *)((char *)export_dir + 32)); // export_dir->AddressOfNames
    funcs = (unsigned long *)((char *)module + *(int *)((char *)export_dir + 28)); // export_dir->AddressOfFunctions
    nameords = (unsigned short *)((char *)module + *(int *)((char *)export_dir + 36)); // export_dir->AddressOfNameOrdinals
    n = *(int *)((char *)export_dir + 24); // export_dir->NumberOfNames
    for (i = 0; i < n; i++) {
        char *string = (char *)module + names[i];
        if (strcmp(name, string) == 0) {
            unsigned short nameord = nameords[i];
            unsigned long funcrva = funcs[nameord];
            return (int)((char *)module + funcrva);
        }
    }
    return 0;
}

dlsym(pctx_t ctx, int m, int n)
{
    int p, r;

    if (m)
        return findsym(m, n);
    p = ctx->mods;
    while (*(int *)p) {
        if ((r = findsym(*(int *)p, n)) != 0)
            return r;
        p = p + 4;
    }
    return 0;
}

myisalnum(c) {
    return c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9';
}

myisspace(c) {
    return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

myisdigit(c) {
    return c >= '0' && c <= '9';
}

mystrncmp(s1, s2, len)
{
    int c;

    while(len--) {
        if ((c = *(char *)s1++ - *(char *)s2++) != 0)
            return c;
    }
    return 0;
}

mystrstr(s1, s2)
{
    int len;

    len = strlen(s2);
    while (*(char *)s1) {
        if (*(char *)s1 == *(char *)s2) {
            if (!mystrncmp(s1, s2, len))
                return s1;
        }
        s1++;
    }
    return 0;
}

mystrtol(s)
{
    int v;
    char c;

    while (myisspace(*(char *)s))
        s = s + 1;
    v = 0;
    if (*(char *)s == '0' && *(char *)(s + 1) == 'x') {
        s = s + 2;
        while ((c = *(char *)s++) >= '0' && c <= '9' || (c | 0x20) >= 'a' && (c | 0x20) <= 'f') {
            if (c > '9')
                c = (c | 0x20) - 39;
            v = v * 16 + (c - '0');
        }
    }
    else {
        while ((c = *(char *)s++) >= '0' && c <= '9') {
            v = v * 10 + (c - '0');
        }
    }
    return v;
}

__declspec(naked) void end_of_code(void) {}

#pragma comment(lib, "crypt32")
__declspec(dllimport) int __stdcall CryptBinaryToStringA (int, int, int, int, int);
#define CRYPT_STRING_BASE64 1

output_thunks(code, codesize, name)
{
    int i, j;
    char m_buffer[20000] = { 0 };
    char *m_thunks[100];

    int buflen = sizeof(m_buffer) / sizeof(*m_buffer);
    CryptBinaryToStringA(code, codesize, CRYPT_STRING_BASE64, m_buffer, &buflen);
    char *src = m_buffer, *p = m_buffer;
    for (j = 0; *src; j++) {
        m_thunks[j] = p;
        for(i = 0; i < 512 && *src; i++) {
            while (*src == '\r' || *src == '\n')
                src++;
            *p++ = *src++;
        }
        *p++ = 0;
    }

    printf("' Size=%d\n", codesize);
    for(i = 0; i < j; i++) {
        if (i % 10 == 0)
            printf("Private Const %s%d = _\n", name, i / 10 + 1);
        printf("    \"%s\"%s\n", m_thunks[i], i % 10 < 9 && i < j-1 ? " & _" : "");
    }
}

main(n, t)
{
    output_thunks(compile, (int)end_of_code - (int)compile, "STR_THUNK");

#ifdef _DEBUG 
    return test(n, t);
#endif
}

#ifdef _DEBUG
#include "test.c"
#endif
