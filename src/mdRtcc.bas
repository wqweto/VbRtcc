Attribute VB_Name = "mdRtcc"
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

'=========================================================================
' Thunk data
'=========================================================================

' Size=7040
Private Const STR_THUNK1 = _
    "VYvsgeyAAAAAU4tdCFaDexQAD4WwAQAAV4t7DI11gIl7HLkfAAAAx0X4JmQqQMdFqC1AJTrHRaxfXkJLx0WwZDw8WsdFtC8wM2XHRbg+PmAvx0W8MDNlPMdFwD0wZj7HRcQ9L2Y8x0XIQC5mPsdFzEAxZj3HRdA9Jmchx0XUPSdnJsdF2CZrfHzHRdwjbCZAx0XgLkJDaMdF5F5ALkLHRehTaXxAx0XsLkIrasdF8H5ALyXHRfRZZCFAx0WAKysjbcdFhC0tJWHHRYhtKkBSx0WMPF4xY8dFkC9AJVvHRZRfW0gzx0WYYyVAJcdFnFtfW0jHRaAzYytAx0WkLkIjZGbHRfxiAPOluRQAAABmpYtDHI11qIt7BIPAfolDDMdFqCBpbnTHRawgaWYgx0WwZWxzZcdFtCB3aGnHRbhsZSBix0W8cmVha8dFwCByZXTHRcR1cm4gx0XIZm9yIMdFzHNob3LHRdB0IGFsx0XUbG9jYcdF2CBfYXPHRdxtICFf" & _
    "x0XgZW1pdMdF5CAhbW/HReh2ICFlx0XsYXggZMdF8GVmaW7HRfRlIG1hZsdF+GluxkX6IPOlZqWki0MEg8BTiUMYiwOJQxRfi1UMi0M8i3MUiVNMiVNIhcB0Hg++CECJSzCJQzyD+QJ1MotDQMdDPAAAAACJQzDrIw++ColLMA++QgHB4AgDyIlLMHQIjUICiUNM6wfHQzD/////i8vorQEAADPSi8voZBUAAIvGXluL5V3DzMzMzMzMzMzMzMzMi9GLSjyFyXQgD74BiUIwjUEBg3owAolCPHU5i0JAx0I8AAAAAIlCMMNWi3JMD74OiUowD75GAcHgCAPBiUIwdAuNRgKJQkyLQjBew8dCMP////9ei0Iww8zMzMzMzMzMVovxg34wXA+FHAEAAItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4tGMIP4" & _
    "bnUJx0YwCgAAAF7Dg/hydQnHRjANAAAAXsOD+HR1CcdGMAkAAABew4P4XHUFiUYwXsOD+HgPhY4AAABXi87o/f7//4v4g/85fgaDzyCD7yeLzujp/v//g/g5fgaDyCCD6CeDx83B5wQD+IN+LAB0VYtGTA++SAEPvgDB4QgDyIP5MHwFg/k5fguDySCD6WGD+QV3MYvO6KX+//+D+Dl+BoPIIIPoJ8HnBIvOA/jojv7//4P4OX4Gg8ggg+gng8fNwecEA/iJfjBfXsPMzMzMzFNWV4vxi04wg/kgdBOD+Ql0DoP5DXQJg/kKdAQz0usFugEAAAAzwIP5Iw+UwAvCD4RiAQAAg/kjD4XwAAAAi048hcl0IQ++AYlGMI1BAYN+MAKJRjx1NYtGQMdGPAAAAACJRjDrJotWTA++ColOMA++QgHB4AgDwYlGMHQIjUICiUZM6wfHRjD/////i87obf///4F+IDADAAB1IovO6F3///+LRhjGACCLRiD/RhjH" & _
    "AAEAAACLTiCLRhiJQQSDfjAKdFaLThiKRjCIAf9GGItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dSyLRkDHRjwAAAAAiUYw6x2LVkwPvgqJTjAPvkIBweAIA8GJRjB0T41CAolGTIN+MAp1qotOGIpGMIgB/0YYi0YYxgAC/0YYi048hcl0NA++AYlGMI1BAYN+MAKJRjwPhbr+//+LRkDHRjwAAAAAiUYw6aj+///HRjD/////6Vv///+LVkwPvgqJTjAPvkIBweAIA8GJRjB0C41CAolGTOl6/v//x0Yw/////+lu/v//g/lMdSCLRkwPvkgBD74AweEIA8GD+CJ0BYP4J3UHuAEAAADrAjPAiUYshcB0TotOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4tGMIvIx0YoAAAAAIlGIIP5YXwFg/l6fhKD+UF8" & _
    "BYP5Wn4IjUHQg/gJdwe6AQAAAOsCM9IzwIP5Xw+UwAvCD4QQAQAAi0YYxgAg/0YYi34YiX5Ei04wg/lhfAWD+Xp+EoP5QXwFg/lafgiNQdCD+Al3B7oBAAAA6wIz0jPAg/lfD5TAC8J0YYpGMIgH/0YYi048i34Yhcl0IQ++AYlGMI1BAYN+MAKJRjx1qItGQMdGPAAAAACJRjDrmYtWTA++ColOMA++QgHB4AgDwYlGMHQLjUICiUZM6Xf////HRjD/////6Wv///+LRiCDwNCD+AkPhrABAADGByCLVkSLTgRK6IcTAAArRgSJRiCLRhjGAACLRiCNDMUAAQAAiU4ggfkwAwAAD46OAQAAi0YQA8GJRiCDOAEPhX0BAACLQASLzolGPItGMIlGQOgO+///6Z78//+LTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dG" & _
    "MP////+LRiCD+CcPhJUBAAAzyYP4Lw+UwTPAg34wKg+UwIXID4T2AAAAi87oj/r//4tOMIXJD4TCAAAAi348kIP5KnRIi9eF0nQgD74KQolOMIv6iVY8g/kCdSyLTkAz0jP/iVY8iU4w6x2LXkwPvguJTjAPvkMBweAIA8iJTjB0K41DAolGTIP5KnW6hf90JQ++D0eJTjCJfjyD+QJ1M4tOQDP/iX48iU4w6ybHRjD/////65GLVkwPvgqJTjAPvkIBweAIA8iJTjB0H41CAolGTIP5L3QfhckPhWD///+LzujZ+f//6Wn7//+Dyf+JTjDpSf///8dGMAAAAACLzui7+f//6Uv7//+LTkToXhIAAIlGJMdGIAIAAABfXlvDi1YcD746hf908o2bAAAAAA++WgGDwgLHRiQAAAAAD74Kg+liiU4oeRwzwOsDjUkAQELB4AYDwYlGJA++CoPpYolOKHjrM8lCO14wD5TBM8CD+0APlMALyDPAO34gD5TA" & _
    "hch1Cw++OoX/daVfXlvDO14wdYiLzugl+f//X8dGIAEAAABeW8OLzsdGIAIAAADobPn//4tGMIvOiUYk6P/4//9fi85eW+n1+P//zMzMzMyF0nQUg/r/dA+LQRSIEP9BFMH6CIXSdezDzMzMzMzMzIXSdBRWi0EUizIrwoPoBIkCi9aF9nXuXsPMzMzMzMzMVovxuLgAAACD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQg0YUBF7DzMzMzMzMzMzMVovxuOkAAACD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQi0YUjUgEiU4UXsPMzMzMVYvsVovyuoXADwDrA41JAIP6/3QPi0EUiBD/QRTB+giF0nXsjZaEAAAAXoXSdBWQg/r/dA+LQRSIEP9BFMH6CIXSdeyLURSLRQiJAotBFI1QBIlRFF3DzMzMzMzMzMzMVovyujnBAACD+v90D4tBFIgQ/0EUwfoIhdJ17Lq4AAAAg/r/dA+LQRSIEP9BFMH6" & _
    "CIXSdeyLQRS6DwAAAMcAAAAAAINBFASD+v90D4tBFIgQ/0EUwfoIhdJ17I2WkAAAAF6F0nQUg/r/dA+LQRSIEP9BFMH6CIXSdey6wAAAAJCD+v90D4tBFIgQ/0EUwfoIhdJ17MPMzMzMzMzMzMzMzFWL7FaL8YHCgwAAAHQWi/+D+v90D4tGFIgQ/0YUwfoIhdJ17ItVCDPAgfoAAgAAD53ASCWAAAAAg8AFdBeNSQCD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQg0YUBF5dw1WL7IPsDFNWi/GJVfRXvwEAAACJffyLXiCD+yIPhUYBAACLVgy5uAAAAI2bAAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiRCDRhQEg34wInRxi87o9vb//4tODIpGMIgB/0YMg34sAItODHQLi0YwwfgIiAH/RgyLTjyFyXQhD74BiUYwjUEBg34wAolGPHUsi0ZAx0Y8AAAAAIlGMOsdi1ZMD74KiU4wD75CAcHg" & _
    "CAPBiUYwdGKNQgKJRkyDfjAidY+DfiwAdAmLRgzGAAD/RgyLRgzGAACLRgyLTjyDwAOD4PyJRgyFyXQ6D74BiUYwjUEBg34wAolGPHVbi0ZAi87HRjwAAAAAiUYw6G73//+LXfTp3wMAAMdGMP/////pLf///4tWTA++ColOMA++QgHB4AgDwYlGMHQVjUICi86JRkzoNvf//4td9OmnAwAAx0Yw/////4vO6CD3//+LXfTpkQMAAItGKIt+JIlF+OgK9///g/sCdSy5uAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiTi/AQAAAINGFATpUgMAAIH7SAIAAA+FjQAAAIvO6HYGAAC5g8ADAJCD+f90D4tGFIgI/0YUwfkIhcl17LmD4PwAjaQkAAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsuSnEAACNpCQAAAAAg/n/dA+LRhSICP9GFMH5CIXJdey5ieAAAI2kJAAAAACD+f8PhPABAACLRhSICP9G" & _
    "FMH5CIXJdeiNeQHpuQIAAIN9+AJ1STPSi87osf3//7m5AAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSL14vOxwAAAAAAg0YUBIP7IQ+FBwEAAOh5/P//jXvg6WoCAACD+yh1FovO6JUFAACLzuje9f//jXvZ6U8CAACD+yoPhUgBAACLzujG9f//i14gi87ovPX//4vO6LX1//+DfiAqdR6Lzuio9f//i87oofX//4vO6Jr1//+LzuiT9f//M9uLzuiK9f//M9KLzugB/f//g34gPQ+FjwAAAIvO6HD1//+5UAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi87oAAUAALlZAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyB+xgCAAB1JblmiQEAg/n/D4S6AAAAi0YUiAj/RhTB+QiFyXXojXkB6YMBAAAz0ovOgfsAAQAAD5TCgcKIAQAA6HL6//+/AQAAAOlhAQAAhdsPhHkAAACB+wABAAB1J7mLAAAAjUkA"
Private Const STR_THUNK2 = _
    "g/n/dDmLRhSICP9GFMH5CIXJdez/RhSNeQHpKgEAADPSi86B+xgCAAAPlcJKgeIAAQAAgcIPvgAA6BL6////RhS/AQAAAOn+AAAAg/smdSOLRiCNU+SLzv8w6JH7//+DxASLzuhn9P//vwEAAADp1gAAAIH7EAMAAHQyixOJVfyF0nUui34IiweFwHQgjUkAi1ZEi8joJgoAAIvQiVX8hdJ1D4tHBIPHBIXAdeMz0olV/ItOIDPAg/k9D5TAhUX0dCeLzugH9P//i87osAMAAIt9/IX/dHJXugYAAACLzugM+///g8QE62CD+Sh0WIXSdBNSuggAAACLzujx+v//i1X8g8QEg34oC3U7hdJ0D1Iz0ovO6Nf6//+DxATrGbmDwAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLViSLzugP+f//6Irz//+LffyDfiAoD4VtAQAAg/8BdRy5UAAAAI1JAIP5/3QPi0YUiAj/RhTB+QiFyXXsuYHsAACNpCQAAAAA" & _
    "g/n/dA+LRhSICP9GFMH5CIXJdeyLRhSLzscAAAAAAItGFIlF9IPABIlGFOgg8///M/+DfiApdD+LzujBAgAAuYmEJACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIk4g0YUBIN+ICx1B4vO6OLy//+DxwSDfiApdcGLRfSLzok46M3y//+LVfyF0nU0i1MEuegAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFF+JEItOFI1BBIlGFF6JSwRbi+Vdw4P6AXVVuf+UJACD+f90D4tGFIgI/0YUwfkIhcl17ItGFLmBxAAAiTiDRhQEjZsAAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhRfxwAEAAAAg0YUBF5bi+VdwytWFLnoAAAAg+oFg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSJEINGFARfXluL5V3DzMzMzMzMzMzMVYvsg+wIVleL+ovxi8dPg/gBdQvoZ/n//19ei+Vdw4vX6Nr////HRfgAAAAAO34o" & _
    "D4VpAQAAU4tGJIvOi14giUX86Lnx//+D/wh+XbmFwA8Ag/n/dA+LRhSICP9GFMH5CIXJdeyLTfyBwYQAAAB0FIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUi9eLTfiJCIvOi14UiV34jUMEiUYU6F////+LVfzpngAAALlQAAAAi/+D+f90D4tGFIgI/0YUwfkIhcl17IvXi87oM////7lZAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLVfwzyYP/BQ+UwTPAg/8ED5TAC8h0DIvO6Hz3//+LVfzrO4vKhdJ0FZCD+f90D4tGFIgI/0YUwfkIhcl17IP7JXUbuZIAAACL/4P5/3QPi0YUiAj/RhTB+QiFyXXsi134O34oD4Tn/v//hdt0SoP/CH5FU4vO6L/2//+LXfyDxASL04v4g/IB6E32//+6BQAAAIvO6HH2//+F/3QSi0YUiw8rx4PoBIkHi/mFyXXui9OLzugi9v//W19ei+Vdw8zMzMzMzMzMzMzM" & _
    "ugsAAADpRv7//8zMzMzMzFa6CwAAAIvx6DP+//+5hcAPAIP5/3QPi0YUiAj/RhTB+QiFyXXsuYQAAADrA41JAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUxwAAAAAAi0YUjUgEiU4UXsPMzMzMzMzMzFWL7FFTVovxi9pXi34ggf8gAQAAD4XPAAAA6MHv//+Lzui67///i87oc////4vOiUX86Knv//+L04vO6MD///+BfiA4AQAAdXWLzuiQ7///uekAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFMcAAAAAAIt+FItV/I1HBIlGFIXSdBKLRhSLCivCg+gEiQKL0YXJde6L04vO6GP///+F/w+E9AIAAItGFIsPK8eD6ASJB4v5hcl17l9eW4vlXcOLVfyF0g+E0AIAAI2kJAAAAACLRhSLCivCg+gEiQKL0YXJde5fXluL5V3DM8mB//gBAAAPlMEzwIH/YAEAAA+UwAvID4TyAAAAi87o0u7//4vO" & _
    "6Mvu//+B/2ABAAB1D4teFIvO6Hn+//+JRfzrcIN+IDt0DLoLAAAAi87oovz//4vO6Jvu//+DfiA7i14Ux0X8AAAAAHQKi87oRP7//4lF/IvO6Hru//+DfiApdDAz0ovO6Fv0//+6CwAAAIvOi/joXfz//yteFIvOjVP76ED0//+L14vO6Ofz//+NXwSLzug97v//jVX8i87oU/7//yteFLnpAAAAg+sFg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSJGINGFASLVfyF0g+EuQEAAItGFIsKK8KD6ASJAovRhcl17l9eW4vlXcOB/4ACAAAPhdYAAACLRgTGQDYgi0YExkA9IItGBMZAQiA5fiAPhZwAAACLzui27f//i0YgPbACAAB1VovO6KXt//+DfiAodQeLzuiY7f//g34gAnUti04khcl1C/9GFOsajZsAAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLzuhl7f//g34gKXUxi87oWO3//+soPegCAAB1" & _
    "IYvO6Ejt//+LzuhB7f//i87oOu3//7oLAAAAi87oLvv//4F+IIACAAAPhGT///+LRgRfxkA2IYtGBMZAPSGLRgReW8ZAQiGL5V3Dg/97dT2Lzuj57P//jVeGi87orwAAAIN+IH0PhJcAAADrA41JAIvTi87o9/z//4N+IH118YvO6Mrs//9fXluL5V3Dgf/AAQAAdTSLzui07P//g34gO3QMugsAAACLzuii+v//i1Y0i87oiPL//4vOiUY06I7s//9fXluL5V3Dgf+QAQAAdSCLzuh47P//ixOLzuhf8v//i86JA+hm7P//X15bi+Vdw4P/O3QMugsAAACLzuhO+v//i87oR+z//19eW4vlXcNVi+yD7AgzwFOL2oXbiV38VovxD5TAM8mJRfhXg34g/w+VwSPIM8CBfiAAAQAAD5TAC8gPhOwBAACNpCQAAAAAi1YggfoAAQAAD4WMAAAAi87o6uv//4N+IDt0c41kJACF23QXg0Y4BItOOItGIPfZ" & _
    "iQiLzujH6///60GLTiCLRgyJAYvO6Lbr//+DfiBbdSiLzuip6///i04ki0YMg8EDA8GLzoPg/IlGDOiR6///i87oiuv//+sEg0YMBIN+ICx1B4vO6Hfr//+DfiA7dZGLzuhq6///6SoBAACLUgSF0nQSi0YUiworwoPoBIkCi9GFyXXui04gi0YUiQGLzug96///i87oNuv//4N+ICm/CAAAAHQii0Ygi86JOIPHBOgc6///g34gLHUHi87oD+v//4N+ICl13ovO6ALr///HRjgAAAAAuVWJ5QDHRjQAAAAAg/n/dA+LRhSICP9GFMH5CIXJdey5gewAAI2bAAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUM9KLzscAAAAAAIteFI1DBIlGFOjB+v//i1Y0hdJ0EotGFIsKK8KD6ASJAovRhcl17rnJAAAAjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyDx/i5wgAAAI1kJACD+f90D4tGFIgI/0YUwfkI" & _
    "hcl17ItGFIk4g0YUAotGOIkDi138M9KDfiD/D5XCM8AjVfiBfiAAAQAAD5TAC9APhRv+//9fXluL5V3DzMzMzFWL7IPsEFZXi/mJVfyLRzyLRDh4A8eLcCCF9nUIXzPAXovlXcNTi1gcjQw+A9+JTfiJXfAz9otYJAPfiV30i1gYhdt+P4sEsYvKA8eNZCQAihE6EHUahNJ0EopRATpQAXUOg8ECg8AChNJ15DPA6wUbwIPIAYXAdBSLVfxGi034O/N8wVtfM8Bei+Vdw4tF9ItN8FsPtwRwiwSBA8dfXovlXcPMzMzMzMzMzMzMzMzMVYvsg+wMU4vaiU34VovRiV30V41LAYoDQ4TAdfmKAivZhMB0UotN9Iv6K/mKIYhl/zrEdTaL84vRhdt0F41JAA++BBeNUgEPvkr/TivBdQ6F9nXsi0X4X15bi+Vdw4XAdPKLVfiLTfSKZf+KQgFCR4lV+ITAdbpfXjPAW4vlXcNTVooZD77Dg/ggdA+D+Al0" & _
    "CoP4DXQFg/gKdQNB6+Qz9oD7MHU+gHkBeHU1g8ECi/+KEY1JAYD6MHwFgPo5fhSK2oDLII1DnzwFdzWA+jl+A41T2YPG/Q++wsHmBAPw69CA+zB8G41JAYD7OX8TD77DjTS2ihmNduiNNHCA+zB95YvGXlvDzMzMzMzMzMzMzMw="

'=========================================================================
' API
'=========================================================================

'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
Private Const MEM_RELEASE                   As Long = &H8000
'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Any, ByVal cb As Long, lpcbNeeded As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ALLOC_SIZE                    As Long = 2 ^ 16 - 32

Private Type UcsRtccBufferType
    m_data(0 To ALLOC_SIZE) As Byte
End Type

Public Type UcsRtccContextType
    m_prog              As Long
    m_sym_stk           As Long
    m_mods              As Long
    m_glo               As Long
    m_vars              As Long
    m_state(0 To 31)    As Long
    m_buffer()          As UcsRtccBufferType
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function RtccCompile(ctx As UcsRtccContextType, sSource As String) As Long
    If ctx.m_prog = 0 Then
        ctx.m_prog = VirtualAlloc(0, ALLOC_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        ReDim ctx.m_buffer(0 To 3) As UcsRtccBufferType
        ctx.m_sym_stk = VarPtr(ctx.m_buffer(0))
        ctx.m_glo = VarPtr(ctx.m_buffer(1))
        ctx.m_vars = VarPtr(ctx.m_buffer(2))
        ctx.m_mods = VarPtr(ctx.m_buffer(3))
        Call EnumProcessModules(GetCurrentProcess(), ByVal ctx.m_mods, ALLOC_SIZE \ 4, ByVal 0)
    End If
    RtccCompile = CallWindowProc(pvGetThunkAddress, VarPtr(ctx), StrPtr(sSource))
End Function

Public Sub RtccFree(ctx As UcsRtccContextType)
    Dim uEmpty          As UcsRtccContextType
    
    If ctx.m_prog <> 0 Then
        Call VirtualFree(ctx.m_prog, 0, MEM_RELEASE)
        ctx = uEmpty
    End If
End Sub

'= private ===============================================================

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    
    If lpThunk = 0 Then
        baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2)
        lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function
