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

' Size=6640
Private Const STR_THUNK1 = _
    "VYvsgey8AAAAU4tdCFaDexQAD4W7AQAAV4t7DI21RP///4l7HLkfAAAAx4VE////KysjbceFSP///y0tJWHHhUz///9tKkBSx4VQ////PF4xY8eFVP///y9AJVvHhVj///9fW0gzx4Vc////YyVAJceFYP///1tfW0jHhWT///8zYytAx4Vo////LkIjZMeFbP///y1AJTrHhXD///9fXkJLx4V0////ZDw8WseFeP///y8wM2XHhXz///8+PmAvx0WAMDNlPMdFhD0wZj7HRYg9L2Y8x0WMQC5mPsdFkEAxZj3HRZQ9Jmchx0WYPVwnZ8dFnCYma3zHRaB8I2wmx0WkQC5CQ8dFqGheQC7HRaxCU2l8x0WwQC5CK8dFtGp+QC/HRbglWWQhx0W8QCZkKmbHRcBAYsZFwgDzpcdFxCBpbnS5DgAAAMdFyCBpZiDHRcxlbHNlZqXHRdAgd2hpx0XUbGUgYsdF2HJlYWuki0McjXXEi3sEg8B/iUMMx0Xc" & _
    "IHJldMdF4HVybiDHReRmb3Igx0Xoc2hvcsdF7HQgZW3HRfBpdCBkx0X0ZWZpbsdF+GUgbWFmx0X8aW7GRf4g86VmpaSLQwSDwDuJQxiLA4lDFF+LVQyLQzyLcxSJU0yJU0iFwHQeD74IQIlLMIlDPIP5AnUyi0NAx0M8AAAAAIlDMOsjD74KiUswD75CAcHgCAPIiUswdAiNQgKJQ0zrB8dDMP////+Ly+iiAQAAM9KLy+gJFAAAi8ZeW4vlXcPMi9GLSjyFyXQgD74BiUIwjUEBg3owAolCPHU5i0JAx0I8AAAAAIlCMMNWi3JMD74OiUowD75GAcHgCAPBiUIwdAuNRgKJQkyLQjBew8dCMP////9ei0Iww8zMzMzMzMzMVovxg34wXA+FHAEAAItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4tGMIP4" & _
    "bnUJx0YwCgAAAF7Dg/hydQnHRjANAAAAXsOD+HR1CcdGMAkAAABew4P4XHUFiUYwXsOD+HgPhY4AAABXi87o/f7//4v4g/85fgaDzyCD7yeLzujp/v//g/g5fgaDyCCD6CeDx83B5wQD+IN+LAB0VYtGTA++SAEPvgDB4QgDyIP5MHwFg/k5fguDySCD6WGD+QV3MYvO6KX+//+D+Dl+BoPIIIPoJ8HnBIvOA/jojv7//4P4OX4Gg8ggg+gng8fNwecEA/iJfjBfXsPMzMzMzFNWV4vxi04wg/kgdBOD+Ql0DoP5DXQJg/kKdAQz0usFugEAAAAzwIP5Iw+UwAvCD4RiAQAAg/kjD4XwAAAAi048hcl0IQ++AYlGMI1BAYN+MAKJRjx1NYtGQMdGPAAAAACJRjDrJotWTA++ColOMA++QgHB4AgDwYlGMHQIjUICiUZM6wfHRjD/////i87obf///4F+IHACAAB1IovO6F3///+LRhjGACCLRiD/RhjH" & _
    "AAEAAACLTiCLRhiJQQSDfjAKdFaLThiKRjCIAf9GGItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dSyLRkDHRjwAAAAAiUYw6x2LVkwPvgqJTjAPvkIBweAIA8GJRjB0T41CAolGTIN+MAp1qotOGIpGMIgB/0YYi0YYxgAC/0YYi048hcl0NA++AYlGMI1BAYN+MAKJRjwPhbr+//+LRkDHRjwAAAAAiUYw6aj+///HRjD/////6Vv///+LVkwPvgqJTjAPvkIBweAIA8GJRjB0C41CAolGTOl6/v//x0Yw/////+lu/v//g/lMdSCLRkwPvkgBD74AweEIA8GD+CJ0BYP4J3UHuAEAAADrAjPAiUYshcB0TotOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4tGMIvIx0YoAAAAAIlGIIP5YXwFg/l6fhKD+UF8" & _
    "BYP5Wn4IjUHQg/gJdwe6AQAAAOsCM9IzwIP5Xw+UwAvCD4QQAQAAi0YYxgAg/0YYi34YiX5Ei04wg/lhfAWD+Xp+EoP5QXwFg/lafgiNQdCD+Al3B7oBAAAA6wIz0jPAg/lfD5TAC8J0YYpGMIgH/0YYi048i34Yhcl0IQ++AYlGMI1BAYN+MAKJRjx1qItGQMdGPAAAAACJRjDrmYtWTA++ColOMA++QgHB4AgDwYlGMHQLjUICiUZM6Xf////HRjD/////6Wv///+LRiCDwNCD+AkPhrABAADGByCLVkSLTgRK6PcRAAArRgSJRiCLRhjGAACLRiCNDMUAAQAAiU4ggflwAgAAD46OAQAAi0YQA8GJRiCDOAEPhX0BAACLQASLzolGPItGMIlGQOgO+///6Z78//+LTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dG" & _
    "MP////+LRiCD+CcPhJUBAAAzyYP4Lw+UwTPAg34wKg+UwIXID4T2AAAAi87oj/r//4tOMIXJD4TCAAAAi348kIP5KnRIi9eF0nQgD74KQolOMIv6iVY8g/kCdSyLTkAz0jP/iVY8iU4w6x2LXkwPvguJTjAPvkMBweAIA8iJTjB0K41DAolGTIP5KnW6hf90JQ++D0eJTjCJfjyD+QJ1M4tOQDP/iX48iU4w6ybHRjD/////65GLVkwPvgqJTjAPvkIBweAIA8iJTjB0H41CAolGTIP5L3QfhckPhWD///+LzujZ+f//6Wn7//+Dyf+JTjDpSf///8dGMAAAAACLzui7+f//6Uv7//+LTkTozhAAAIlGJMdGIAIAAABfXlvDi1YcD746hf908o2bAAAAAA++WgGDwgLHRiQAAAAAD74Kg+liiU4oeRwzwOsDjUkAQELB4AYDwYlGJA++CoPpYolOKHjrM8lCO14wD5TBM8CD+0APlMALyDPAO34gD5TA" & _
    "hch1Cw++OoX/daVfXlvDO14wdYiLzugl+f//X8dGIAEAAABeW8OLzsdGIAIAAADobPn//4tGMIvOiUYk6P/4//9fi85eW+n1+P//zMzMzMyF0nQUg/r/dA+LQRSIEP9BFMH6CIXSdezDzMzMzMzMzIXSdBRWi0EUizIrwoPoBIkCi9aF9nXuXsPMzMzMzMzMVovxuLgAAACD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQg0YUBF7DzMzMzMzMzMzMVovxuOkAAACD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQi0YUjUgEiU4UXsPMzMzMVYvsVovyuoXADwDrA41JAIP6/3QPi0EUiBD/QRTB+giF0nXsjZaEAAAAXoXSdBWQg/r/dA+LQRSIEP9BFMH6CIXSdeyLURSLRQiJAotBFI1QBIlRFF3DzMzMzMzMzMzMVovyujnBAACD+v90D4tBFIgQ/0EUwfoIhdJ17Lq4AAAAg/r/dA+LQRSIEP9BFMH6" & _
    "CIXSdeyLQRS6DwAAAMcAAAAAAINBFASD+v90D4tBFIgQ/0EUwfoIhdJ17I2WkAAAAF6F0nQUg/r/dA+LQRSIEP9BFMH6CIXSdey6wAAAAJCD+v90D4tBFIgQ/0EUwfoIhdJ17MPMzMzMzMzMzMzMzFWL7FaL8YHCgwAAAHQWi/+D+v90D4tGFIgQ/0YUwfoIhdJ17ItVCDPAgfoAAgAAD53ASCWAAAAAg8AFdBeNSQCD+P90D4tOFIgB/0YUwfgIhcB17ItGFIkQg0YUBF5dw1WL7IPsDFNWi/GJVfRXvwEAAACJffyLXiCD+yIPhUYBAACLVgy5uAAAAI2bAAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiRCDRhQEg34wInRxi87o9vb//4tODIpGMIgB/0YMg34sAItODHQLi0YwwfgIiAH/RgyLTjyFyXQhD74BiUYwjUEBg34wAolGPHUsi0ZAx0Y8AAAAAIlGMOsdi1ZMD74KiU4wD75CAcHg" & _
    "CAPBiUYwdGKNQgKJRkyDfjAidY+DfiwAdAmLRgzGAAD/RgyLRgzGAACLRgyLTjyDwAOD4PyJRgyFyXQ6D74BiUYwjUEBg34wAolGPHVbi0ZAi87HRjwAAAAAiUYw6G73//+LXfTp6gAAAMdGMP/////pLf///4tWTA++ColOMA++QgHB4AgDwYlGMHQVjUICi86JRkzoNvf//4td9OmyAAAAx0Yw/////4vO6CD3//+LXfTpnAAAAItGKIt+JIlF+OgK9///g/sCdSS5uAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiTiDRhQE62CDffgCdUMz0ovO6FL+//+5uQAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUi9eLzscAAAAAAINGFASD+yEPhQACAADoGv3//+sXg/soD4UNAQAAi87oWAUAAIvO6IH2//+/AQAAAIN+ICgPhYIDAACD/wF1GblQAAAAg/n/dA+LRhSICP9GFMH5CIXJdey5gewA" & _
    "AIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUi87HAAAAAACLRhSJRfSDwASJRhToH/b//zP/g34gKXRKjaQkAAAAAIvO6NkEAAC5iYQkAI1kJACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIk4g0YUBIN+ICx1B4vO6Nb1//+DxwSDfiApdb2LRfSLzok46MH1//+LVfyF0g+FQwIAAItTBLnoAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhRfiRCLThSNQQSJRhReiUsEW4vlXcOD+yoPhTwBAACLzuhy9f//i14gi87oaPX//4vO6GH1//+DfiAqdR6LzuhU9f//i87oTfX//4vO6Eb1//+Lzug/9f//M9uLzug29f//M9KLzuit/P//g34gPQ+FjgAAAIvO6Bz1//+5UAAAAI2kJAAAAACD+f90D4tGFIgI/0YUwfkIhcl17IvO6MUDAAC5WQAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsgfsYAgAAdSK5ZokB"
Private Const STR_THUNK2 = _
    "AIP5/w+ERf7//4tGFIgI/0YUwfkIhcl16Okx/v//M9KLzoH7AAEAAA+UwoHCiAEAAOga+v//6RT+//+F2w+EDP7//4H7AAEAAHUhuYsAAACD+f90NotGFIgI/0YUwfkIhcl17P9GFOnj/f//M9KLzoH7GAIAAA+VwkqB4gABAACBwg++AADoxfn///9GFOm8/f//g/smdReLRiCNU+SLzv8w6En7//+DxATpmf3//4s7iX38hf91KYt+CIsHhcB0G4tWRIvI6GUIAACJRfyFwHVCi0cEg8cEhcB15TP/iX38i04gM8CD+T0PlMCFRfR0KIvO6Njz//+LzuihAgAAV7oGAAAAi87o5Pr//4PEBOlA/f//i33868iD+SgPhDL9//9XuggAAACLzujB+v//g8QEg34oCw+FGP3//1cz0ovO6Kr6//+LViSDxASLzuj9+P//6Hjz///p9/z//4P6AXVWuf+UJACD+f90D4tGFIgI/0YUwfkIhcl17ItGFLmB" & _
    "xAAAiTiDRhQEjaQkAAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUX8cABAAAAINGFAReW4vlXcMrVhS56AAAAIPqBYP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiRCDRhQEX15bi+Vdw8zMzMzMzMzMzFWL7IPsCFZXi/qL8YvHT4P4AXUL6Ef6//9fXovlXcOL1+ja////x0X4AAAAADt+KA+FaQEAAFOLRiSLzoteIIlF/OiZ8v//g/8Ifl25hcAPAIP5/3QPi0YUiAj/RhTB+QiFyXXsi038gcGEAAAAdBSD+f90D4tGFIgI/0YUwfkIhcl17ItGFIvXi034iQiLzoteFIld+I1DBIlGFOhf////i1X86Z4AAAC5UAAAAIv/g/n/dA+LRhSICP9GFMH5CIXJdeyL14vO6DP///+5WQAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi1X8M8mD/wUPlMEzwIP/BA+UwAvIdAyLzuhc+P//i1X86zuLyoXSdBWQ" & _
    "g/n/dA+LRhSICP9GFMH5CIXJdeyD+yV1G7mSAAAAi/+D+f90D4tGFIgI/0YUwfkIhcl17Itd+Dt+KA+E5/7//4XbdEqD/wh+RVOLzuif9///i138g8QEi9OL+IPyAegt9///ugUAAACLzuhR9///hf90EotGFIsPK8eD6ASJB4v5hcl17ovTi87oAvf//1tfXovlXcPMzMzMzMzMzMzMzLoLAAAA6Ub+///MzMzMzMxWugsAAACL8egz/v//uYXADwCD+f90D4tGFIgI/0YUwfkIhcl17LmEAAAA6wONSQCD+f90D4tGFIgI/0YUwfkIhcl17ItGFMcAAAAAAItGFI1IBIlOFF7DzMzMzMzMzMxVi+xRU1aL8YvaV4t+IIH/IAEAAA+FzwAAAOih8P//i87omvD//4vO6HP///+LzolF/OiJ8P//i9OLzujA////gX4gOAEAAHV1i87ocPD//7npAAAAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhTHAAAA" & _
    "AACLfhSLVfyNRwSJRhSF0nQSi0YUiworwoPoBIkCi9GFyXXui9OLzuhj////hf8PhIQCAACLRhSLDyvHg+gEiQeL+YXJde5fXluL5V3Di1X8hdIPhGACAACNpCQAAAAAi0YUiworwoPoBIkCi9GFyXXuX15bi+VdwzPJgf/4AQAAD5TBM8CB/2ABAAAPlMALyA+E8gAAAIvO6LLv//+Lzuir7///gf9gAQAAdQ+LXhSLzuh5/v//iUX863CDfiA7dAy6CwAAAIvO6KL8//+Lzuh77///g34gO4teFMdF/AAAAAB0CovO6ET+//+JRfyLzuha7///g34gKXQwM9KLzug79f//ugsAAACLzov46F38//8rXhSLzo1T++gg9f//i9eLzujH9P//jV8Ei87oHe///41V/IvO6FP+//8rXhS56QAAAIPrBYP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiRiDRhQEi1X8hdIPhEkBAACLRhSLCivCg+gEiQKL0YXJ" & _
    "de5fXluL5V3Dgf9IAgAAdW2Lzui47v//i87ose7//4N+ICkPhAkBAACNpCQAAAAAg34gAnUgi04khcl1Bf9GFOsUg/n/dA+LRhSICP9GFMH5CIXJdeyLzuhz7v//g34gLHUHi87oZu7//4N+ICl1wIvO6Fnu//9fXluL5V3Dg/97dTqLzuhG7v//jVeGi87orAAAAIN+IH0PhJQAAACL/4vTi87oZ/3//4N+IH118YvO6Bru//9fXluL5V3Dgf/AAQAAdTSLzugE7v//g34gO3QMugsAAACLzugS+///i1Y0i87o2PP//4vOiUY06N7t//9fXluL5V3Dgf+QAQAAdSCLzujI7f//ixOLzuiv8///i86JA+i27f//X15bi+Vdw4P/O3QMugsAAACLzui++v//i87ol+3//19eW4vlXcNVi+yD7AgzwFOL2oXbiV38VovxD5TAM8mJRfhXg34g/w+VwSPIM8CBfiAAAQAAD5TAC8gPhKwBAACNpCQAAAAA" & _
    "i1YggfoAAQAAdVOLzug+7f//g34gO3Q6hdt0EINGOASLTjiLRiD32YkI6wyLTiCLRgyJAYNGDASLzugR7f//g34gLHUHi87oBO3//4N+IDt1xovO6Pfs///pJwEAAItSBIXSdBKLRhSLCivCg+gEiQKL0YXJde6LTiCLRhSJAYvO6Mrs//+LzujD7P//g34gKb8IAAAAdCKLRiCLzok4g8cE6Kns//+DfiAsdQeLzuic7P//g34gKXXei87oj+z//8dGOAAAAAC5VYnlAMdGNAAAAACD+f90D4tGFIgI/0YUwfkIhcl17LmB7AAAjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhQz0ovOxwAAAAAAi14UjUMEiUYU6HH7//+LVjSF0nQSi0YUiworwoPoBIkCi9GFyXXuuckAAACNSQCD+f90D4tGFIgI/0YUwfkIhcl17IPH+LnCAAAAjWQkAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiTiDRhQCi0Y4" & _
    "iQOLXfwz0oN+IP8PlcIzwCNV+IF+IAABAAAPlMAL0A+FW/7//19eW4vlXcPMzMzMVYvsg+wQVleL+YlV/ItHPItEOHgDx4twIIX2dQhfM8Bei+Vdw1OLWByNDD4D34lN+Ild8DP2i1gkA9+JXfSLWBiF234/iwSxi8oDx41kJACKEToQdRqE0nQSilEBOlABdQ6DwQKDwAKE0nXkM8DrBRvAg8gBhcB0FItV/EaLTfg783zBW18zwF6L5V3Di0X0i03wWw+3BHCLBIEDx19ei+Vdw8zMzMzMzMzMzMzMzMxVi+yD7AxTi9qJTfhWi9GJXfRXjUsBigNDhMB1+YoCK9mEwHRSi030i/or+YohiGX/OsR1Novzi9GF23QXjUkAD74EF41SAQ++Sv9OK8F1DoX2deyLRfhfXluL5V3DhcB08otV+ItN9Ipl/4pCAUJHiVX4hMB1ul9eM8Bbi+Vdw1NWihkPvsOD+CB0D4P4CXQKg/gNdAWD+Ap1A0Hr5DP2" & _
    "gPswdT6AeQF4dTWDwQKL/4oRjUkBgPowfAWA+jl+FIragMsgjUOfPAV3NYD6OX4DjVPZg8b9D77CweYEA/Dr0ID7MHwbjUkBgPs5fxMPvsONNLaKGY126I00cID7MH3li8ZeW8PMzMzMzMzMzMzMzA=="

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

Private Const ALLOC_SIZE                    As Long = 10000

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
