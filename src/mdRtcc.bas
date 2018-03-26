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

' Size=6464
Private Const STR_THUNK1 = _
    "VYvsgey0AAAAU4tdCFaDexQAD4WXAQAAV4t7DI11gIl7HLkfAAAAx0WAKysjbcdFhC0tJWHHRYhtKkBSx0WMPF4xY8dFkC9AJVvHRZRfW0gzx0WYYyVAJcdFnFtfW0jHRaAzYytAx0WkLkIjZMdFqC1AJTrHRaxfXkJLx0WwZDw8WsdFtC8wM2XHRbg+PmAvx0W8MDNlPMdFwD0wZj7HRcQ9L2Y8x0XIQC5mPsdFzEAxZj3HRdA9Jmchx0XUPVwnZ8dF2CYma3zHRdx8I2wmx0XgQC5CQ8dF5GheQC7HRehCU2l8x0XsQC5CK8dF8Gp+QC/HRfQlWWQhx0X4QCZkKmbHRfxAYsZF/gDzpbkMAAAAx4VM////IGludMeFUP///yBpZiDHhVT///9lbHNlZqXHhVj///8gd2hpx4Vc////bGUgYseFYP///3JlYWuki0McjbVM////i3sEg8B/iUMMx4Vk////IHJldMeFaP///3VybiDHhWz///9mb3Ig" & _
    "x4Vw////ZGVmaceFdP///25lIG3HhXj///9haW4g86WLQwSDwDCJQxiLA4lDFF+LVQyLQzyLcxSJU0yJU0iFwHQeD74IQIlLMIlDPIP5AnUyi0NAx0M8AAAAAIlDMOsjD74KiUswD75CAcHgCAPIiUswdAiNQgKJQ0zrB8dDMP////+Ly+imAQAAM9KLy+htEwAAi8ZeW4vlXcPMzMzMzIvRi0o8hcl0IA++AYlCMI1BAYN6MAKJQjx1OYtCQMdCPAAAAACJQjDDVotyTA++DolKMA++RgHB4AgDwYlCMHQLjUYCiUJMi0IwXsPHQjD/////XotCMMPMzMzMzMzMzFaL8YN+MFwPhRwBAACLTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dGMP////+LRjCD+G51CcdGMAoAAABew4P4cnUJx0YwDQAAAF7Dg/h0dQnH" & _
    "RjAJAAAAXsOD+Fx1BYlGMF7Dg/h4D4WOAAAAV4vO6P3+//+L+IP/OX4Gg88gg+8ni87o6f7//4P4OX4Gg8ggg+gng8fNwecEA/iDfiwAdFWLRkwPvkgBD74AweEIA8iD+TB8BYP5OX4Lg8kgg+lhg/kFdzGLzuil/v//g/g5fgaDyCCD6CfB5wSLzgP46I7+//+D+Dl+BoPIIIPoJ4PHzcHnBAP4iX4wX17DzMzMzMxTVleL8YtOMIP5IHQTg/kJdA6D+Q10CYP5CnQEM9LrBboBAAAAM8CD+SMPlMALwg+EYgEAAIP5Iw+F8AAAAItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4vO6G3///+BfiAYAgAAdSKLzuhd////i0YYxgAgi0Yg/0YYxwABAAAAi04gi0YYiUEEg34wCnRWi04YikYwiAH/RhiL" & _
    "TjyFyXQhD74BiUYwjUEBg34wAolGPHUsi0ZAx0Y8AAAAAIlGMOsdi1ZMD74KiU4wD75CAcHgCAPBiUYwdE+NQgKJRkyDfjAKdaqLThiKRjCIAf9GGItGGMYAAv9GGItOPIXJdDQPvgGJRjCNQQGDfjACiUY8D4W6/v//i0ZAx0Y8AAAAAIlGMOmo/v//x0Yw/////+lb////i1ZMD74KiU4wD75CAcHgCAPBiUYwdAuNQgKJRkzpev7//8dGMP/////pbv7//4P5THUgi0ZMD75IAQ++AMHhCAPBg/gidAWD+Cd1B7gBAAAA6wIzwIlGLIXAdE6LTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dGMP////+LRjCLyMdGKAAAAACJRiCD+WF8BYP5en4Sg/lBfAWD+Vp+CI1B0IP4CXcHugEAAADrAjPSM8CD+V8PlMAL" & _
    "wg+EEAEAAItGGMYAIP9GGIt+GIl+RItOMIP5YXwFg/l6fhKD+UF8BYP5Wn4IjUHQg/gJdwe6AQAAAOsCM9IzwIP5Xw+UwAvCdGGKRjCIB/9GGItOPIt+GIXJdCEPvgGJRjCNQQGDfjACiUY8daiLRkDHRjwAAAAAiUYw65mLVkwPvgqJTjAPvkIBweAIA8GJRjB0C41CAolGTOl3////x0Yw/////+lr////i0Ygg8DQg/gJD4awAQAAxgcgi1ZEi04ESuhXEQAAK0YEiUYgi0YYxgAAi0YgjQzFAAEAAIlOIIH5GAIAAA+OlQEAAItGEAPBiUYggzgBD4WEAQAAi0AEi86JRjyLRjCJRkDoDvv//+me/P//i048hcl0IQ++AYlGMI1BAYN+MAKJRjx1NYtGQMdGPAAAAACJRjDrJotWTA++ColOMA++QgHB4AgDwYlGMHQIjUICiUZM6wfHRjD/////i0Ygg/gnD4SVAQAAM8mD+C8PlMEzwIN+MCoP" & _
    "lMCFyA+E/QAAAIvO6I/6//+LTjCFyQ+EwgAAAIt+PJCD+Sp0SIvXhdJ0IA++CkKJTjCL+olWPIP5AnUsi05AM9Iz/4lWPIlOMOsdi15MD74LiU4wD75DAcHgCAPIiU4wdCuNQwKJRkyD+Sp1uoX/dCUPvg9HiU4wiX48g/kCdTOLTkAz/4l+PIlOMOsmx0Yw/////+uRi1ZMD74KiU4wD75CAcHgCAPIiU4wdB+NQgKJRkyD+S90H4XJD4Vg////i87o2fn//+lp+///g8n/iU4w6Un////HRjAAAAAAi87ou/n//+lL+///agBqAP92ROgqEAAAg8QMiUYkx0YgAgAAAF9eW8OLVhwPvjqF/3TyD75aAYPCAsdGJAAAAAAPvgqD6WKJTih5GzPAjWQkAEBCweAGA8GJRiQPvgqD6WKJTih46zPJQjteMA+UwTPAg/tAD5TAC8gzwDt+IA+UwIXIdQsPvjqF/3WmX15bwzteMHWPi87oJfn//1/HRiAB" & _
    "AAAAXlvDi87HRiACAAAA6Gz5//+LRjCLzolGJOj/+P//X4vOXlvp9fj//8zMzMzMhdJ0FIP6/3QPi0EUiBD/QRTB+giF0nXsw8zMzMzMzMyF0nQUVotBFIsyK8KD6ASJAovWhfZ17l7DzMzMzMzMzFaL8bi4AAAAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEINGFARew8zMzMzMzMzMzFaL8bjpAAAAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEItGFI1IBIlOFF7DzMzMzFWL7FaL8rqFwA8A6wONSQCD+v90D4tBFIgQ/0EUwfoIhdJ17I2WhAAAAF6F0nQVkIP6/3QPi0EUiBD/QRTB+giF0nXsi1EUi0UIiQKLQRSNUASJURRdw8zMzMzMzMzMzFaL8ro5wQAAg/r/dA+LQRSIEP9BFMH6CIXSdey6uAAAAIP6/3QPi0EUiBD/QRTB+giF0nXsi0EUug8AAADHAAAAAACDQRQEg/r/dA+LQRSI" & _
    "EP9BFMH6CIXSdeyNlpAAAABehdJ0FIP6/3QPi0EUiBD/QRTB+giF0nXsusAAAACQg/r/dA+LQRSIEP9BFMH6CIXSdezDzMzMzMzMzMzMzMxVi+xWi/GBwoMAAAB0Fov/g/r/dA+LRhSIEP9GFMH6CIXSdeyLVQgzwIH6AAIAAA+dwEglgAAAAIPABXQXjUkAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEINGFAReXcNVi+yD7AxTVovxiVX0V78BAAAAiX38i14gg/siD4VGAQAAi1YMubgAAACNmwAAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIkQg0YUBIN+MCJ0cYvO6Pb2//+LTgyKRjCIAf9GDIN+LACLTgx0C4tGMMH4CIgB/0YMi048hcl0IQ++AYlGMI1BAYN+MAKJRjx1LItGQMdGPAAAAACJRjDrHYtWTA++ColOMA++QgHB4AgDwYlGMHRijUICiUZMg34wInWPg34sAHQJi0YMxgAA" & _
    "/0YMi0YMxgAAi0YMi048g8ADg+D8iUYMhcl0Og++AYlGMI1BAYN+MAKJRjx1W4tGQIvOx0Y8AAAAAIlGMOhu9///i1306eoAAADHRjD/////6S3///+LVkwPvgqJTjAPvkIBweAIA8GJRjB0FY1CAovOiUZM6Db3//+LXfTpsgAAAMdGMP////+Lzugg9///i1306ZwAAACLRiiLfiSJRfjoCvf//4P7AnUkubgAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIk4g0YUBOtgg334AnVDM9KLzuhS/v//ubkAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIvXi87HAAAAAACDRhQEg/shD4XLAQAA6Br9///rF4P7KA+FDQEAAIvO6CgFAACLzuiB9v//vwEAAACDfiAoD4VPAwAAg/8BdRm5UAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsuYHsAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIvOxwAAAAAA" & _
    "i0YUiUX0g8AEiUYU6B/2//8z/4N+ICl0So2kJAAAAACLzuipBAAAuYmEJACNZCQAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSJOINGFASDfiAsdQeLzujW9f//g8cEg34gKXW9i0X0i86JOOjB9f//i1X8hdIPhRICAACLUwS56AAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUX4kQi04UjUEEiUYUXolLBFuL5V3Dg/sqD4UGAQAAi87ocvX//4teIIvO6Gj1//+Lzuhh9f//g34gKnUei87oVPX//4vO6E31//+LzuhG9f//i87oP/X//zPbi87oNvX//zPSi87orfz//4N+ID11XYvO6CD1//+5UAAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi87ooAMAALlZAAAAg/n/dA+LRhSICP9GFMH5CIXJdewz0ovOgfsAAQAAD5TCgcKIAQAA6E/6///pSf7//4XbD4RB/v//gfsAAQAAdSa5iwAAAOsDjUkA"
Private Const STR_THUNK2 = _
    "g/n/dDCLRhSICP9GFMH5CIXJdez/RhTpE/7//7kPvgAAg/n/dA+LRhSICP9GFMH5CIXJdez/RhTp8v3//4P7JnUXi0YgjVPki87/MOh/+///g8QE6c/9//+LO4l9/IX/dS6LfgiLB4XAdCDrA41JAItWRIvI6PYHAACJRfyFwHVCi0cEg8cEhcB15TP/iX38i04gM8CD+T0PlMCFRfR0KIvO6An0//+LzuiiAgAAV7oGAAAAi87oFfv//4PEBOlx/f//i33868iD+SgPhGP9//9XuggAAACLzujy+v//g8QEg34oCw+FSf3//1cz0ovO6Nv6//+LViSDxASLzugu+f//6Knz///pKP3//4P6AXVPuf+UJACD+f90D4tGFIgI/0YUwfkIhcl17ItGFLmBxAAAiTiDRhQEg/n/dA+LRhSICP9GFMH5CIXJdeyLRhRfxwAEAAAAg0YUBF5bi+VdwytWFLnoAAAAg+oF6wONSQCD+f90D4tGFIgI/0YUwfkI" & _
    "hcl17ItGFIkQg0YUBF9eW4vlXcPMzMzMzMzMzMzMzMxVi+yD7AhWV4v6i/GLx0+D+AF1C+h3+v//X16L5V3Di9fo2v///8dF+AAAAAA7figPhWkBAABTi0Yki86LXiCJRfzoyfL//4P/CH5duYXADwCD+f90D4tGFIgI/0YUwfkIhcl17ItN/IHBhAAAAHQUg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSL14tN+IkIi86LXhSJXfiNQwSJRhToX////4tV/OmeAAAAuVAAAACL/4P5/3QPi0YUiAj/RhTB+QiFyXXsi9eLzugz////uVkAAACD+f90D4tGFIgI/0YUwfkIhcl17ItV/DPJg/8FD5TBM8CD/wQPlMALyHQMi87ojPj//4tV/Os7i8qF0nQVkIP5/3QPi0YUiAj/RhTB+QiFyXXsg/sldRu5kgAAAIv/g/n/dA+LRhSICP9GFMH5CIXJdeyLXfg7figPhOf+//+F23RKg/8IfkVTi87oz/f/" & _
    "/4td/IPEBIvTi/iD8gHoXff//7oFAAAAi87ogff//4X/dBKLRhSLDyvHg+gEiQeL+YXJde6L04vO6DL3//9bX16L5V3DzMzMzMzMzMzMzMy6CwAAAOlG/v//zMzMzMzMVroLAAAAi/HoM/7//7mFwA8Ag/n/dA+LRhSICP9GFMH5CIXJdey5hAAAAOsDjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhTHAAAAAACLRhSNSASJThRew8zMzMzMzMzMVYvsUVNWi/GL2leLfiCB/yABAAAPhc8AAADo0fD//4vO6Mrw//+Lzuhz////i86JRfzoufD//4vTi87owP///4F+IDgBAAB1dYvO6KDw//+56QAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUxwAAAAAAi34Ui1X8jUcEiUYUhdJ0EotGFIsKK8KD6ASJAovRhcl17ovTi87oY////4X/D4QUAgAAi0YUiw8rx4PoBIkHi/mFyXXuX15bi+Vdw4tV" & _
    "/IXSD4TwAQAAjaQkAAAAAItGFIsKK8KD6ASJAovRhcl17l9eW4vlXcMzyYH/+AEAAA+UwTPAgf9gAQAAD5TAC8gPhPIAAACLzuji7///i87o2+///4H/YAEAAHUPi14Ui87oef7//4lF/Otwg34gO3QMugsAAACLzuii/P//i87oq+///4N+IDuLXhTHRfwAAAAAdAqLzuhE/v//iUX8i87oiu///4N+ICl0MDPSi87oa/X//7oLAAAAi86L+Ohd/P//K14Ui86NU/voUPX//4vXi87o9/T//41fBIvO6E3v//+NVfyLzuhT/v//K14UuekAAACD6wWD+f90D4tGFIgI/0YUwfkIhcl17ItGFIkYg0YUBItV/IXSD4TZAAAAi0YUiworwoPoBIkCi9GFyXXuX15bi+Vdw4P/e3U/i87o6+7//41XhovO6LEAAACDfiB9D4SZAAAAjaQkAAAAAIvTi87o1/3//4N+IH118YvO6Lru//9fXluL5V3Dgf/A" & _
    "AQAAdTSLzuik7v//g34gO3QMugsAAACLzuiC+///i1Y0i87oePT//4vOiUY06H7u//9fXluL5V3Dgf+QAQAAdSCLzuho7v//ixOLzuhP9P//i86JA+hW7v//X15bi+Vdw4P/O3QMugsAAACLzugu+///i87oN+7//19eW4vlXcNVi+yD7AgzwFOL2oXbiV38VovxD5TAM8mJRfhXg34g/w+VwSPIM8CBfiAAAQAAD5TAC8gPhKwBAACNpCQAAAAAi1YggfoAAQAAdVOLzuje7f//g34gO3Q6hdt0EINGOASLTjiLRiD32YkI6wyLTiCLRgyJAYNGDASLzuix7f//g34gLHUHi87opO3//4N+IDt1xovO6Jft///pJwEAAItSBIXSdBKLRhSLCivCg+gEiQKL0YXJde6LTiCLRhSJAYvO6Grt//+Lzuhj7f//g34gKb8IAAAAdCKLRiCLzok4g8cE6Ent//+DfiAsdQeLzug87f//g34gKXXei87oL+3/" & _
    "/8dGOAAAAAC5VYnlAMdGNAAAAACD+f90D4tGFIgI/0YUwfkIhcl17LmB7AAAjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhQz0ovOxwAAAAAAi14UjUMEiUYU6OH7//+LVjSF0nQSi0YUiworwoPoBIkCi9GFyXXuuckAAACNSQCD+f90D4tGFIgI/0YUwfkIhcl17IPH+LnCAAAAjWQkAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiTiDRhQCi0Y4iQOLXfwz0oN+IP8PlcIzwCNV+IF+IAABAAAPlMAL0A+FW/7//19eW4vlXcPMzMzMVYvsg+wQVleL+YlV/ItHPItEOHgDx4twIIX2dQhfM8Bei+Vdw1OLWByNDD4D34lN+Ild8DP2i1gkA9+JXfSLWBiF234/iwSxi8oDx41kJACKEToQdRqE0nQSilEBOlABdQ6DwQKDwAKE0nXkM8DrBRvAg8gBhcB0FItV/EaLTfg783zBW18zwF6L5V3Di0X0" & _
    "i03wWw+3BHCLBIEDx19ei+Vdw8zMzMzMzMzMzMzMzMxVi+yD7AxTi9qJTfhWi9GJXfRXjUsBigNDhMB1+YoCK9mEwHRSi030i/or+YohiGX/OsR1Novzi9GF23QXjUkAD74EF41SAQ++Sv9OK8F1DoX2deyLRfhfXluL5V3DhcB08otV+ItN9Ipl/4pCAUJHiVX4hMB1ul9eM8Bbi+Vdw1WL7ItFCFNWihgPvsuD+SB0D4P5CXQKg/kNdAWD+Qp1A0Dr5DP2gPswdT2AeAF4dTSDwAKKEI1AAYD6MHwFgPo5fhWK2oDLII1Ln4D5BXc7gPo5fgONU9mDxv0PvsrB5gQD8evPgPswfCGNmwAAAACNQAGA+zl/Ew++y400tooYjXbojTRxgPswfeWLxl5bXcPMzMzMzMzMzMzMzMzMzMw="

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