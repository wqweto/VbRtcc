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

' Size=6656
Private Const STR_THUNK1 = _
    "VYvsgeyAAAAAU4tdCFaDexQAD4WNAQAAV4t7DI11gIl7HLkfAAAAx0X4JmQqQMdFvDAzZTzHRcA9MGY+x0XEPS9mPMdFyEAuZj7HRcxAMWY9x0XQPSZnIcdF1D0nZybHRdgma3x8x0XcI2wmQMdF4C5CQ2jHReReQC5Cx0XoU2l8QMdF7C5CK2rHRfB+QC8lx0X0WWQhQMdFgCsrI23HRYQtLSVhx0WIbSpAUsdFjDxeMWPHRZAvQCVbx0WUX1tIM8dFmGMlQCXHRZxbX1tIx0WgM2MrQMdFpC5CI2THRagtQCU6x0WsX15CS8dFsGQ8PFrHRbQvMDNlx0W4Pj5gL2bHRfxiAPOluQ8AAABmpYtDHI11vIt7BIPAfolDDMdFvCBpbnTHRcAgaWYgx0XEZWxzZcdFyCB3aGnHRcxsZSBix0XQcmVha8dF1CByZXTHRdh1cm4gx0XcZm9yIMdF4HNob3LHReR0IGVtx0XoaXQgZcdF7GF4IGTHRfBlZmlu" & _
    "x0X0ZSBtYWbHRfhpbsZF+iDzpWalpItDBIPAP4lDGIsDiUMUX4tVDItDPItzFIlTTIlTSIXAdB4PvghAiUswiUM8g/kCdTKLQ0DHQzwAAAAAiUMw6yMPvgqJSzAPvkIBweAIA8iJSzB0CI1CAolDTOsHx0Mw/////4vL6LABAAAz0ovL6EcUAACLxl5bi+Vdw8zMzMzMzMzMzMzMzMzMzIvRi0o8hcl0IA++AYlCMI1BAYN6MAKJQjx1OYtCQMdCPAAAAACJQjDDVotyTA++DolKMA++RgHB4AgDwYlCMHQLjUYCiUJMi0IwXsPHQjD/////XotCMMPMzMzMzMzMzFaL8YN+MFwPhRwBAACLTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dGMP////+LRjCD+G51CcdGMAoAAABew4P4cnUJx0YwDQAAAF7Dg/h0dQnH" & _
    "RjAJAAAAXsOD+Fx1BYlGMF7Dg/h4D4WOAAAAV4vO6P3+//+L+IP/OX4Gg88gg+8ni87o6f7//4P4OX4Gg8ggg+gng8fNwecEA/iDfiwAdFWLRkwPvkgBD74AweEIA8iD+TB8BYP5OX4Lg8kgg+lhg/kFdzGLzuil/v//g/g5fgaDyCCD6CfB5wSLzgP46I7+//+D+Dl+BoPIIIPoJ4PHzcHnBAP4iX4wX17DzMzMzMxTVleL8YtOMIP5IHQTg/kJdA6D+Q10CYP5CnQEM9LrBboBAAAAM8CD+SMPlMALwg+EYgEAAIP5Iw+F8AAAAItOPIXJdCEPvgGJRjCNQQGDfjACiUY8dTWLRkDHRjwAAAAAiUYw6yaLVkwPvgqJTjAPvkIBweAIA8GJRjB0CI1CAolGTOsHx0Yw/////4vO6G3///+BfiCQAgAAdSKLzuhd////i0YYxgAgi0Yg/0YYxwABAAAAi04gi0YYiUEEg34wCnRWi04YikYwiAH/RhiL" & _
    "TjyFyXQhD74BiUYwjUEBg34wAolGPHUsi0ZAx0Y8AAAAAIlGMOsdi1ZMD74KiU4wD75CAcHgCAPBiUYwdE+NQgKJRkyDfjAKdaqLThiKRjCIAf9GGItGGMYAAv9GGItOPIXJdDQPvgGJRjCNQQGDfjACiUY8D4W6/v//i0ZAx0Y8AAAAAIlGMOmo/v//x0Yw/////+lb////i1ZMD74KiU4wD75CAcHgCAPBiUYwdAuNQgKJRkzpev7//8dGMP/////pbv7//4P5THUgi0ZMD75IAQ++AMHhCAPBg/gidAWD+Cd1B7gBAAAA6wIzwIlGLIXAdE6LTjyFyXQhD74BiUYwjUEBg34wAolGPHU1i0ZAx0Y8AAAAAIlGMOsmi1ZMD74KiU4wD75CAcHgCAPBiUYwdAiNQgKJRkzrB8dGMP////+LRjCLyMdGKAAAAACJRiCD+WF8BYP5en4Sg/lBfAWD+Vp+CI1B0IP4CXcHugEAAADrAjPSM8CD+V8PlMAL" & _
    "wg+EEAEAAItGGMYAIP9GGIt+GIl+RItOMIP5YXwFg/l6fhKD+UF8BYP5Wn4IjUHQg/gJdwe6AQAAAOsCM9IzwIP5Xw+UwAvCdGGKRjCIB/9GGItOPIt+GIXJdCEPvgGJRjCNQQGDfjACiUY8daiLRkDHRjwAAAAAiUYw65mLVkwPvgqJTjAPvkIBweAIA8GJRjB0C41CAolGTOl3////x0Yw/////+lr////i0Ygg8DQg/gJD4awAQAAxgcgi1ZEi04ESugnEgAAK0YEiUYgi0YYxgAAi0YgjQzFAAEAAIlOIIH5kAIAAA+OjgEAAItGEAPBiUYggzgBD4V9AQAAi0AEi86JRjyLRjCJRkDoDvv//+me/P//i048hcl0IQ++AYlGMI1BAYN+MAKJRjx1NYtGQMdGPAAAAACJRjDrJotWTA++ColOMA++QgHB4AgDwYlGMHQIjUICiUZM6wfHRjD/////i0Ygg/gnD4SVAQAAM8mD+C8PlMEzwIN+MCoP" & _
    "lMCFyA+E9gAAAIvO6I/6//+LTjCFyQ+EwgAAAIt+PJCD+Sp0SIvXhdJ0IA++CkKJTjCL+olWPIP5AnUsi05AM9Iz/4lWPIlOMOsdi15MD74LiU4wD75DAcHgCAPIiU4wdCuNQwKJRkyD+Sp1uoX/dCUPvg9HiU4wiX48g/kCdTOLTkAz/4l+PIlOMOsmx0Yw/////+uRi1ZMD74KiU4wD75CAcHgCAPIiU4wdB+NQgKJRkyD+S90H4XJD4Vg////i87o2fn//+lp+///g8n/iU4w6Un////HRjAAAAAAi87ou/n//+lL+///i05E6P4QAACJRiTHRiACAAAAX15bw4tWHA++OoX/dPKNmwAAAAAPvloBg8ICx0YkAAAAAA++CoPpYolOKHkcM8DrA41JAEBCweAGA8GJRiQPvgqD6WKJTih46zPJQjteMA+UwTPAg/tAD5TAC8gzwDt+IA+UwIXIdQsPvjqF/3WlX15bwzteMHWIi87oJfn//1/HRiAB" & _
    "AAAAXlvDi87HRiACAAAA6Gz5//+LRjCLzolGJOj/+P//X4vOXlvp9fj//8zMzMzMhdJ0FIP6/3QPi0EUiBD/QRTB+giF0nXsw8zMzMzMzMyF0nQUVotBFIsyK8KD6ASJAovWhfZ17l7DzMzMzMzMzFaL8bi4AAAAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEINGFARew8zMzMzMzMzMzFaL8bjpAAAAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEItGFI1IBIlOFF7DzMzMzFWL7FaL8rqFwA8A6wONSQCD+v90D4tBFIgQ/0EUwfoIhdJ17I2WhAAAAF6F0nQVkIP6/3QPi0EUiBD/QRTB+giF0nXsi1EUi0UIiQKLQRSNUASJURRdw8zMzMzMzMzMzFaL8ro5wQAAg/r/dA+LQRSIEP9BFMH6CIXSdey6uAAAAIP6/3QPi0EUiBD/QRTB+giF0nXsi0EUug8AAADHAAAAAACDQRQEg/r/dA+LQRSI" & _
    "EP9BFMH6CIXSdeyNlpAAAABehdJ0FIP6/3QPi0EUiBD/QRTB+giF0nXsusAAAACQg/r/dA+LQRSIEP9BFMH6CIXSdezDzMzMzMzMzMzMzMxVi+xWi/GBwoMAAAB0Fov/g/r/dA+LRhSIEP9GFMH6CIXSdeyLVQgzwIH6AAIAAA+dwEglgAAAAIPABXQXjUkAg/j/dA+LThSIAf9GFMH4CIXAdeyLRhSJEINGFAReXcNVi+yD7BBTVovxiVXwV78BAAAAiX38i14gg/siD4VGAQAAi1YMubgAAACNmwAAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIkQg0YUBIN+MCJ0cYvO6Pb2//+LTgyKRjCIAf9GDIN+LACLTgx0C4tGMMH4CIgB/0YMi048hcl0IQ++AYlGMI1BAYN+MAKJRjx1LItGQMdGPAAAAACJRjDrHYtWTA++ColOMA++QgHB4AgDwYlGMHRijUICiUZMg34wInWPg34sAHQJi0YMxgAA" & _
    "/0YMi0YMxgAAi0YMi048g8ADg+D8iUYMhcl0Og++AYlGMI1BAYN+MAKJRjx1W4tGQIvOx0Y8AAAAAIlGMOhu9///i13w6S8DAADHRjD/////6S3///+LVkwPvgqJTjAPvkIBweAIA8GJRjB0FY1CAovOiUZM6Db3//+LXfDp9wIAAMdGMP////+Lzugg9///i13w6eECAACLRiiJRfSLRiSJRfjoB/f//4P7AnUqubgAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFItV+IkQg0YUBOmhAgAAg330AnVLM9KLzuhJ/v//ubkAAACNZCQAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSLzotV+McAAAAAAINGFASD+yEPhQIBAADoDP3//+lQAgAAg/sodROLzuh7BQAAi87odPb//+k4AgAAg/sqD4U5AQAAi87oX/b//4teIIvO6FX2//+LzuhO9v//g34gKnUei87oQfb//4vO6Dr2//+Lzugz9v//i87o" & _
    "LPb//zPbi87oI/b//zPSi87omv3//4N+ID0PhYsAAACLzugJ9v//uVAAAACNZCQAg/n/dA+LRhSICP9GFMH5CIXJdeyLzujlBAAAuVkAAACD+f90D4tGFIgI/0YUwfkIhcl17IH7GAIAAHUiuWaJAQCD+f8PhH8BAACLRhSICP9GFMH5CIXJdejpawEAADPSi86B+wABAAAPlMKBwogBAADoCvv//+lOAQAAhdsPhEYBAACB+wABAAB1IbmLAAAAg/n/dDaLRhSICP9GFMH5CIXJdez/RhTpHQEAADPSi86B+xgCAAAPlcJKgeIAAQAAgcIPvgAA6LX6////RhTp9gAAAIP7JnUei0YgjVPki87/MOg5/P//g8QEi87oD/X//+nTAAAAgftwAgAAdC+LC4lN/IXJdSuLfgiLB4XAdB2LVkSLyOh2CQAAi8iJTfyFyXUPi0cEg8cEhcB14zPJiU38i1YgM8CD+j0PlMCFRfB0J4vO6Lf0//+LzuiwAwAA"
Private Const STR_THUNK2 = _
    "i338hf90cle6BgAAAIvO6Lz7//+DxATrYIP6KHRYhcl0E1G6CAAAAIvO6KH7//+LTfyDxASDfigLdTuFyXQPUTPSi87oh/v//4PEBOsZuYPAAACD+f90D4tGFIgI/0YUwfkIhcl17ItWJIvO6L/5///oOvT//4t9/IN+ICgPhW0BAACD/wF1HLlQAAAAjUkAg/n/dA+LRhSICP9GFMH5CIXJdey5gewAAI2kJAAAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFIvOxwAAAAAAi0YUiUXwg8AEiUYU6NDz//8z/4N+ICl0P4vO6MECAAC5iYQkAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUiTiDRhQEg34gLHUHi87okvP//4PHBIN+ICl1wYtF8IvOiTjoffP//4tV/IXSdTSLUwS56AAAAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUX4kQi04UjUEEiUYUXolLBFuL5V3Dg/oBdVW5/5QkAIP5/3QPi0YU" & _
    "iAj/RhTB+QiFyXXsi0YUuYHEAACJOINGFASNmwAAAACD+f90D4tGFIgI/0YUwfkIhcl17ItGFF/HAAQAAACDRhQEXluL5V3DK1YUuegAAACD6gWD+f90D4tGFIgI/0YUwfkIhcl17ItGFIkQg0YUBF9eW4vlXcPMzMzMzMzMzMxVi+yD7AhWV4v6i/GLx0+D+AF1C+gX+v//X16L5V3Di9fo2v///8dF+AAAAAA7figPhWkBAABTi0Yki86LXiCJRfzoafL//4P/CH5duYXADwCD+f90D4tGFIgI/0YUwfkIhcl17ItN/IHBhAAAAHQUg/n/dA+LRhSICP9GFMH5CIXJdeyLRhSL14tN+IkIi86LXhSJXfiNQwSJRhToX////4tV/OmeAAAAuVAAAACL/4P5/3QPi0YUiAj/RhTB+QiFyXXsi9eLzugz////uVkAAACD+f90D4tGFIgI/0YUwfkIhcl17ItV/DPJg/8FD5TBM8CD/wQPlMALyHQMi87o" & _
    "LPj//4tV/Os7i8qF0nQVkIP5/3QPi0YUiAj/RhTB+QiFyXXsg/sldRu5kgAAAIv/g/n/dA+LRhSICP9GFMH5CIXJdeyLXfg7figPhOf+//+F23RKg/8IfkVTi87ob/f//4td/IPEBIvTi/iD8gHo/fb//7oFAAAAi87oIff//4X/dBKLRhSLDyvHg+gEiQeL+YXJde6L04vO6NL2//9bX16L5V3DzMzMzMzMzMzMzMy6CwAAAOlG/v//zMzMzMzMVroLAAAAi/HoM/7//7mFwA8Ag/n/dA+LRhSICP9GFMH5CIXJdey5hAAAAOsDjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyLRhTHAAAAAACLRhSNSASJThRew8zMzMzMzMzMVYvsUVNWi/GL2leLfiCB/yABAAAPhc8AAADocfD//4vO6Grw//+Lzuhz////i86JRfzoWfD//4vTi87owP///4F+IDgBAAB1dYvO6EDw//+56QAAAIP5/3QPi0YUiAj/" & _
    "RhTB+QiFyXXsi0YUxwAAAAAAi34Ui1X8jUcEiUYUhdJ0EotGFIsKK8KD6ASJAovRhcl17ovTi87oY////4X/D4SEAgAAi0YUiw8rx4PoBIkHi/mFyXXuX15bi+Vdw4tV/IXSD4RgAgAAjaQkAAAAAItGFIsKK8KD6ASJAovRhcl17l9eW4vlXcMzyYH/+AEAAA+UwTPAgf9gAQAAD5TAC8gPhPIAAACLzuiC7///i87oe+///4H/YAEAAHUPi14Ui87oef7//4lF/Otwg34gO3QMugsAAACLzuii/P//i87oS+///4N+IDuLXhTHRfwAAAAAdAqLzuhE/v//iUX8i87oKu///4N+ICl0MDPSi87oC/X//7oLAAAAi86L+Ohd/P//K14Ui86NU/vo8PT//4vXi87ol/T//41fBIvO6O3u//+NVfyLzuhT/v//K14UuekAAACD6wWD+f90D4tGFIgI/0YUwfkIhcl17ItGFIkYg0YUBItV/IXSD4RJAQAA" & _
    "i0YUiworwoPoBIkCi9GFyXXuX15bi+Vdw4H/SAIAAHVti87oiO7//4vO6IHu//+DfiApD4QJAQAAjaQkAAAAAIN+IAJ1IItOJIXJdQX/RhTrFIP5/3QPi0YUiAj/RhTB+QiFyXXsi87oQ+7//4N+ICx1B4vO6Dbu//+DfiApdcCLzugp7v//X15bi+Vdw4P/e3U6i87oFu7//41XhovO6KwAAACDfiB9D4SUAAAAi/+L04vO6Gf9//+DfiB9dfGLzujq7f//X15bi+Vdw4H/wAEAAHU0i87o1O3//4N+IDt0DLoLAAAAi87oEvv//4tWNIvO6Kjz//+LzolGNOiu7f//X15bi+Vdw4H/kAEAAHUgi87omO3//4sTi87of/P//4vOiQPohu3//19eW4vlXcOD/zt0DLoLAAAAi87ovvr//4vO6Gft//9fXluL5V3DVYvsg+wIM8BTi9qF24ld/FaL8Q+UwDPJiUX4V4N+IP8PlcEjyDPAgX4gAAEAAA+U" & _
    "wAvID4SsAQAAjaQkAAAAAItWIIH6AAEAAHVTi87oDu3//4N+IDt0OoXbdBCDRjgEi044i0Yg99mJCOsMi04gi0YMiQGDRgwEi87o4ez//4N+ICx1B4vO6NTs//+DfiA7dcaLzujH7P//6ScBAACLUgSF0nQSi0YUiworwoPoBIkCi9GFyXXui04gi0YUiQGLzuia7P//i87ok+z//4N+ICm/CAAAAHQii0Ygi86JOIPHBOh57P//g34gLHUHi87obOz//4N+ICl13ovO6F/s///HRjgAAAAAuVWJ5QDHRjQAAAAAg/n/dA+LRhSICP9GFMH5CIXJdey5gewAAI1JAIP5/3QPi0YUiAj/RhTB+QiFyXXsi0YUM9KLzscAAAAAAIteFI1DBIlGFOhx+///i1Y0hdJ0EotGFIsKK8KD6ASJAovRhcl17rnJAAAAjUkAg/n/dA+LRhSICP9GFMH5CIXJdeyDx/i5wgAAAI1kJACD+f90D4tGFIgI/0YUwfkI" & _
    "hcl17ItGFIk4g0YUAotGOIkDi138M9KDfiD/D5XCM8AjVfiBfiAAAQAAD5TAC9APhVv+//9fXluL5V3DzMzMzFWL7IPsEFZXi/mJVfyLRzyLRDh4A8eLcCCF9nUIXzPAXovlXcNTi1gcjQw+A9+JTfiJXfAz9otYJAPfiV30i1gYhdt+P4sEsYvKA8eNZCQAihE6EHUahNJ0EopRATpQAXUOg8ECg8AChNJ15DPA6wUbwIPIAYXAdBSLVfxGi034O/N8wVtfM8Bei+Vdw4tF9ItN8FsPtwRwiwSBA8dfXovlXcPMzMzMzMzMzMzMzMzMVYvsg+wMU4vaiU34VovRiV30V41LAYoDQ4TAdfmKAivZhMB0UotN9Iv6K/mKIYhl/zrEdTaL84vRhdt0F41JAA++BBeNUgEPvkr/TivBdQ6F9nXsi0X4X15bi+Vdw4XAdPKLVfiLTfSKZf+KQgFCR4lV+ITAdbpfXjPAW4vlXcNTVooZD77Dg/ggdA+D+Al0" & _
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
