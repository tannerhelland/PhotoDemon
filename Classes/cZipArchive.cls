VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This code was originally written by Vladimir Vissoultchev.
'
'Thank you to Vlad for his many contributions to the VB6 community,
' including this all-in-one zip archive manager.
'
'This file has been modified for use within PhotoDemon.  You can download an original,
' unmodified copy from the following link (good as of January 2023):
' https://github.com/wqweto/ZipArchive
'
'This class is MIT-licensed and Copyright 2017-2026 by wqweto@gmail.com.
' Full license details available at https://github.com/wqweto/ZipArchive/blob/master/LICENSE
'
'Original code follows, with any PhotoDemon-specific changes clearly noted at the end of the file.

'=========================================================================
'
' cZipArchive (c) 2017-2022 by wqweto@gmail.com
'
' A single-class pure VB6 library for zip archives management
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cZipArchive"

#Const ImplCompress = ZIP_NOCOMPRESS = 0
#Const ImplDecompress = ZIP_NOEXTRACT = 0
#Const ImplCrypto = ZIP_CRYPTO <> 0
#Const ImplUseShared = ZIP_USESHARED <> 0
#Const ImplInflate = ZIP_INFLATE <> 0
#Const ImplLogging = DEBUG_MODE

'=========================================================================
' Public enums
'=========================================================================

Public Enum ZipFileInfoIndexEnum
    zipIdxFileName = 0
    zipIdxAttributes
    zipIdxCrc32
    zipIdxSize
    zipIdxCompressedSize
    zipIdxComment
    zipIdxLastModified
    zipIdxMethod
    zipIdxOffset
    zipIdxFlags
End Enum

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Currency, ByVal Total As Currency, Cancel As Boolean)
Event Error(ByVal FileIdx As Long, Source As String, Description As String, Cancel As Boolean)
#If ImplCompress Then
    Event BeforeCompress(ByVal FileIdx As Long, Level As Long, SkipFile As Boolean, Cancel As Boolean)
#End If ' ImplCompress
#If ImplDecompress Then
    Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, Cancel As Boolean)
    Event ExtractComplete(ByVal FileIdx As Long, File As Variant)
#End If ' ImplDecompress

'=========================================================================
' Thunk data
'=========================================================================

' Auto-generated on 12.1.2018 17:15:52, CodeSize=6640, DataSize=984, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "UYtEJAhTi1wkEFWLbCQYVleLeEQD64tEJCSJbCQQiwD2wwN0HDvddBQPthNDD7bIM9HB6AgzBJf2wwN16IlcJByL1cdEJBgAAAAAK9OD4vyNDBqDwgPB6gI7yxvJ99EjyolMJCB0eYvpjaQkAAAAAIsbi8vB6QgPtvGLyMHpCA+2yTPxi8vB6RAPttGLyMHpEIu0twAIAAAPtskz0YvIwekYM7SXAAQAAIvTweoYM9EPtsgPtsOLXCQcM8iDwwQzNJeJXCQcM7SPAAwAAItMJBiLxkGJTCQYO811lItsJBCLzTP2K8s76xvt99Uj6XQWD7YTjVsBD7bIRjPRwegIMwSXO/V16otMJCRfXl2JAVtZwhAAzMzMzMzMzMzMzMzMVot0JBCF9n41i1QkDE6LRCQIwe4ERoMCAYsKdQP/QgSJCItKBIlIBMdACAAAAADHQAwAAAAAg8AQg+4BddhewhAAzMzMzMzMzMzMzMzMzMyLVCQMhdJ0GotEJAhWi3Qk" & _
    "CCvwigwGjUABMEj/g+oBdfJewhAAzMzMzMzMzMzMzMyDfCQMAItEJASLSAyJTCQED46OAAAAU1VWi3QkFFeLfCQgi++D5QGD5wKNmwAAAACKHoXtdBWLSAiDyQKL0YPyAQ+v0cHqCDLTiBaF/3QCih4PthBGD7bLi1wkFDPRiwjB6QiLFJOLWAgz0Q+2ygNIBGnJBYQICIkQD7bTwesIQYlIBA+2SAcz0YtMJBQzHJGLTCQcSYlYCIlMJByFyX+QX15dW8IQAMzMzMzMzMzMzFaLdCQIV2oMi0Yo/9CL+FeJN+gJCQAAi04oahT/0cdAEAAAAADHQAwAAAAAiUcIi8dfXsIQAMzMzMzMzMzMzMxWi3QkCFeLPv92CItHMP/Q/3YEi0cw/9CLRzBW/9BfXsIQAMzMzMzMzMzMzMzMzMyLRCQMVot0JAxXi3wkDIXAdA1Q/3YE/zb/N+gh/f///3YcjUYM/3YY/3YU/3YQUI1GCFD/dgT/NlfoIgkAAF9e" & _
    "whAAzMzMzMzMzMzMzMzMzP90JATo5w8AAMIQAMzMzMxWi3QkCItODI1GDFeLPoXJdAw7TgR0B1BW6IIRAACLThCNRhCFyXQMO04IdAdQVuhsEQAAg34UAI1GFHQHUFboXBEAAI1GBFBW6FIRAACNRghQVuhIEQAAi0cwVv/QX17CEADMzMzMzMzMzMzMzMzMi0QkCFNVi2wkDFZXjVgMU414CFf/cAT/MFXoIQkAAItMJByL8IXJdBiF9nQUiweFwHQMUf8zUP91AOgx/P//i8ZfXl1bwhAAzMzMzMzMzMxWi3QkCIPK/4uGBIAEAA+3TMYCjQTGZjvKdAkPv8FmiRTG6xCLQASD+P90CGaJlEYIgAQAi4YEgAQAi0wkEIlMxgSLhgSABABmiVTGAo0UTouOBIAEAA+3ggiABABmiQTOi4YEgAQAD78Mxg+3wGaJggiABACD+f90DGaLhgSABABmiUTOAouOBIAEAIpEJAyIhA4AAAQAi4YEgAQAQCX/" & _
    "fwAAiYYEgAQAXsIMAMzMzIPsDFNVi2wkIFZXi3wkIDP2i38EiXwkGDm3DIAGAA+O4gAAAIuPDIAGAIvBK8YDxYP4BA+MnAAAAItUJCQ78X0KD7aEPgiABgDrCIvGK8EPtgQQjV4CiEQkFI1D/zvBfQoPtoQ+CYAGAOsJi8YrwQ+2RBABiEQkFTvZfQoPtoQ+CoAGAOsJi8YrwQ+2RBACiEQkFo1DATvBfQoPtoQ+C4AGAOsJi8YrwQ+2RBADiEQkF4tEJBRpyL2nNR7B6RBRUFfoiv7//0Y7twyABgAPjFH////rMYvOO7cMgAYAfSeNlwiABgDrCY2kJAAAAACL/4qEDwiABgCNUgGIQv9BO48MgAYAfOoptwyABgCF7Q+O9QEAAOsKjaQkAAAAAI1JAItcJCSD/QQPjFkBAACLhwAABAAz7YmHAIAEAGkDvac1HsHoEA+/lEcIgAQAg/r/D4QxAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/" & _
    "fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJCx9CQ+/FNeD+v91wIXtD47pAAAAi0QkKDlEJDB+BolEJDDrBItEJDC+BAAAADvGfnGNmwAAAACKBB4zyTPbiEQkE4XtfliQi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCSKBDiLfCQYOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkMEaL64tcJCQ78HyZi1wkJItsJCg79X09i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkKHzUi2wkKFb/txCABgD/dCQo6KAOAACF9n596ziLTCQgigM8j4sRi3EID7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFZS6EcSAAC+AQAAAItsJCiD/QR8FmkDvac1HsHoEFAPtgNQV+h1/P//6xWLjwyABgCKA4iEDwiABgD/hwyA" & _
    "BgBOQ02F9n/JiWwkKIlcJCSF7Q+PF/7//19eXVuDxAzCFADMzMzMzMzMzMzMzMzMg+wUU1WLbCQoVleLfCQoM/aLfwSJfCQgObcMgAYAD47iAAAAi48MgAYAi8ErxgPFg/gED4ycAAAAi1QkLDvxfQoPtoQ+CIAGAOsIi8YrwQ+2BBCNXgKIRCQYjUP/O8F9Cg+2hD4JgAYA6wmLxivBD7ZEEAGIRCQZO9l9Cg+2hD4KgAYA6wmLxivBD7ZEEAKIRCQajUMBO8F9Cg+2hD4LgAYA6wmLxivBD7ZEEAOIRCQbi0QkGGnIvac1HsHpEFFQV+hq+///Rju3DIAGAA+MUf///+sxi847twyABgB9J42XCIAGAOsJjaQkAAAAAIv/ioQPCIAGAI1SAYhC/0E7jwyABgB86im3DIAGADPAM8mJRCQUiUwkGIlEJByF7Q+OCQIAAOsKjaQkAAAAAI1JAItcJCyD/QQPjMABAACLhwAABAAz7YmHAIAEAGkDvac1" & _
    "HsHoEA+/lEcIgAQAg/r/D4SYAQAAi48EgAQAi8Irwb4AgAAAJf9/AAAr8CvOgeH/fwAAi4Q5AAAEADsDdQ6JtK8QgAYARTtsJDR9CQ+/FNeD+v91wIXtD45MAQAAi1QkMItEJDg7wn4Gi8KJRCQ4vgQAAAA7xn53igQeM8kz24hEJBOF7X5g6weNpCQAAAAAi5SPEIAGAIvGK8J5GIuHBIAEACvCA8Yl/38AAIqEOAAABADrC4t8JCyKBDiLfCQgOEQkE3UIiZSfEIAGAENBO818vYP7AX4Pi0QkOEaL64tcJCw78HyRi1wkLItUJDA78n05i5cQgAYAi84ryo0sGoXJeRiLhwSABAArwgPGJf9/AACKhDgAAAQA6wOKBBk4BCl1CEZBO3QkMHzUi2wkFIXtfl6NRQE78H5Pi0QkKIsQi2gIi0QkHDyPD7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFVS6CAPAAAPtgOLjxCABgCJTCQYiXQk" & _
    "FIlEJBzpkAAAAItEJBhVUOsli4cQgAYAiUQkGA+2A4l0JBSJRCQc63GLTCQYi2wkFIXtfjJVUf90JDDo8AoAADPAjXX/iUQkFIX2f1KLbCQwi0wkGIXtD48D/v//X15dW4PEFMIUAItMJCiKAzyPixGLcQgPtsiLQjx3CQ+2RAgwagjrDQ+2BAhqCY0ERQEAAABQVlLodg4AAL4BAAAAi2wkMIP9BHwWaQO9pzUewegQUA+2A1BX6KT4///rFYuPDIAGAIoDiIQPCIAGAP+HDIAGAE5DTYX2f8mJbCQwiVwkLOlq////zMzMzMzMzMzMVleLfCQMaLCPBgCLB4tAKP/Qi/CF9nUFX17CBACJdwSNTgK6AIAAAIPI/+sDjUkAx0EC/////41JCGaJQfhmiUH2g+oBdemNvgiABAC5AIAAAPOrX4mWBIAEAI1CAYmWDIAGAF7CBADMzMzMzMzMzFZXi3wkDLkSAAAAvgBQ51W4kCnnVfOli0wkEIPAYF9e" & _
    "xwEAEOdViUEEx0EIoEDnVcdBDHhE51XCCADMzMzMzMxTi1wkCFZXaACAAACLO4tzCItHKP/QiQYzwDlEJCRqAw+VwMdGCACAAACDwALHRgQAAAAAUFZX6DQNAACDfCQoAP90JDD/dCQw/3QkIP90JCBTdAfoF/j//+sF6DD7//9qB2oAVlfoBQ0AAIN8JCQAdBiLThCFyXQRuAgAAAArwVBqAFZX6OYMAACLRCQciw5fiQiLRCQci04EXluJCLgBAAAAwiQAzMzMzMzMzMzMzFFTVVaLdCQUV2gAgAAAix6JXCQci0Mo/9CLbCQgiYakgQAAx4asgQAAAIAAAMeGqIEAAAAAAACF7X8Ng76cAQAAAA+OtgUAAIO+nAEAABiNvpwBAAB9KotUJBzrA41JAIXtfhkPtgJNiw9C0+AJhpgBAACNQQiJB4P4GHzjiVQkHItGGIXAdQnHRhgBAAAA66aD+AF1c4sXg/oDD4xdBQAAi4aYAQAAg8L90eiLyIkX" & _
    "wegCiYaYAQAAg+EDdR2LysdGGAoAAACD4Qcr0dPoiReJhpgBAADpXP///4P5AXUYi0YEiUYMi0YIiUYQx0YYBgAAAOk/////g/kCD4U2////iU4Y6S7///+D+AJ1bosXg/oOD4zlBAAAi46YAQAAi8GD4B/B6QUFAQEAAMdGLAAAAACJRiCLwYPgH8HpBUDHRhgDAAAAiUYki8GD4A/B6QSDwASJjpgBAACJRiiNQvKJBzPAiUZEiUZIiUZMiUZQZolGVIhGVum7/v//g/gDdW85Bw+MdQQAAItGLDtGKH0xgz8DfCyLS0CLRiyKlpgBAACA4gcPtgQBiFQwRP9GLItGLIMH/cGumAEAAAM7Rih8z4tGLDtGKA+Faf7//2oTjUZEUFboxgkAAIlGFMdGGAQAAADHRiwAAAAA6Uf+//+D+AQPheIAAACLRiSLTiADwTlGLHw/UY1GV1BW6I8JAAD/diSJRgyLRiCDwFcDxlBW6HoJAACNfhSJRhBXVui9"
Private Const STR_THUNK2 As String = _
    "BQAAxwcAAAAAx0YYBgAAAOny/f///3YUjYaYAQAAV1DoGwYAAIvIg/n/D4SeAwAAg/n+D4RiAwAAg/kQfQ+LRiyITDBX/0Ys6br9//91B7gCAAAA6w8zwIP5EQ+VwI0EhQMAAACJRjAzwIP5Eg+UwI0ExQMAAACJRjSD+RB1G4tGLIXAfhQPtkQwVolGPMdGGAUAAADpbf3//zPAx0YYBQAAAIlGPOlc/f//g/gFdVyLH4tOMDvZD4wRAwAAi4aYAQAAugEAAADT4ivZSokfI9DT6ANWNImGmAEAAIXSfhyLRiQDRiCLTiw7yH0PikY8SohEMVf/RiyF0n/ki1wkGMdGGAQAAADp+/z//4P4Bg+FnAAAAP92DI1eDFeNhpgBAABQ6BgFAACD+P8PhJ0CAACD+P4PhF0CAAA9AAEAAH0QUFboCAQAAItcJBjptvz//3VAiwPHRhgBAAAAO0YEdA1TVuhYBAAAxwMAAAAAi0YQjX4Qi1wkGDtGCA+Ehvz/" & _
    "/1dW6DgEAADHBwAAAADpdPz//4tcJBg9HgEAAA+NZfz//8dGGAcAAACJRhzpVvz//4P4B3VUi0YcLQEBAACNDECLQzSNHIgPv0sCOQ8PjPsBAACLlpgBAAC4AQAAANPgSCPCA0MEiUY4D79DAikHiksCi1wkGNPqiZaYAQAAx0YYCAAAAOn9+///g/gIdTr/dhCNhpgBAABXUOghBAAAg/j/D4SmAQAAg/j+D4RqAQAAg/geD41hAQAAx0YYCQAAAIlGHOm++///g/gJD4WPAAAAi0YcixeNDECLQziNBIgPv0gCiUQkIIlMJBA70Q+MWgEAAIuGmAEAALsBAAAAK1QkENPji0wkIEsj2ANZBIkXikkC0+iDfjgAiYaYAQAAx0YYBgAAAHQsjaQkAAAAAIuGoIEAAP9OOCvDJf9/AAAPtoQwoAEAAFBW6IECAACDfjgAddv/TjiLXCQY6Sb7//+D+Ap1MYsXg/oQD4zdAAAAi46YAQAAD7fBiUZAjULw" & _
    "wekQiQeJjpgBAADHRhgLAAAA6fD6//+D+At1P4sHg/gQD4ynAAAAi46YAQAAg8DwD7fRiQeB8v//AACLRkDB6RCJjpgBAAA7wnVQ99gbwIPgC0CJRhjprPr//4P4DA+Fo/r//4M/CHxlD7aGmAEAAFBW6NkBAACDB/jBrpgBAAAIg0ZA/w+FfPr//8dGGAEAAADpcPr//4tcJBj/tqSBAACLQzD/0ItEJCTHhqSBAAAAAAAAX17HAAAAAACLRCQgXVvHAAAAAAAzwFnCFACLTCQkuAEAAACLlqSBAABfiRGLlqiBAACLTCQkXl1biRFZwhQAzMzMzMzMzMzMzMzMzIHsIAEAAFNWi7QkLAEAAFdosIEAAItGKP/Qi9iNfCQMuAgICAi5JAAAAGggAQAAiTPzq7gJCQkJjbwkoAAAALkcAAAA86uNRCQQx4QkEAEAAAcHBwdQU8eEJBwBAAAHBwcHx4QkIAEAAAcHBwfHhCQkAQAABwcHB8eEJCgBAAAH" & _
    "BwcHx4QkLAEAAAcHBwfHhCQwAQAACAgICMeEJDQBAAAICAgI6McEAACJQwSNRCQMaiBQU8dEJBgFBQUFx0QkHAUFBQXHRCQgBQUFBcdEJCQFBQUFx0QkKAUFBQXHRCQsBQUFBcdEJDAFBQUFx0QkNAUFBQXodwQAAF+JQwiLw17HQxgAAAAAx0MUAAAAAMdDEAAAAADHQwwAAAAAx4OYAQAAAAAAAMeDnAEAAAAAAADHg6CBAAAAAAAAW4HEIAEAAMIEAMzMzMzMzMzMzMzMzFOLXCQMVot0JAyLhqCBAACLDoicMKABAACLhqCBAABAJf9/AACJhqCBAACLhqiBAAA7hqyBAAB8GgPAUP+2pIEAAImGrIEAAItBLP/QiYakgQAAi4aogQAAi46kgQAAiBwB/4aogQAAXlvCCADMzMxTi1wkCFWLbCQQiwOJRCQMhe11CF2DyP9bwggAV4t9AIX/dEtWM/Y5N3wj6wONSQCLRwSNBPCDwASDOAB0B1BT" & _
    "6Lv///9GOzd+5otEJBT/dwSLQDD/0ItEJBTHRwQAAAAAV4tAMP/Qx0UAAAAAAF5fXTPAW8IIAMzMzMzMzMzMzItEJAxTi1wkDFaLCItABIszV4t8JBCLFyPKjQTID7YIO85/Jg+2CNPqK/Fmg3gC/3Usi0AEhcB0GosIi0AEI8qNBMgPtgg7zn7aX16DyP9bwgwAX164/v///1vCDACJF4kzD79AAl9eW8IMAMzMzMyD7AiLRCQMVYsoi0AIiUQkEItEJBiJbCQEhcAPji0BAABTVlc9BAEAAH4HvgIBAADrDIvwPQIBAAB+A41w/SvGux0AAACJRCQkg8//i0U0i+iJRCQUjQQ7mSvC0fiNDEA7dI0EfQSL2OvrO3SNCH4Ei/jr4YtcJByNBECNPIUAAAAAuRcBAAAPtwQvA/2LbCQQZjvBD7/Ii0U8fwwPtoRIAP7//2oH6wcPtkQBqGoIUFNV6CkDAAAPt0cCZoXAdA0rdwSYUFZTVegTAwAAi0U4" & _
    "g87/i1wkIL8eAAAAi+iL/40EN5krwtH4jQxAO1yNBH0Ei/jr6ztcjQh+BIvw6+GLfCQcjQRAjTSFAAAAAAP1i2wkEGoFD78Oi0U8D7YEyFBXVei5AgAAD7dGAmaFwHQPmFCLwytGBFBXVeihAgAAi0QkJIXAD4/Z/v//X15bXYPECMIMAMzMzMzMzMzMzMzMU1WLbCQMVldqCItFKP/Qi1wkKIv4i8u4AQAAANPgvgEAAACLTCQsSNPmiUQkKI0E9QAAAABQi0Uo/9CJRwQzyY1G/4PK/4kHhcB4I41kJACLRwRmiVTIAotHBMYEyACLRwTHRMgEAAAAAEE7D37hM+05bCQgfmmLdCQc6wONSQAPtgQuO8N+UYtEJBiLBKiLyCNMJCg7TCQkdT6Ly9P4Iwc7B380i08EZolswQKLTwQPthQuK9ONNMEPtg47yn0CiBaLdCQcugEAAAAPtgwuK8vT4gPCOwd+zEU7bCQgfKAz9jk3fGeLVCQsi2wkHIv/" & _
    "i0cEjQzwD7YBO8J+S4PI/2aJQQKLRwSNDPAPtgErwoP4B34FuAcAAABQjQQTiBFQi8uLxtPgC0QkLFD/dCQsVf90JCz/dCQs6MP+//+LTwSLVCQsiUTxBEY7N36ji8dfXl1bwhwAzMzMzMzMzMzMzIHsBAUAADPAM9JVi6wkFAUAAIlUJASJRCQMiUQkEIlEJBSJRCQYiUQkHIlEJCCJRCQkiUQkKIlEJCyJRCQwiUQkNIlEJDiJRCQ8iUQkQIlEJERWi7QkFAUAAIXtfhmL/w+2DDD/RIwMO9F9AovRQDvFfO2JVCQIM9IzyesDjUkAiVQMUANUDBCDwQQD0oP5PHzuVzP/he1+SVONnCSUAAAAD7Y0N8cDAAAAAItUtFSNQgGJRLRUhfZ0GDPAjWQkAIvKA8CD4QHR+gvBg+4BdfCJA4u0JBwFAABHg8MEO/18wFuLRCQMX4P4CXwFuAkAAABQagBqAFX/tCQkBQAAjYQkoAAAAFCLhCQoBQAA/zDo" & _
    "jP3//15dgcQEBQAAwgwAzItEJAxWi3QkDItOENPgCUYMi0QkFAPBiUYQg/gIfDlXi3wkDItGCDlGBHwPA8BQ/zaJRgiLRyz/0IkGi1YEiw6KRgyIBAr/RgTBbgwIg0YQ+IN+EAh9zV9ewhAAVYvsiwCAQMAgoGDgEJBQ0DCwcPAIiEjIKKho6BiYWNg4uHj4BIRExCSkZOQUlFTUNLR09AyMTMwsrGzsHJxc3Dy8fPwCgkLCIqJi4hKSUtIysnLyCopKyiqqauoamlraOrp6+gaGRsYmpmbmFpZW1ja2dvYOjk7OLq5u7h6eXt4+vn7+AYFBwSGhYeERkVHRMbFx8QmJSckpqWnpGZlZ2Tm5efkFhUXFJaVl5RWVVdU1tXX1DY1NzS2tbe0dnV3dPb19/QODQ8Mjo2PjE5NT0zOzc/MLi0vLK6tr6xubW9s7u3v7B4dHxyenZ+cXl1fXN7d39w+PT88vr2/vH59f3z+/f/8BAQAAAwAAAAMAAAACAQAA" & _
    "BAAAAAQAAAADAQAABQAAAAUAAAAEAQAABgAAAAYAAAAFAQAABwAAAAcAAAAGAQAACAAAAAgAAAAHAQAACQAAAAkAAAAIAQAACgAAAAoAAAAJAQEACwAAAAwAAAAKAQEADQAAAA4AAAALAQEADwAAABAAAAAMAQEAEQAAABIAAAANAQIAEwAAABYAAAAOAQIAFwAAABoAAAAPAQIAGwAAAB4AAAAQAQIAHwAAACIAAAARAQMAIwAAACoAAAASAQMAKwAAADIAAAATAQMAMwAAADoAAAAUAQMAOwAAAEIAAAAVAQQAQwAAAFIAAAAWAQQAUwAAAGIAAAAXAQQAYwAAAHIAAAAYAQQAcwAAAIIAAAAZAQUAgwAAAKIAAAAaAQUAowAAAMIAAAAbAQUAwwAAAOIAAAAcAQUA4wAAAAEBAAAdAQAAAgEAAAIBAAAQERIACAcJBgoFCwQMAw0CDgEPAAAAAAABAAAAAQAAAAEAAAACAAAAAgAAAAIAAAADAAAA" & _
    "AwAAAAMAAAAEAAAABAAAAAQAAQAFAAAABgAAAAUAAQAHAAAACAAAAAYAAgAJAAAADAAAAAcAAgANAAAAEAAAAAgAAwARAAAAGAAAAAkAAwAZAAAAIAAAAAoABAAhAAAAMAAAAAsABAAxAAAAQAAAAAwABQBBAAAAYAAAAA0ABQBhAAAAgAAAAA4ABgCBAAAAwAAAAA8ABgDBAAAAAAEAABAABwABAQAAgAEAABEABwCBAQAAAAIAABIACAABAgAAAAMAABMACAABAwAAAAQAABQACQABBAAAAAYAABUACQABBgAAAAgAABYACgABCAAAAAwAABcACgABDAAAABAAABgACwABEAAAABgAABkACwABGAAAACAAABoADAABIAAAADAAABsADAABMAAAAEAAABwADQABQAAAAGAAAB0ADQABYAAAAIAAAA=="
Private Const STR_THUNK_OFFSETS As String = "592|656|704|784|800|912|0|288|368|416|0|0|0|6896|7264|6640|7244"
Private Const STR_THUNK_BUILDDATE As String = "12.1.2018 17:15:52"
' end of generated code

'=========================================================================
' API
'=========================================================================

Private Const VT_BYREF                      As Long = &H4000
'--- for FindFirstFile
Private Const MAX_PATH                      As Long = 260
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for CreateFile
Private Const CREATE_ALWAYS                 As Long = 2
Private Const OPEN_EXISTING                 As Long = &H3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000
'--- for SetFilePointer
Private Const FILE_BEGIN                    As Long = 0
Private Const FILE_CURRENT                  As Long = 1
Private Const FILE_END                      As Long = 2
'--- for MultiByteToWideChar
Private Const CP_UTF8                       As Long = 65001
Private Const CP_OEMCP                      As Long = 1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for CryptBinaryToString/CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for CNG
Private Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
Private Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"
Private Const BCRYPT_AES_ALGORITHM          As String = "AES"
Private Const BCRYPT_HASH_LENGTH            As String = "HashDigestLength"
Private Const BCRYPT_OBJECT_LENGTH          As String = "ObjectLength"
Private Const BCRYPT_CHAINING_MODE          As String = "ChainingMode"
Private Const BCRYPT_CHAIN_MODE_ECB         As String = "ChainingModeECB"
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
'--- for GetStdHandle
Private Const STD_INPUT_HANDLE              As Long = -10
Private Const STD_OUTPUT_HANDLE             As Long = -11
Private Const STD_ERROR_HANDLE              As Long = -12

Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, lpNewFilePointer As Any, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime As SYSTEMTIME, pvTime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32" (ByVal vTime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If ImplCompress Then
    Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
#End If ' ImplCompress
#If ImplDecompress Then
    Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
#End If ' ImplDecompress
#If ImplCrypto Then
    Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
#End If ' ImplCrypto
#If ImplUseShared = 0 Then
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageW" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Args As Any) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
    Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
#End If ' Not ImplUseShared

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_SIG_LOCAL         As Long = &H4034B50
Private Const LNG_SIG_CENTRAL       As Long = &H2014B50
Private Const LNG_SIG_END           As Long = &H6054B50
Private Const LNG_SIG_DATADESC      As Long = &H7064B50
Private Const LNG_SIG_EXTLOCAL      As Long = &H8074B50
Private Const LNG_SIG_EOCDR         As Long = &H6064B50
Private Const LNG_SIG_LOCATOR       As Long = &H7064B50
Private Const LNG_LEN_LOCAL         As Long = 30
Private Const LNG_LEN_CENTRAL       As Long = 46
Private Const LNG_LEN_END           As Long = 22
Private Const LNG_LEN_EXTLOCAL      As Long = 16
Private Const LNG_LEN_EOCDR         As Long = 56
Private Const LNG_LEN_LOCATOR       As Long = 20
Private Const LNG_VER_DEFAULT       As Long = 10
Private Const LNG_VER_DEFLATE       As Long = 20
Private Const LNG_VER_ZIP64         As Long = 45
Private Const LNG_METHOD_DEFLATE    As Long = 8
Private Const LNG_METHOD_STORE      As Long = 0
Private Const LNG_IOBUF_SIZE        As Long = 65536
Private Const LNG_METHOD_AES        As Long = 99
Private Const LNG_EXTRADATA_AES_HEADER As Integer = &H9901
Private Const LNG_EXTRADATA_AES_VENDOR As Integer = &H4541 '--- "AE"
Private Const LNG_EXTRADATA_AES_SIZE As Long = 7
Private Const LNG_EXTRADATA_AE_2_MAXSIZE As Long = 20
Private Const LNG_EXTRADATA_ZIP64_HEADER As Long = 1
Private Const LNG_EXTRADATA_ZIP64_LOCAL_SIZE As Long = 16
Private Const LNG_ENC_HEADER_SIZE   As Long = 12
Private Const LNG_AES_AUTHCODE_SIZE As Long = 10
Private Const LNG_AES_BLOCK_SIZE    As Long = 16
Private Const LNG_PIPE_ENDED_ERROR  As Long = 109
Private Const LNG_SIZE_LIMIT        As Currency = 2 ^ 32 - 2
Private Const LNG_ENTRIES_LIMIT     As Long = 2 ^ 16 - 2
'--- errors
Private Const ERR_USER_CANCEL       As String = "User cancelled"
Private Const ERR_INIT_COMPRESSOR   As String = "Cannot init deflate compressor"
Private Const ERR_COMPRESSING       As String = "Error compressing"
Private Const ERR_EMPTY_ARCHIVE     As String = "Empty archive"
Private Const ERR_INVALID_ARCHIVE   As String = "Invalid archive"
Private Const ERR_ARCHIVE_NOT_OPEN  As String = "Archive not opened"
Private Const ERR_INVALID_LOCALHDR  As String = "Invalid local header"
Private Const ERR_UNSUPPORTED_METHOD As String = "Unsupported compression method %1"
Private Const ERR_INIT_DECOMPRESSOR As String = "Cannot init inflate decompressor"
Private Const ERR_READING_ARCHIVE   As String = "Error reading archive"
Private Const ERR_DECOMPRESSING     As String = "Error decompressing"
Private Const ERR_CRC_CHECK         As String = "CRC check failed"
Private Const ERR_ENTRY_INVALID_SIG As String = "Entry %1 has invalid signature"
Private Const ERR_ENTRY_NO_FILENAME As String = "Entry %1 has no filename"
Private Const ERR_UNSUPPORTED_ENCRYPTION As String = "Unsupported encryption"
Private Const ERR_PASSWORD_REQUIRED As String = "Password required"
Private Const ERR_INVALID_PASSWORD  As String = "Invalid password"
Private Const ERR_INVALID_AUTHCODE  As String = "Invalid authentication code"
Private Const ERR_INVALID_BYTEARRAY As String = "Invalid byte array"
Private Const ERR_MISSING_ZIP64     As String = "Missing Zip64 extension above 4GB limit"
Private Const ERR_NO_CRYPTO_COMPILED As String = "Crypto not compiled (ZIP_CRYPTO = 1)"
Private Const STR_BUFFER            As String = "[buffer]"
Private Const STR_STREAM            As String = "[stream]"
Private Const STR_ENTRY             As String = "[entry %1]"

Private m_uRtbl                 As ZipRelocTableType
Private m_sLastError            As String
Private m_bCancel               As Boolean
Private m_lFileCount            As Long
Private m_uFiles()              As ZipFileInfo
Private m_lCurrentFile          As Long
Private m_lCodePage             As Long
#If ImplDecompress Then
    Private m_sComment          As String
    Private m_vArchiveFile      As Variant
#End If ' ImplDecompress
#If ImplCrypto Then
    Private m_uCrypto           As ZipCryptoType
#End If ' ImplCrypto

Private Type ZipVfsType
    Handle              As Long
    Stream              As Object
    BufferArray         As Variant
    BufferBase          As LongPtr
    bufferSize          As Long
    BufferPtr           As LongPtr
    Data                As WIN32_FIND_DATA
    FileName            As String
End Type

Private Type ZipRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MemNonce            As Long
    MemXor              As Long
    ZipCrypt            As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Table          As Long
End Type

Private Type ZipBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
End Type

Private Enum ZipRelocationIndexesEnum
    zipRelCompressInit = 0
    zipRelCompressCleanup
    zipRelCompressBlock
    zipRelDecompressInit
    zipRelDecompressCleanup
    zipRelDecompressBlock
    zipRelCalcCrc32
    zipRelMemNonce
    zipRelMemXor
    zipRelZipCrypt
    zipRelMallocImpl
    zipRelReallocImpl
    zipRelFreeImpl
    zipRelLenCodes
    zipRelDistCodes
    zipRelMirrorBytes
    zipRelLenLenMap
End Enum

Private Enum ZipHeaderFlagsEnum
    zipFlagEncrypted = 2 ^ 0                         ' If set, indicates that the file is encrypted
    zipFlagCompressOptionMask = 2 ^ 1 Or 2 ^ 2       ' 0 - normal, 1 - maximum, 2 - fast, 3 - super fast
    zipFlagHasDataDescriptor = 2 ^ 3                 ' Values are put in the data descriptor immediately following the compressed data
    zipFlagStrongEncrypted = 2 ^ 6                   ' Strong encryption
    zipFlagUseUtf8 = 2 ^ 11                          ' Language encoding flag (EFS)
End Enum

Private Type Zip64ExtraDataType
    HeaderId            As Integer
    DataSize            As Integer
    USize               As Currency
    cSize               As Currency
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type Zip64EndHeaderType
    Signature           As Long
    Size                As Currency
    VerMade             As Integer
    VerExt              As Integer
    DiskNum             As Long
    DiskStart           As Long
    Entries             As Currency
    TotEntr             As Currency
    CenSize             As Currency
    CenOff              As Currency
End Type

Private Type Zip64EndLocatorType
    Signature           As Long
    DiskEocdr           As Long
    Offset              As Currency
    DiskNum             As Long
End Type

Private Type ZipLocalHeaderType
    Signature           As Long                     ' Signature
    VerExt              As Integer                  ' version needed to extract
    Flags               As Integer                  ' encrypt and compression flags
    Method              As Integer                  ' compression method
    FTime               As Integer                  ' time last modifies, dos format
    FDate               As Integer                  ' date last modifies, dos format
    '--- padding
    Crc32               As Long                     ' CRC32 for uncompressed file
    cSize               As Long                     ' compressed size
    USize               As Long                     ' uncompressed size
    LenFname            As Integer                  ' Length filename
    LenExt              As Integer                  ' Length for extra field
    Ext                 As Zip64ExtraDataType
End Type

Private Type ZipCentralHeaderType
    Signature           As Long                     ' Signature
    VerMade             As Integer                  ' version made by
    VerExt              As Integer                  ' version needed to extract
    Flags               As Integer                  ' encrypt and compression flags
    Method              As Integer                  ' compression method
    FTime               As Integer                  ' time last modifies, dos format
    FDate               As Integer                  ' date last modifies, dos format
    '--- padding
    Crc32               As Long                     ' CRC32 for uncompressed file
    cSize               As Long                     ' compressed size
    USize               As Long                     ' uncompressed size
    LenFname            As Integer                  ' Length filename
    LenExt              As Integer                  ' Length for extra field
    LenCom              As Integer                  ' Length for comment field
    DiskStart           As Integer                  ' start disk number
    AttribI             As Integer                  ' internal file attributes
    '--- padding
    AttribX             As Long                     ' external file attributes
    Offset              As Long                     ' relative offset of local header
    Ext                 As Zip64ExtraDataType
End Type

Private Type ZipEndHeaderType
    Signature           As Long                     ' Signature
    DiskNum             As Integer                  ' This disk number
    DiskStart           As Integer                  ' Start disk number
    Entries             As Integer                  ' Entries on this disk
    TotEntr             As Integer                  ' Number of total entries
    CenSize             As Long                     ' Size of entire cetral directory
    CenOff              As Long                     ' Offset of central on starting disk
    LenCom              As Integer                  ' Length of comment field
    Ext                 As Zip64EndHeaderType
End Type

Private Type ZipAesExtraDataType
    HeaderId            As Integer                  ' Extra field header ID (0x9901)
    DataSize            As Integer                  ' Data size (currently 7, but subject to possible increase in the future)
    Version             As Integer                  ' Integer version number specific to the zip vendor
    VendorId            As Integer                  ' 2-character vendor ID
    strength            As Byte                     ' Integer mode value indicating AES encryption strength
    '--- padding
    Method              As Integer                  ' The actual compression method used to compress the file
End Type

Private Type ZipFileInfo
    FileName            As String
    Attributes          As Long
    Crc32               As Long
    Size                As Currency
    compressedSize      As Currency
    Comment             As String
    LastModified        As Date
    Method              As Long
    Offset              As Currency
    Flags               As Long
    Extra()             As Byte
    SourceFile          As Variant
    Level               As Long
#If ImplCrypto Then
    Aes                 As ZipAesExtraDataType
    DecDat              As Byte
    Password            As String
#End If ' ImplCrypto
End Type

Private Type ZipCryptoType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HmacLen             As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 1)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    TradKey(0 To 3)     As Long
End Type

'=========================================================================
' Properties
'=========================================================================

Public Property Get SemVersion() As String
    SemVersion = "0.3.0"
End Property

Public Property Get ThunkBuildDate() As String
    ThunkBuildDate = STR_THUNK_BUILDDATE
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property

Public Property Get FileInfo(FileIdx As Variant, Optional ByVal InfoIdx As ZipFileInfoIndexEnum = -1) As Variant
    Dim sFileName       As String
    Dim lIdx            As Long
    
    If VarType(FileIdx) = vbString Then
        sFileName = LCase$(FileIdx)
        For lIdx = 0 To m_lFileCount - 1
            If LCase$(m_uFiles(lIdx).FileName) = sFileName Then
                Exit For
            End If
        Next
    Else
        lIdx = FileIdx
    End If
    If lIdx >= 0 And lIdx < m_lFileCount Then
        With m_uFiles(lIdx)
            Select Case InfoIdx
            Case Is < 0
                FileInfo = Array(.FileName, .Attributes, .Crc32, .Size, .compressedSize, .Comment, .LastModified, .Method, .Offset, .Flags)
            Case zipIdxFileName
                FileInfo = .FileName
            Case zipIdxAttributes
                FileInfo = .Attributes
            Case zipIdxCrc32
                FileInfo = .Crc32
            Case zipIdxSize
                FileInfo = .Size
            Case zipIdxCompressedSize
                FileInfo = .compressedSize
            Case zipIdxComment
                FileInfo = .Comment
            Case zipIdxLastModified
                FileInfo = .LastModified
            Case zipIdxMethod
                FileInfo = .Method
            Case zipIdxOffset
                FileInfo = .Offset
            Case zipIdxFlags
                FileInfo = .Flags
            End Select
        End With
    End If
End Property

Public Property Get CodePage() As Long
    CodePage = m_lCodePage
End Property

Public Property Let CodePage(ByVal lValue As Long)
    m_lCodePage = lValue
End Property

#If ImplDecompress Then

Public Property Get Comment() As String
    Comment = m_sComment
End Property
    
#End If ' ImplDecompress

'=========================================================================
' Methods
'=========================================================================

#If ImplCompress Then

Public Function AddFile( _
            File As Variant, _
            Optional Name As String, _
            Optional Comment As String, _
            Optional Password As String, _
            Optional EncrStrength As Long, _
            Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFile"
    Dim uFile           As ZipVfsType
    
    On Error GoTo EH
    pvSetError
    If m_lFileCount = 0 Then
        ReDim m_uFiles(0 To 2) As ZipFileInfo
    ElseIf m_lFileCount > UBound(m_uFiles) Then
        ReDim Preserve m_uFiles(0 To 2 * UBound(m_uFiles)) As ZipFileInfo
    End If
    uFile = pvVfsOpen(File)
    With m_uFiles(m_lFileCount)
        If LenB(Name) <> 0 Then
            .FileName = Name
        Else
            .FileName = Mid$(uFile.FileName, InStrRev(uFile.FileName, "\") + 1)
        End If
        If Right$(.FileName, 1) = "\" Then
            .Size = 0
            .Attributes = vbDirectory
        Else
            .Size = pvToInt64(uFile.Data.nFileSizeLow, uFile.Data.nFileSizeHigh)
            .Attributes = uFile.Data.dwFileAttributes
        End If
        .Comment = Comment
        .LastModified = pvFromFileTime(uFile.Data.ftLastWriteTime)
        .Extra = vbNullString
        If IsObject(File) Then
            '--- remove VT_BYREF if any
            Set .SourceFile = C_Obj(File)
        Else
            .SourceFile = File
        End If
        .Level = Level
        #If ImplCrypto Then
            .Password = Password
            .Aes.strength = EncrStrength
        #Else
            If LenB(Password) <> 0 Then
                On Error GoTo 0
                Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
            End If
        #End If ' ImplCrypto
    End With
    m_lFileCount = m_lFileCount + 1
    '--- success
    AddFile = True
QH:
    On Error Resume Next
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function CompressArchive( _
            ArchiveFile As Variant, _
            Optional Comment As String, _
            Optional ByVal Level As Long = 6, _
            Optional ByVal useUTF8 As VbTriState = vbUseDefault, _
            Optional ByVal UseZip64 As VbTriState = vbUseDefault) As Boolean
    Const FUNC_NAME     As String = "CompressArchive"
    Dim baBuffer()      As Byte
    Dim baNext()        As Byte
    Dim uArchiveFile    As ZipVfsType
    Dim uLocal          As ZipLocalHeaderType
    Dim lIdx            As Long
    Dim uFile           As ZipVfsType
    Dim lSize           As Long
    Dim cHeaderOffset   As Currency
    Dim cBeginOffset    As Currency
    Dim hCtx            As Long
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    Dim uCentral        As ZipCentralHeaderType
    Dim uEndHdr         As ZipEndHeaderType
    Dim lEntries        As Long
    Dim lComprLevel     As Long
    Dim bSkip           As Boolean
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
    Dim lJdx            As Long
    Dim uAes            As ZipAesExtraDataType
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    ReDim baNext(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsCreate(ArchiveFile)
    uCentral.VerExt = LNG_VER_DEFAULT
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            lComprLevel = IIf(.Level >= 0, .Level, Level)
            RaiseEvent BeforeCompress(lIdx, lComprLevel, bSkip, m_bCancel)
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            lComprLevel = Clamp(lComprLevel, 0, 9)
            uBuf.Greedy = (lComprLevel <= 4)
            uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), lComprLevel)
            uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), lComprLevel)
            .Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            On Error GoTo EH_Continue
            uFile = pvVfsOpen(.SourceFile)
            pvToDosDateTime .LastModified, uLocal.FDate, uLocal.FTime
            uLocal.Method = LNG_METHOD_DEFLATE
            uLocal.VerExt = LNG_VER_DEFLATE
            uLocal.Crc32 = -1
            uLocal.Ext.USize = 0
            uLocal.Ext.cSize = 0
            If useUTF8 = vbUseDefault Then
                uLocal.Flags = IIf(pvFromOemString(pvToOemString(.FileName, m_lCodePage), m_lCodePage) <> .FileName, zipFlagUseUtf8, 0)
            Else
                uLocal.Flags = IIf(useUTF8 = vbTrue, zipFlagUseUtf8, 0)
            End If
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    uLocal.Flags = uLocal.Flags Or zipFlagEncrypted
                    If .Aes.strength = 0 Then
                        uLocal.Flags = uLocal.Flags Or zipFlagHasDataDescriptor
                        If Not pvCryptoTradInit(ToUtf8Array(.Password)) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                        '--- encrypt 12-byte random header w/ last byte used for password check
                        pvCryptoGetRandom LNG_ENC_HEADER_SIZE, baSalt
                        Call CopyMemory(baSalt(LNG_ENC_HEADER_SIZE - 1), ByVal UnsignedAdd(VarPtr(uLocal.FDate), -1), 1)
                        If Not pvCryptoTradCrypt(baSalt) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                    Else
                        lSaltSize = (.Aes.strength + 1) * 4
                        pvCryptoGetRandom lSaltSize, baSalt
                        If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                            Err.Raise vbObjectError, , m_sLastError
                        End If
                        '--- reserve space in extra data for AE-1/2
                        lJdx = UBound(.Extra) + 1
                        ReDim Preserve .Extra(0 To lJdx + LNG_EXTRADATA_AES_SIZE + 3) As Byte
                    End If
                End If
            #End If ' ImplCrypto
            cHeaderOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    '--- prepend salt/enc_header before file data
                    pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), UBound(baSalt) + 1
                    If .Aes.strength > 0 Then
                        '--- prepend Password Verification Value for AES
                        pvVfsWrite uArchiveFile, VarPtr(nPassVer), 2
                    End If
                End If
            #End If ' ImplCrypto
            cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            If (.Attributes And vbDirectory + vbVolume) = 0 Then
                If lComprLevel > 0 Then
                    hCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
                    End If
                    lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                    Do
                        Call CopyMemory(baBuffer(0), baNext(0), lSize)
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lSize
                        lSize = pvVfsRead(uFile, VarPtr(baNext(0)), LNG_IOBUF_SIZE)
                        uBuf.Final = (lSize = 0)
                        lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf), VarPtr(uLocal.Crc32), lComprLevel) '--- level ignored
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_COMPRESSING
                        End If
                        #If ImplCrypto Then
                            If LenB(.Password) <> 0 Then
                                If UBound(baBuffer) < uBuf.OutLen Then
                                    ReDim baBuffer(0 To (uBuf.OutLen And -2048) + 2047) As Byte
                                End If
                                Call CopyMemory(baBuffer(0), ByVal uBuf.OutBlock, uBuf.OutLen)
                                If .Aes.strength = 0 Then
                                    If Not pvCryptoTradCrypt(baBuffer, Size:=uBuf.OutLen) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                Else
                                    If Not pvCryptoAesCrypt(baBuffer, Size:=uBuf.OutLen, HashAfter:=True) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                End If
                                pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), uBuf.OutLen
                                Call CoTaskMemFree(uBuf.OutBlock)
                                uBuf.OutBlock = 0
                            End If
                        #End If ' ImplCrypto
                        If uBuf.OutBlock <> 0 Then
                            pvVfsWrite uArchiveFile, uBuf.OutBlock, uBuf.OutLen
                            Call CoTaskMemFree(uBuf.OutBlock)
                            uBuf.OutBlock = 0
                        End If
                        uLocal.Ext.USize = uLocal.Ext.USize + uBuf.InLen
                        uLocal.Ext.cSize = uLocal.Ext.cSize + uBuf.OutLen
                        .compressedSize = uLocal.Ext.cSize
                        RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                    Loop While lSize <> 0
                    If hCtx <> 0 Then
                        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                        hCtx = 0
                    End If
                    If uLocal.Ext.cSize = 0 And uLocal.Ext.USize = 0 Then
                        uLocal.Method = LNG_METHOD_STORE
                        uLocal.VerExt = LNG_VER_DEFAULT
                    End If
                End If
                '--- if data is incompressible and not encrypted -> store file
                If lComprLevel = 0 Or (uLocal.Ext.cSize > uLocal.Ext.USize And (uLocal.Flags And zipFlagEncrypted) = 0) Then
                    pvVfsSeek uFile, 0, FILE_BEGIN
                    pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
                    uLocal.Method = LNG_METHOD_STORE
                    uLocal.VerExt = LNG_VER_DEFAULT
                    uLocal.Crc32 = -1
                    uLocal.Ext.USize = 0
                    uLocal.Ext.cSize = 0
                    #If ImplCrypto Then
                        If LenB(.Password) <> 0 Then
                            If .Aes.strength = 0 Then
                                pvVfsSeek uArchiveFile, cHeaderOffset, FILE_BEGIN
                                uLocal.Flags = uLocal.Flags And Not zipFlagEncrypted And Not zipFlagHasDataDescriptor
                                pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
                                .Password = vbNullString
                            Else
                                If Not pvCryptoAesInit(ToUtf8Array(.Password), baSalt, lSaltSize * 2, nPassVer) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            End If
                        End If
                    #End If ' ImplCrypto
                    Do
                        RaiseEvent Progress(lIdx, uLocal.Ext.USize, .Size, m_bCancel)
                        If m_bCancel Then
                            m_sLastError = ERR_USER_CANCEL
                            GoTo QH
                        End If
                        lSize = pvVfsRead(uFile, VarPtr(baBuffer(0)), LNG_IOBUF_SIZE)
                        If lSize = 0 Then
                            Exit Do
                        End If
                        Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lSize, VarPtr(uLocal.Crc32))
                        #If ImplCrypto Then
                            If LenB(.Password) <> 0 Then
                                If .Aes.strength = 0 Then
                                    If Not pvCryptoTradCrypt(baBuffer, Size:=lSize) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                Else
                                    If Not pvCryptoAesCrypt(baBuffer, Size:=lSize, HashAfter:=True) Then
                                        Err.Raise vbObjectError, , m_sLastError
                                    End If
                                End If
                            End If
                        #End If ' ImplCrypto
                        pvVfsWrite uArchiveFile, VarPtr(baBuffer(0)), lSize
                        uLocal.Ext.USize = uLocal.Ext.USize + lSize
                        uLocal.Ext.cSize = uLocal.Ext.cSize + lSize
                    Loop
                End If
            End If
            pvVfsClose uFile
            uLocal.Crc32 = uLocal.Crc32 Xor -1
            #If ImplCrypto Then
                If LenB(.Password) <> 0 Then
                    If .Aes.strength = 0 Then
                        uLocal.Ext.cSize = uLocal.Ext.cSize + LNG_ENC_HEADER_SIZE
                    Else
                        '--- append HMAC-SHA1 authentication tag after file data
                        baSalt = pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)
                        pvVfsWrite uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE
                        '--- update AE-1/2 in extra data
                        uAes.HeaderId = LNG_EXTRADATA_AES_HEADER
                        uAes.DataSize = LNG_EXTRADATA_AES_SIZE
                        uAes.Version = IIf(.Size <= LNG_EXTRADATA_AE_2_MAXSIZE, 2, 1)
                        uAes.VendorId = LNG_EXTRADATA_AES_VENDOR
                        uAes.strength = lSaltSize / 4 - 1
                        uAes.Method = uLocal.Method
                        lJdx = UBound(.Extra) - LNG_EXTRADATA_AES_SIZE - 3
                        Debug.Assert VarPtr(uAes.strength) + LenB(uAes.strength) - VarPtr(uAes.HeaderId) = 9
                        Call CopyMemory(.Extra(lJdx + 0), uAes.HeaderId, 9)
                        Call CopyMemory(.Extra(lJdx + 9), uAes.Method, LenB(uAes.Method))
                        '--- update local header
                        uLocal.Method = LNG_METHOD_AES
                        uLocal.Ext.cSize = uLocal.Ext.cSize + lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                        If uAes.Version = 2 Then
                            uLocal.Crc32 = 0
                        End If
                    End If
                End If
            #End If ' ImplCrypto
            .Crc32 = uLocal.Crc32
            Debug.Assert .Size = uLocal.Ext.USize
            .Size = uLocal.Ext.USize
            .compressedSize = uLocal.Ext.cSize
            .Method = uLocal.Method
            .Flags = uLocal.Flags
            If (uLocal.Flags And zipFlagHasDataDescriptor) <> 0 Then
                pvOutputDataDecriptor uArchiveFile, uLocal, UseZip64
                uLocal.Crc32 = 0
            End If
            cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            pvOutputLocalHeader uArchiveFile, uLocal, .FileName, .Extra, .Size, UseZip64
            pvVfsSeek uArchiveFile, cBeginOffset, FILE_BEGIN
            If uLocal.VerExt > uCentral.VerExt Then
                uCentral.VerExt = uLocal.VerExt
            End If
            If False Then
SkipFile:
                On Error GoTo EH
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                .Offset = -1
                If hCtx <> 0 Then
                    Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
                    hCtx = 0
                End If
            End If
            On Error GoTo EH
        End With
    Next
    m_lCurrentFile = -1
    cBeginOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
    uCentral.VerMade = LNG_VER_ZIP64
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            If .Offset >= 0 Then
                uCentral.Method = .Method
                pvToDosDateTime .LastModified, uCentral.FDate, uCentral.FTime
                uCentral.Crc32 = .Crc32
                uCentral.AttribX = .Attributes
                uCentral.Flags = .Flags
                uCentral.Ext.cSize = .compressedSize
                uCentral.Ext.USize = .Size
                uCentral.Ext.Offset = .Offset
                pvOutputCentralHeader uArchiveFile, uCentral, .FileName, .Extra, .Comment, UseZip64
                lEntries = lEntries + 1
            End If
        End With
    Next
    m_lCurrentFile = -1
    If lEntries = 0 Then
        Err.Raise vbObjectError, , ERR_EMPTY_ARCHIVE
    End If
    uEndHdr.Ext.Entries = lEntries
    uEndHdr.Ext.TotEntr = lEntries
    uEndHdr.Ext.CenSize = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT) - cBeginOffset
    uEndHdr.Ext.CenOff = cBeginOffset
    pvOutputEndHeader uArchiveFile, uEndHdr, Comment, UseZip64
    pvVfsSetEof uArchiveFile, "[output_archive]"
    '--- success
    CompressArchive = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uFile
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Public Function AddFromFolder( _
            sFolderAndMask As String, _
            Optional Recursive As Boolean, _
            Optional TargetFolder As String, _
            Optional IncludeEmptyFolders As Boolean, _
            Optional Password As String, _
            Optional EncrStrength As Long, _
            Optional Level As Long = -1) As Boolean
    Const FUNC_NAME     As String = "AddFromFolder"
    Dim lIdx            As Long
    Dim sFolder         As String
    Dim sMask           As String
    Dim lRootOffset     As Long
    Dim cFolders        As Collection
    Dim vElem           As Variant
    Dim sName           As String
    
    On Error GoTo EH
    pvSetError
    Set cFolders = New Collection
    lIdx = InStrRev(sFolderAndMask, "\")
    If lIdx > 0 Then
        sFolder = Left$(sFolderAndMask, lIdx - 1)
    End If
    sMask = Mid$(sFolderAndMask, lIdx + 1)
    If LenB(sFolder) = 0 Then
        sFolder = CurDir$()
    End If
    lRootOffset = Len(sFolder)
    cFolders.Add sFolder
    Do While cFolders.Count > 0
        sFolder = cFolders.Item(1)
        cFolders.Remove 1
        lIdx = 0
        For Each vElem In pvEnumFiles(sFolder, sMask, -vbDirectory)
            If AddFile(vElem, PathCombine(TargetFolder, Mid$(vElem, lRootOffset + 2)), Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                lIdx = lIdx + 1
                '--- success (entries added)
                AddFromFolder = True
            End If
        Next
        If lIdx = 0 And IncludeEmptyFolders Then
            sName = PathCombine(TargetFolder, Mid$(sFolder, lRootOffset + 2))
            If LenB(sName) <> 0 Then
                If AddFile(sFolder & "\", sName & "\", Password:=Password, EncrStrength:=EncrStrength, Level:=Level) Then
                    '--- success (entries added)
                    AddFromFolder = True
                End If
            End If
        End If
        If Recursive Then
            For Each vElem In pvEnumFiles(sFolder, "*.*", vbDirectory)
                cFolders.Add vElem
            Next
        End If
    Loop
QH:
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Private Function pvEnumFiles(sFolder As String, sMask As String, ByVal eAttrib As VbFileAttribute) As Collection
    Dim sFile           As String
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    
    On Error GoTo EH
    Set pvEnumFiles = New Collection
    sFile = PathCombine(sFolder, "*.*")
    hFind = FindFirstFile(StrPtr(sFile), VarPtr(uData))
    If hFind = INVALID_HANDLE_VALUE Then
        If Err.LastDllError <> 2 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "pvEnumFiles", GetSystemMessage(Err.LastDllError) & " (" & sFile & ")"
        End If
    Else
        Do
            If eAttrib > 0 And (uData.dwFileAttributes And eAttrib) <> 0 _
                    Or eAttrib < 0 And (uData.dwFileAttributes And -eAttrib) = 0 Then
                sFile = Left$(uData.cFileName, InStr(uData.cFileName, Chr$(0)) - 1)
                If sFile <> "." And sFile <> ".." Then
                    If PathMatchSpecW(StrPtr(sFile), StrPtr(sMask)) <> 0 Then
                        pvEnumFiles.Add PathCombine(sFolder, sFile)
                    End If
                End If
            End If
        Loop While FindNextFile(hFind, VarPtr(uData)) <> 0
        Call FindClose(hFind)
        hFind = 0
    End If
    Exit Function
EH:
    If hFind <> 0 And hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        hFind = 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub pvOutputLocalHeader(uFile As ZipVfsType, uHdr As ZipLocalHeaderType, FileName As String, baExtra() As Byte, ByVal cSize As Currency, ByVal UseZip64 As VbTriState)
    Dim baFName()       As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.cSize > LNG_SIZE_LIMIT Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or cSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.cSize = -1
        uHdr.USize = -1
    Else
        uHdr.cSize = pvToInt32(uHdr.Ext.cSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    baFName = pvToOemString(Replace(FileName, "\", "/"), IIf((uHdr.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage))
    uHdr.Signature = LNG_SIG_LOCAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + LNG_EXTRADATA_ZIP64_LOCAL_SIZE, 0) + UBound(baExtra) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.LenExt) + 2 - VarPtr(uHdr.Crc32) = LNG_LEN_LOCAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 14
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 14
    Debug.Assert VarPtr(uHdr.LenExt) + LenB(uHdr.LenExt) - VarPtr(uHdr.Crc32) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 16
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        uHdr.Ext.DataSize = LNG_EXTRADATA_ZIP64_LOCAL_SIZE
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.cSize = uHdr.Ext.cSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.cSize) + LenB(uHdr.Ext.cSize) - VarPtr(uHdr.Ext.HeaderId) = 20
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 20
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.cSize = uHdr.Ext.cSize * 10000@
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
End Sub

Private Sub pvOutputDataDecriptor(uFile As ZipVfsType, uHdr As ZipLocalHeaderType, ByVal UseZip64 As VbTriState)
    pvVfsWrite uFile, VarPtr(LNG_SIG_DATADESC), 4
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 4
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.cSize > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.USize = uHdr.Ext.USize / 10000@
        uHdr.Ext.cSize = uHdr.Ext.cSize / 10000@
        Debug.Assert VarPtr(uHdr.Ext.USize) + LenB(uHdr.Ext.USize) - VarPtr(uHdr.Ext.cSize) = 16
        pvVfsWrite uFile, VarPtr(uHdr.Ext.cSize), 16
        uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        uHdr.Ext.cSize = uHdr.Ext.cSize * 10000@
    Else
        uHdr.cSize = pvToInt32(uHdr.Ext.cSize)
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
        Debug.Assert VarPtr(uHdr.USize) + LenB(uHdr.USize) - VarPtr(uHdr.cSize) = 8
        pvVfsWrite uFile, VarPtr(uHdr.cSize), 8
        uHdr.cSize = 0
        uHdr.USize = 0
    End If
End Sub

Private Sub pvOutputCentralHeader(uFile As ZipVfsType, uHdr As ZipCentralHeaderType, FileName As String, baExtra() As Byte, Comment As String, ByVal UseZip64 As VbTriState)
    Dim lCodePage       As Long
    Dim baFName()       As Byte
    Dim baComment()     As Byte
    
    If UseZip64 = vbTrue Or uHdr.Ext.USize > LNG_SIZE_LIMIT Or uHdr.Ext.cSize > LNG_SIZE_LIMIT Or uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        UseZip64 = vbTrue
        uHdr.VerExt = LNG_VER_ZIP64
        uHdr.Ext.DataSize = 0
    End If
    If uHdr.Ext.cSize > LNG_SIZE_LIMIT Then
        uHdr.cSize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.cSize = pvToInt32(uHdr.Ext.cSize)
    End If
    If uHdr.Ext.USize > LNG_SIZE_LIMIT Then
        uHdr.USize = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.USize = pvToInt32(uHdr.Ext.USize)
    End If
    If uHdr.Ext.Offset > LNG_SIZE_LIMIT Then
        uHdr.Offset = -1
        uHdr.Ext.DataSize = uHdr.Ext.DataSize + 8
    Else
        uHdr.Offset = pvToInt32(uHdr.Ext.Offset)
    End If
    lCodePage = IIf((uHdr.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage)
    baFName = pvToOemString(Replace(FileName, "\", "/"), lCodePage)
    baComment = pvToOemString(Comment, lCodePage)
    uHdr.Signature = LNG_SIG_CENTRAL
    uHdr.LenFname = UBound(baFName) + 1
    uHdr.LenExt = IIf(UseZip64 = vbTrue, 4 + uHdr.Ext.DataSize, 0) + UBound(baExtra) + 1
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) + VarPtr(uHdr.AttribI) + 2 - VarPtr(uHdr.Crc32) + VarPtr(uHdr.Offset) + 4 - VarPtr(uHdr.AttribX) = LNG_LEN_CENTRAL
    Debug.Assert VarPtr(uHdr.FDate) + LenB(uHdr.FDate) - VarPtr(uHdr.Signature) = 16
    pvVfsWrite uFile, VarPtr(uHdr.Signature), 16
    Debug.Assert VarPtr(uHdr.AttribI) + LenB(uHdr.AttribI) - VarPtr(uHdr.Crc32) = 22
    pvVfsWrite uFile, VarPtr(uHdr.Crc32), 22
    Debug.Assert VarPtr(uHdr.Offset) + LenB(uHdr.Offset) - VarPtr(uHdr.AttribX) = 8
    pvVfsWrite uFile, VarPtr(uHdr.AttribX), 8
    Debug.Assert uHdr.LenFname > 0
    If uHdr.LenFname > 0 Then
        pvVfsWrite uFile, VarPtr(baFName(0)), uHdr.LenFname
    End If
    If UseZip64 = vbTrue Then
        uHdr.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER
        Debug.Assert VarPtr(uHdr.Ext.DataSize) + LenB(uHdr.Ext.DataSize) - VarPtr(uHdr.Ext.HeaderId) = 4
        pvVfsWrite uFile, VarPtr(uHdr.Ext.HeaderId), 4
        If uHdr.USize = -1 Then
            uHdr.Ext.USize = uHdr.Ext.USize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.USize), LenB(uHdr.Ext.USize)
            uHdr.Ext.USize = uHdr.Ext.USize * 10000@
        End If
        If uHdr.cSize = -1 Then
            uHdr.Ext.cSize = uHdr.Ext.cSize / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.cSize), LenB(uHdr.Ext.cSize)
            uHdr.Ext.cSize = uHdr.Ext.cSize * 10000@
        End If
        If uHdr.Offset = -1 Then
            uHdr.Ext.Offset = uHdr.Ext.Offset / 10000@
            pvVfsWrite uFile, VarPtr(uHdr.Ext.Offset), LenB(uHdr.Ext.Offset)
            uHdr.Ext.Offset = uHdr.Ext.Offset * 10000@
        End If
    End If
    If UBound(baExtra) >= 0 Then
        pvVfsWrite uFile, VarPtr(baExtra(0)), UBound(baExtra) + 1
    End If
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvOutputEndHeader(uFile As ZipVfsType, uHdr As ZipEndHeaderType, Comment As String, ByVal UseZip64 As VbTriState)
    Dim baComment()     As Byte
    Dim uLocator        As Zip64EndLocatorType
    
    
    If UseZip64 = vbTrue Or uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Or uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Or uHdr.Ext.CenSize > LNG_SIZE_LIMIT Or uHdr.Ext.CenOff > LNG_SIZE_LIMIT Then
        If UseZip64 = vbFalse Then
            Err.Raise vbObjectError, , ERR_MISSING_ZIP64
        End If
        uHdr.CenOff = -1
        '--- Zip64 End of Central Directory record
        uHdr.Ext.Signature = LNG_SIG_EOCDR
        uHdr.Ext.Size = (LNG_LEN_EOCDR - 12) / 10000@
        uHdr.Ext.VerMade = LNG_VER_ZIP64
        uHdr.Ext.VerExt = LNG_VER_ZIP64
        uHdr.Ext.Entries = uHdr.Ext.Entries / 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr / 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize / 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff / 10000@
        uLocator.Offset = pvVfsSeek(uFile, 0, FILE_CURRENT) / 10000@
        Debug.Assert VarPtr(uHdr.Ext.CenOff) + 8 - VarPtr(uHdr.Ext.Signature) = LNG_LEN_EOCDR
        pvVfsWrite uFile, VarPtr(uHdr.Ext.Signature), LNG_LEN_EOCDR
        uHdr.Ext.Entries = uHdr.Ext.Entries * 10000@
        uHdr.Ext.TotEntr = uHdr.Ext.TotEntr * 10000@
        uHdr.Ext.CenSize = uHdr.Ext.CenSize * 10000@
        uHdr.Ext.CenOff = uHdr.Ext.CenOff * 10000@
        '--- Zip64 End of Central Directory Locator record
        uLocator.Signature = LNG_SIG_LOCATOR
        Debug.Assert VarPtr(uLocator.DiskNum) + 4 - VarPtr(uLocator.Signature) = LNG_LEN_LOCATOR
        pvVfsWrite uFile, VarPtr(uLocator.Signature), LNG_LEN_LOCATOR
    Else
        uHdr.CenOff = pvToInt32(uHdr.Ext.CenOff)
    End If
    If uHdr.Ext.Entries > LNG_ENTRIES_LIMIT Then
        uHdr.Entries = -1
    Else
        uHdr.Entries = pvToInt32(uHdr.Ext.Entries) And &HFFFF&
    End If
    If uHdr.Ext.TotEntr > LNG_ENTRIES_LIMIT Then
        uHdr.TotEntr = -1
    Else
        uHdr.TotEntr = pvToInt32(uHdr.Ext.TotEntr) And &HFFFF&
    End If
    If uHdr.Ext.CenSize > LNG_SIZE_LIMIT Then
        uHdr.CenSize = -1
    Else
        uHdr.CenSize = pvToInt32(uHdr.Ext.CenSize)
    End If
    baComment = pvToOemString(Comment, m_lCodePage)
    uHdr.Signature = LNG_SIG_END
    uHdr.LenCom = UBound(baComment) + 1
    Debug.Assert VarPtr(uHdr.LenCom) + 2 - VarPtr(uHdr.Signature) = LNG_LEN_END
    pvVfsWrite uFile, VarPtr(uHdr.Signature), LNG_LEN_END
    If uHdr.LenCom > 0 Then
        pvVfsWrite uFile, VarPtr(baComment(0)), uHdr.LenCom
    End If
End Sub

Private Sub pvToDosDateTime(dDate As Date, nDate As Integer, nTime As Integer)
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME
    
    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call FileTimeToDosDateTime(uFileTime, VarPtr(nDate), VarPtr(nTime))
End Sub

Private Function pvToOemString(sText As String, ByVal lCodePage As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    If Len(sText) <> 0 Then
        lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
        If lSize > 0 Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            lSize = WideCharToMultiByte(lCodePage, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), UBound(baRetVal) + 1, 0, 0)
            pvToOemString = baRetVal
        Else
            pvToOemString = vbNullString
        End If
    Else
        pvToOemString = vbNullString
    End If
End Function

#End If ' ImplCompress

#If ImplDecompress Then

Public Function OpenArchive(ArchiveFile As Variant) As Boolean
    Const FUNC_NAME     As String = "OpenArchive"
    Const MAX_END_SEEK  As Long = 10000
    Dim uArchiveFile    As ZipVfsType
    Dim uEndHdr         As ZipEndHeaderType
    Dim lIdx            As Long
    Dim sArchiveComment As String
    Dim uFiles()        As ZipFileInfo
    Dim uCentral        As ZipCentralHeaderType
    Dim baComment()     As Byte
    Dim cOffset         As Currency
    Dim baBuffer()      As Byte
    Dim uLocator        As Zip64EndLocatorType
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(ArchiveFile)
    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_END)
    If cOffset >= LNG_LEN_END Then
        For lIdx = 0 To IIf(cOffset - LNG_LEN_END < MAX_END_SEEK, cOffset - LNG_LEN_END, MAX_END_SEEK)
            cOffset = pvVfsSeek(uArchiveFile, -LNG_LEN_END - lIdx, FILE_END)
            Debug.Assert VarPtr(uEndHdr.LenCom) + 2 - VarPtr(uEndHdr.Signature) = LNG_LEN_END
            pvVfsRead uArchiveFile, VarPtr(uEndHdr), LNG_LEN_END
            If uEndHdr.Signature = LNG_SIG_END And uEndHdr.LenCom = lIdx Then
                If uEndHdr.LenCom > 0 Then
                    ReDim baComment(0 To uEndHdr.LenCom - 1) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baComment(0)), uEndHdr.LenCom
                    sArchiveComment = pvFromOemString(baComment, m_lCodePage)
                End If
                If uEndHdr.CenOff = -1 Or uEndHdr.Ext.Entries = -1 Then
                    If cOffset <= LNG_LEN_LOCATOR + LNG_LEN_EOCDR Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, cOffset - LNG_LEN_LOCATOR, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uLocator), LNG_LEN_LOCATOR
                    If uLocator.Signature <> LNG_SIG_LOCATOR Or uLocator.DiskEocdr <> uEndHdr.DiskNum Then
                        Exit For
                    End If
                    pvVfsSeek uArchiveFile, uLocator.Offset * 10000@, FILE_BEGIN
                    pvVfsRead uArchiveFile, VarPtr(uEndHdr.Ext), LNG_LEN_EOCDR
                    If uEndHdr.Ext.Signature <> LNG_SIG_EOCDR Then
                        Exit For
                    End If
                    uEndHdr.Ext.CenOff = uEndHdr.Ext.CenOff * 10000@
                    uEndHdr.Ext.Entries = uEndHdr.Ext.Entries * 10000@
                Else
                    uEndHdr.Ext.CenOff = pvToInt64(uEndHdr.CenOff)
                    uEndHdr.Ext.Entries = uEndHdr.Entries And &HFFFF&
                End If
                lIdx = -1
                Exit For
            End If
        Next
    End If
    If lIdx < 0 Then
        '--- note: redim one more (last ignored)
        ReDim uFiles(0 To uEndHdr.Ext.Entries) As ZipFileInfo
        pvVfsSeek uArchiveFile, uEndHdr.Ext.CenOff, FILE_BEGIN
        For lIdx = 0 To uEndHdr.Ext.Entries - 1
            m_lCurrentFile = lIdx
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) + VarPtr(uCentral.AttribI) + 2 - VarPtr(uCentral.Crc32) + VarPtr(uCentral.Offset) + 4 - VarPtr(uCentral.AttribX) = LNG_LEN_CENTRAL
            Debug.Assert VarPtr(uCentral.FDate) + LenB(uCentral.FDate) - VarPtr(uCentral.Signature) = 16
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 16
            Debug.Assert VarPtr(uCentral.AttribI) + LenB(uCentral.AttribI) - VarPtr(uCentral.Crc32) = 22
            pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 22
            Debug.Assert VarPtr(uCentral.Offset) + LenB(uCentral.Offset) - VarPtr(uCentral.AttribX) = 8
            pvVfsRead uArchiveFile, VarPtr(uCentral.AttribX), 8
            If uCentral.Signature <> LNG_SIG_CENTRAL Then
                pvSetError MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
                GoTo QH
            End If
            If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(lIdx), lIdx + 1) Then
                GoTo QH
            End If
        Next
    Else
        '--- fall-back to sequential lookup from start of file
        pvVfsSeek uArchiveFile, 0, FILE_BEGIN
        uEndHdr.Ext.Entries = 0
        Do
            uCentral.Offset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
            pvVfsRead uArchiveFile, VarPtr(uCentral.Signature), 4
            Select Case uCentral.Signature
            Case LNG_SIG_LOCAL
                pvVfsRead uArchiveFile, VarPtr(uCentral.VerExt), 10
                pvVfsRead uArchiveFile, VarPtr(uCentral.Crc32), 16
                uEndHdr.Ext.Entries = uEndHdr.Ext.Entries + 1
                If uEndHdr.Ext.Entries = 1 Then
                    ReDim uFiles(0 To 3) As ZipFileInfo
                ElseIf uEndHdr.Ext.Entries > UBound(uFiles) Then
                    ReDim Preserve uFiles(0 To 2 * uEndHdr.Ext.Entries) As ZipFileInfo
                End If
                If Not pvLoadFileInfo(uArchiveFile, uCentral, uFiles(uEndHdr.Ext.Entries - 1), uEndHdr.Ext.Entries) Then
                    GoTo QH
                End If
                If uCentral.Ext.cSize = 0 And uCentral.Ext.USize <> 0 Then
                    cOffset = pvVfsSeek(uArchiveFile, 0, FILE_CURRENT)
                    ReDim baBuffer(0 To uCentral.Ext.USize - LNG_LEN_EXTLOCAL) As Byte
                    pvVfsRead uArchiveFile, VarPtr(baBuffer(0)), UBound(baBuffer) + 1
                    For lIdx = 0 To UBound(baBuffer) - LNG_LEN_EXTLOCAL
                        If Peek(VarPtr(baBuffer(lIdx))) = LNG_SIG_EXTLOCAL Then
                            If Peek(VarPtr(baBuffer(lIdx + 4))) = uCentral.Crc32 Or uCentral.Crc32 = 0 Then
                                uCentral.Crc32 = Peek(VarPtr(baBuffer(lIdx + 4)))
                                uCentral.Ext.cSize = Peek(VarPtr(baBuffer(lIdx + 8)))
                                uFiles(uEndHdr.Ext.Entries - 1).Crc32 = uCentral.Crc32
                                uFiles(uEndHdr.Ext.Entries - 1).compressedSize = uCentral.Ext.cSize
                                Exit For
                            End If
                        End If
                    Next
                    If uCentral.Ext.cSize = 0 Then
                        Exit Do
                    End If
                    pvVfsSeek uArchiveFile, cOffset, FILE_BEGIN
                End If
                pvVfsSeek uArchiveFile, uCentral.Ext.cSize, FILE_CURRENT
            Case LNG_SIG_EXTLOCAL
                If uEndHdr.Ext.Entries > 0 Then
                    pvVfsSeek uArchiveFile, LNG_LEN_EXTLOCAL - 4, FILE_CURRENT
                End If
            Case Else
                Exit Do
            End Select
        Loop
        '--- note: redim one more (last ignored)
        ReDim Preserve uFiles(0 To uEndHdr.Ext.Entries) As ZipFileInfo
    End If
    '--- commit archive info to member vars
    AssignVariant m_vArchiveFile, ArchiveFile
    m_lFileCount = uEndHdr.Ext.Entries
    m_sComment = sArchiveComment
    m_uFiles = uFiles
    '--- success
    OpenArchive = True
QH:
    pvVfsClose uArchiveFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Private Function pvLoadFileInfo(uArchiveFile As ZipVfsType, uCentral As ZipCentralHeaderType, uFile As ZipFileInfo, ByVal lEntry As Long) As Boolean
    Const FUNC_NAME     As String = "pvLoadFileInfo"
    Dim baFName()       As Byte
    Dim baExtra()       As Byte
    Dim baComment()     As Byte
    Dim lCodePage       As Long
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim aSizes(0 To 2)  As Currency
    
    With uFile
        .Method = uCentral.Method
        .LastModified = pvFromDosDateTime(uCentral.FDate, uCentral.FTime)
        .Crc32 = uCentral.Crc32
        .Attributes = uCentral.AttribX
        .Flags = uCentral.Flags
        lCodePage = IIf((.Flags And zipFlagUseUtf8) <> 0, CP_UTF8, m_lCodePage)
        Debug.Assert uCentral.LenFname > 0
        If uCentral.LenFname > 0 Then
            ReDim baFName(0 To uCentral.LenFname - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baFName(0)), uCentral.LenFname
            .FileName = pvToWinFileName(baFName, lCodePage)
        Else
            If pvSetError(MODULE_NAME & "." & FUNC_NAME, ERR_INVALID_ARCHIVE & ". " & Replace(ERR_ENTRY_NO_FILENAME, "%1", lEntry), CanContinue:=True) Then
                GoTo QH
            End If
            .FileName = Replace(STR_ENTRY, "%1", lEntry)
        End If
        uCentral.Ext.USize = uCentral.USize
        uCentral.Ext.cSize = uCentral.cSize
        uCentral.Ext.Offset = uCentral.Offset
        If uCentral.LenExt > 0 Then
            ReDim baExtra(0 To uCentral.LenExt - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baExtra(0)), uCentral.LenExt
            .Extra = baExtra
            lPos = 0
            Do While lPos + 3 < uCentral.LenExt
                Call CopyMemory(uCentral.Ext.HeaderId, baExtra(lPos + 0), 4)
                If uCentral.Ext.HeaderId = LNG_EXTRADATA_ZIP64_HEADER Then
                    If uCentral.Ext.DataSize >= 8 Then
                        Call CopyMemory(aSizes(0), baExtra(lPos + 4), Clamp(uCentral.Ext.DataSize, 0, 24))
                        If uCentral.USize = -1 Then
                            uCentral.Ext.USize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.cSize = -1 Then
                            uCentral.Ext.cSize = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                        If uCentral.Offset = -1 Then
                            uCentral.Ext.Offset = aSizes(lIdx) * 10000@
                            lIdx = lIdx + 1
                        End If
                    End If
                    Exit Do
                End If
                lPos = lPos + uCentral.Ext.DataSize + 4
            Loop
        End If
        .compressedSize = uCentral.Ext.cSize
        .Size = uCentral.Ext.USize
        .Offset = uCentral.Ext.Offset
        If uCentral.LenCom > 0 Then
            ReDim baComment(0 To uCentral.LenCom - 1) As Byte
            pvVfsRead uArchiveFile, VarPtr(baComment(0)), uCentral.LenCom
            .Comment = pvFromOemString(baComment, lCodePage)
        End If
        #If ImplCrypto Then
            If (.Flags And zipFlagEncrypted) <> 0 Then
                If .Method = LNG_METHOD_AES Then
                    lPos = 0
                    Do While lPos + 3 < UBound(baExtra) + 1
                        Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 4)
                        If .Aes.HeaderId = LNG_EXTRADATA_AES_HEADER Then
                            If .Aes.DataSize >= LNG_EXTRADATA_AES_SIZE Then
                                Debug.Assert VarPtr(.Aes.strength) + 1 - VarPtr(.Aes.HeaderId) = 9
                                Call CopyMemory(.Aes.HeaderId, baExtra(lPos + 0), 9)
                                Call CopyMemory(.Aes.Method, baExtra(lPos + 9), 2)
                                .Method = .Aes.Method
                            End If
                            Exit Do
                        End If
                        lPos = lPos + .Aes.DataSize + 4
                    Loop
                End If
                '--- store for traditional ZipCrypt password check
                If (.Flags And zipFlagHasDataDescriptor) <> 0 Then
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.FDate), -1), 1)
                Else
                    Call CopyMemory(.DecDat, ByVal UnsignedAdd(VarPtr(uCentral.cSize), -1), 1)
                End If
            End If
        #End If ' ImplCrypto
    End With
    '--- success
    pvLoadFileInfo = True
QH:
End Function

Public Function Extract( _
            OutputTarget As Variant, _
            Optional Filter As Variant, _
            Optional Password As String) As Boolean
    Const FUNC_NAME     As String = "Extract"
    Dim baBuffer()      As Byte
    Dim uArchiveFile    As ZipVfsType
    Dim lIdx            As Long
    Dim uLocal          As ZipLocalHeaderType
    Dim baFName()       As Byte
    Dim baExt()         As Byte
    Dim cRead           As Currency
    Dim lSize           As Long
    Dim lResult         As Long
    Dim hCtx            As Long
    Dim uBuf            As ZipBuffersType
    Dim vFileName       As Variant
    Dim uFile           As ZipVfsType
    Dim lCrc32          As Long
    Dim pCrc32          As Long
    Dim bSkip           As Boolean
    Dim sFilterFolder   As String
    Dim sFilterMask     As String
    Dim sOutputFolder   As String
    Dim sOutputFile     As String
#If ImplCrypto Then
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim nPassVer        As Integer
#End If ' ImplCrypto
    
    On Error GoTo EH
    pvSetError
    m_lCurrentFile = -1
    If IsEmpty(m_vArchiveFile) Then
        Err.Raise vbObjectError, , ERR_ARCHIVE_NOT_OPEN
    End If
    ReDim baBuffer(0 To LNG_IOBUF_SIZE - 1) As Byte
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    If VarType(Filter) = vbString Then
        sFilterMask = Replace(Filter, "/", "\")
        lIdx = InStrRev(sFilterMask, "\")
        If lIdx > 0 Then
            sFilterFolder = Left$(sFilterMask, lIdx) & "*"
            sFilterMask = Mid$(sFilterMask, lIdx + 1)
        End If
    End If
    If VarType(OutputTarget) = vbString Then
        sOutputFolder = OutputTarget
        If LenB(sOutputFolder) <> 0 Then
            If (FileAttribs(sOutputFolder) And vbDirectory) = 0 Then
                sOutputFolder = pvPathGetFolder(sOutputFolder, sOutputFile)
                If (FileAttribs(sOutputFolder) And vbDirectory) = 0 Then
                    sOutputFolder = OutputTarget
                    sOutputFile = vbNullString
                End If
            End If
        End If
    End If
    '--- touch to open for read
    pvVfsRead uArchiveFile, 0, 0
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            '--- check if current entry is filtered
            If IsMissing(Filter) Then
                '--- do nothing
            ElseIf IsArray(Filter) Then
                If Not At(Filter, lIdx) Then
                    GoTo SkipFile
                End If
            ElseIf VarType(Filter) = vbString Then
                If LenB(sFilterFolder) <> 0 Then
                    If Not .FileName Like sFilterFolder Then
                        GoTo SkipFile
                    End If
                End If
                If LenB(sFilterMask) <> 0 Then
                    If PathMatchSpecW(StrPtr(Mid$(.FileName, InStrRev(.FileName, "\") + 1)), StrPtr(sFilterMask)) = 0 Then
                        GoTo SkipFile
                    End If
                End If
            ElseIf Not lIdx = Filter Then
                GoTo SkipFile
            End If
            On Error GoTo EH_Continue
            '--- read local header
            pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) + VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = LNG_LEN_LOCAL
            Debug.Assert VarPtr(uLocal.FDate) + 2 - VarPtr(uLocal.Signature) = 14
            pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), 14
            Debug.Assert VarPtr(uLocal.LenExt) + 2 - VarPtr(uLocal.Crc32) = 16
            pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), 16
            Debug.Assert uLocal.LenFname > 0
            If uLocal.LenFname > 0 Then
                ReDim baFName(0 To uLocal.LenFname - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baFName(0)), uLocal.LenFname
            End If
            If uLocal.LenExt > 0 Then
                ReDim baExt(0 To uLocal.LenExt - 1) As Byte
                pvVfsRead uArchiveFile, VarPtr(baExt(0)), uLocal.LenExt
            End If
            '--- sanity checks
            If uLocal.Signature <> LNG_SIG_LOCAL Then
                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & ". " & Replace(ERR_ENTRY_INVALID_SIG, "%1", lIdx + 1)
            End If
'            If pvToWinFileName(baFName, .Flags) <> .FileName Then
'                Err.Raise vbObjectError, , ERR_INVALID_LOCALHDR & " (Filename)"
'            End If
            If .Method <> LNG_METHOD_DEFLATE And .Method <> LNG_METHOD_STORE Then
                If pvSetError(MODULE_NAME & "." & FUNC_NAME, Replace(ERR_UNSUPPORTED_METHOD, "%1", .Method), CanContinue:=True) Then
                    GoTo QH
                Else
                    .Method = LNG_METHOD_STORE
                End If
            End If
            bSkip = False
            If VarType(OutputTarget) = vbString Then
                If LenB(sOutputFile) = 0 Then
                    vFileName = .FileName
                ElseIf InStr(sOutputFile, "*.*") > 0 Then
                    vFileName = Replace(sOutputFile, "*.*", .FileName)
                Else
                    vFileName = Replace(sOutputFile, "*", pvFileGetFName(.FileName))
                End If
                vFileName = PathCombine(sOutputFolder, CStr(vFileName))
                RaiseEvent BeforeExtract(lIdx, vFileName, bSkip, m_bCancel)
            Else
                RaiseEvent BeforeExtract(lIdx, OutputTarget, bSkip, m_bCancel)
            End If
            If m_bCancel Then
                m_sLastError = ERR_USER_CANCEL
                GoTo QH
            End If
            If bSkip Then
                GoTo SkipFile
            End If
            '--- allow output to buffer, stream object or (customized) filename
            If VarType(OutputTarget) = vbString Then
                uFile = pvVfsCreate(vFileName)
            Else
                If IsArray(OutputTarget) And .Size > 0 Then
                    ReDim OutputTarget(0 To .Size - 1) As Byte
                End If
                uFile = pvVfsCreate(OutputTarget)
            End If
            If .compressedSize > 0 Then
                lCrc32 = -1
                pCrc32 = VarPtr(lCrc32)
                cRead = 0
                If (.Flags And zipFlagEncrypted) <> 0 Then
                    RaiseEvent Progress(lIdx, cRead, .compressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    #If ImplCrypto Then
                        If (.Flags And zipFlagStrongEncrypted) <> 0 Then
                            Err.Raise vbObjectError, , ERR_UNSUPPORTED_ENCRYPTION
                        End If
                        If LenB(Password) = 0 Then
                            Err.Raise vbObjectError, , ERR_PASSWORD_REQUIRED
                        End If
                        If .Aes.strength = 0 Then
                            '--- init traditional ZipCrypt
                            If Not pvCryptoTradInit(ToUtf8Array(Password)) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), LNG_ENC_HEADER_SIZE)
                            If lResult = 0 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            If Not pvCryptoTradCrypt(baBuffer, 0, LNG_ENC_HEADER_SIZE, UseAfter:=True) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If baBuffer(LNG_ENC_HEADER_SIZE - 1) <> .DecDat Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            cRead = LNG_ENC_HEADER_SIZE
                        Else
                            '--- init WinZip AES
                            lSaltSize = (.Aes.strength + 1) * 4
                            lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSaltSize + 2)
                            If lResult < lSaltSize + 2 Then
                                Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                            End If
                            ReDim baSalt(0 To lSaltSize - 1) As Byte
                            Call CopyMemory(baSalt(0), baBuffer(0), UBound(baSalt) + 1)
                            If Not pvCryptoAesInit(ToUtf8Array(Password), baSalt, lSaltSize * 2, nPassVer) Then
                                Err.Raise vbObjectError, , m_sLastError
                            End If
                            If nPassVer <> PeekInt(VarPtr(baBuffer(lSaltSize))) Then
                                Err.Raise vbObjectError, , ERR_INVALID_PASSWORD
                            End If
                            '--- don't calc CRC32 for AE-2
                            If .Aes.Version = 2 Then
                                pCrc32 = 0
                            End If
                            '--- note: LNG_AES_AUTHCODE_SIZE chunk comes after file data
                            cRead = lSaltSize + 2 + LNG_AES_AUTHCODE_SIZE
                        End If
                    #Else
                        Err.Raise vbObjectError, , ERR_NO_CRYPTO_COMPILED
                    #End If ' ImplCrypto
                End If
                If .Method = LNG_METHOD_DEFLATE Then
                    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
                    If hCtx = 0 Then
                        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
                    End If
                End If
                Do ' While cRead < .CompressedSize
                    RaiseEvent Progress(lIdx, cRead, .compressedSize, m_bCancel)
                    If m_bCancel Then
                        m_sLastError = ERR_USER_CANCEL
                        GoTo QH
                    End If
                    If cRead >= .compressedSize Then
                        Exit Do
                    End If
                    lSize = LNG_IOBUF_SIZE
                    If lSize > .compressedSize - cRead Then
                        lSize = .compressedSize - cRead
                    End If
                    lResult = pvVfsRead(uArchiveFile, VarPtr(baBuffer(0)), lSize)
                    If lResult = 0 Then
                        Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                    End If
                    #If ImplCrypto Then
                        If (.Flags And zipFlagEncrypted) <> 0 Then
                            If .Aes.strength = 0 Then
                                If Not pvCryptoTradCrypt(baBuffer, 0, lResult, UseAfter:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            Else
                                If Not pvCryptoAesCrypt(baBuffer, 0, lResult, HashBefore:=True) Then
                                    Err.Raise vbObjectError, , m_sLastError
                                End If
                            End If
                        End If
                    #End If ' ImplCrypto
                    cRead = cRead + lResult
                    If .Method = LNG_METHOD_DEFLATE Then
                        uBuf.InBlock = VarPtr(baBuffer(0))
                        uBuf.InLen = lResult
                        lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf), pCrc32)
                        If lResult = 0 Or uBuf.OutBlock = 0 Then
                            Err.Raise vbObjectError, , ERR_DECOMPRESSING
                        End If
                        pvVfsWrite uFile, uBuf.OutBlock, uBuf.OutLen
                        Call CoTaskMemFree(uBuf.OutBlock)
                        uBuf.OutBlock = 0
                    Else ' If .Method = LNG_METHOD_STORE Then
                        If pCrc32 <> 0 Then
                            Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baBuffer(0)), lResult, pCrc32)
                        End If
                        pvVfsWrite uFile, VarPtr(baBuffer(0)), lResult
                    End If
                Loop
                pvVfsSetEof uFile, .FileName
                #If ImplCrypto Then
                    If .Aes.strength <> 0 Then
                        '--- check HMAC-SHA1 auth code
                        ReDim baSalt(0 To LNG_AES_AUTHCODE_SIZE - 1) As Byte
                        lResult = pvVfsRead(uArchiveFile, VarPtr(baSalt(0)), LNG_AES_AUTHCODE_SIZE)
                        If lResult < LNG_AES_AUTHCODE_SIZE Then
                            Err.Raise vbObjectError, , ERR_READING_ARCHIVE
                        End If
                        If Not pvArrayEqual(baSalt, pvCryptoAesGetFinalHash(LNG_AES_AUTHCODE_SIZE)) Then
                            Err.Raise vbObjectError, , ERR_INVALID_AUTHCODE
                        End If
                    End If
                #End If ' ImplCrypto
                If pCrc32 <> 0 Then
                    If .Crc32 <> (lCrc32 Xor -1) Then
                        Err.Raise vbObjectError, , ERR_CRC_CHECK
                    End If
                End If
            Else
                If (.Attributes And vbDirectory + vbVolume) = 0 And Right$(.FileName, 1) <> "\" Then
                    pvVfsSetEof uFile, .FileName
                End If
                '--- note: Total=1 to prevent division by zero in client code
                RaiseEvent Progress(lIdx, 1, 1, m_bCancel)
                If m_bCancel Then
                    m_sLastError = ERR_USER_CANCEL
                    GoTo QH
                End If
            End If
            If VarType(OutputTarget) = vbString Then
                RaiseEvent ExtractComplete(lIdx, vFileName)
            Else
                RaiseEvent ExtractComplete(lIdx, OutputTarget)
            End If
            '--- success (entries extracted)
            Extract = True
SkipFile:
            On Error GoTo EH
            pvVfsClose uFile
            If hCtx <> 0 Then
                Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
                hCtx = 0
            End If
        End With
    Next
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    pvVfsClose uArchiveFile
    pvVfsClose uFile
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
EH_Continue:
    If pvSetError(MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description, CanContinue:=True) Then
        Resume QH
    Else
        Resume SkipFile
    End If
End Function

Private Function pvFromDosDateTime(ByVal nDate As Integer, ByVal nTime As Integer) As Date
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    If nDate <> 0 Or nTime <> 0 Then
        Call DosDateTimeToFileTime(nDate, nTime, uFileTime)
        Call FileTimeToSystemTime(uFileTime, uSysTime)
        With uSysTime
            pvFromDosDateTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
    End If
End Function

Private Function pvToWinFileName(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    '--- replaces non-ANSI symbols with '?' too
    vSplit = Split(Replace(StrConv(StrConv(pvFromOemString(baBuffer, lCodePage), _
        vbFromUnicode), vbUnicode), "/", "\"), "\")
    For lIdx = 0 To UBound(vSplit)
        vSplit(lIdx) = SanitizeFileName(CStr(vSplit(lIdx)), "_")
    Next
    pvToWinFileName = Join(vSplit, "\")
End Function

#End If ' ImplDecompress

#If ImplInflate Then

Public Function InflateBase64(sText As String, baOutput() As Byte) As Boolean
    InflateBase64 = Inflate(FromBase64Array(sText), baOutput)
End Function

Public Function Inflate(baBuffer() As Byte, baOutput() As Byte) As Boolean
    Const FUNC_NAME     As String = "Inflate"
    Dim hCtx            As Long
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    hCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If hCtx = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_DECOMPRESSOR
    End If
    uBuf.InBlock = VarPtr(baBuffer(0))
    uBuf.InLen = UBound(baBuffer) + 1
    lResult = CallWindowProc(m_uRtbl.DecompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_DECOMPRESSING
    End If
    ReDim baOutput(0 To uBuf.OutLen - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal uBuf.OutBlock, uBuf.OutLen)
    '--- success
    Inflate = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hCtx)
        hCtx = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function DeflateBase64(baInput() As Byte, sText As String, Optional ByVal Level As Long = 6) As Boolean
    Dim baOutput()      As Byte
    
    DeflateBase64 = Deflate(baInput, baOutput, Level)
    If DeflateBase64 Then
        sText = ToBase64String(baOutput)
    End If
End Function

Public Function Deflate(baBuffer() As Byte, baOutput() As Byte, Optional ByVal Level As Long = 6) As Boolean
    Const FUNC_NAME     As String = "Deflate"
    Dim hCtx            As Long
    Dim lOutputPtr      As Long
    Dim lOutputSize     As Long
    
    On Error GoTo EH
    hCtx = DeflateInit()
    If Not DeflateBlob(hCtx, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, lOutputPtr, lOutputSize, Level, 1) Then
        GoTo QH
    End If
    ReDim baOutput(0 To lOutputSize - 1) As Byte
    Call CopyMemory(baOutput(0), ByVal lOutputPtr, lOutputSize)
    '--- success
    Deflate = True
QH:
    On Error Resume Next
    If lOutputPtr <> 0 Then
        Call CoTaskMemFree(lOutputPtr)
    End If
    If hCtx <> 0 Then
        DeflateEnd hCtx
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Function DeflateInit() As Long
    DeflateInit = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If DeflateInit = 0 Then
        Err.Raise vbObjectError, , ERR_INIT_COMPRESSOR
    End If
End Function

Public Function DeflateBlob(ByVal hCtx As Long, ByVal lInputPtr As Long, ByVal lInputSize As Long, lOutputPtr As Long, lOutputSize As Long, Optional ByVal Level As Long = 6, Optional ByVal Final As Boolean) As Boolean
    Const FUNC_NAME     As String = "DeflateBlob"
    Dim uBuf            As ZipBuffersType
    Dim lResult         As Long
    
    On Error GoTo EH
    uBuf.Greedy = (Level <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), Level)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), Level)
    uBuf.InBlock = lInputPtr
    uBuf.InLen = lInputSize
    uBuf.Final = -Final
    lResult = CallWindowProc(m_uRtbl.CompressBlock, hCtx, VarPtr(uBuf))
    If lResult = 0 Or uBuf.OutBlock = 0 Then
        Err.Raise vbObjectError, , ERR_COMPRESSING
    End If
    '--- commit
    lOutputPtr = uBuf.OutBlock
    lOutputSize = uBuf.OutLen
    uBuf.OutBlock = 0
    '--- success
    DeflateBlob = True
QH:
    On Error Resume Next
    If uBuf.OutBlock <> 0 Then
        Call CoTaskMemFree(uBuf.OutBlock)
        uBuf.OutBlock = 0
    End If
    Exit Function
EH:
    pvSetError MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
    Resume QH
End Function

Public Sub DeflateEnd(ByVal hCtx As Long)
    If hCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCtx)
    End If
End Sub

#End If ' ImplInflate

'= Crypto ================================================================

#If ImplCrypto Then

Private Sub pvCryptoGetRandom(ByVal lSize As Long, baOutput() As Byte)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call RtlGenRandom(baOutput(0), lSize)
End Sub

Private Function pvCryptoAesInit(baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long, nPassVer As Integer) As Boolean
    Dim baDerivedKey()  As Byte
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- init member vars
    m_uCrypto.Nonce(0) = 0
    m_uCrypto.Nonce(1) = 0
    m_uCrypto.EncrData = vbNullString
    m_uCrypto.EncrPos = 0
    '--- generate RFC 2898 based derived key
    If m_uCrypto.hPbkdf2Alg = 0 Then
        '--- CNG API missing on XP
        On Error GoTo EH_Unsupported
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hPbkdf2Alg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        On Error GoTo 0
    End If
    ReDim baDerivedKey(0 To 2 * lKeyLen + 1) As Byte
    '--- PBKDF2 API missing on Vista
    On Error GoTo EH_Unsupported
    hResult = BCryptDeriveKeyPBKDF2(m_uCrypto.hPbkdf2Alg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, 1000, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    On Error GoTo 0
    '--- extract Password Verification Value from last 2 bytes of derived key
    Call CopyMemory(nPassVer, baDerivedKey(2 * lKeyLen), 2)
    '--- init AES key from first half of the derived key
    If m_uCrypto.hAesAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hAesAlg, StrPtr(BCRYPT_AES_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), 0)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_AES_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_OBJECT_LENGTH), m_uCrypto.AesKeyObjLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_OBJECT_LENGTH & ")"
            GoTo QH
        End If
        hResult = BCryptSetProperty(m_uCrypto.hAesAlg, StrPtr(BCRYPT_CHAINING_MODE), StrPtr(BCRYPT_CHAIN_MODE_ECB), LenB(BCRYPT_CHAIN_MODE_ECB), 0)
        If hResult < 0 Then
            sApiSource = "BCryptSetProperty(" & BCRYPT_CHAINING_MODE & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    ReDim m_uCrypto.AesKeyObjData(0 To m_uCrypto.AesKeyObjLen - 1) As Byte
    hResult = BCryptGenerateSymmetricKey(m_uCrypto.hAesAlg, m_uCrypto.hAesKey, m_uCrypto.AesKeyObjData(0), m_uCrypto.AesKeyObjLen, baDerivedKey(0), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptGenerateSymmetricKey"
        GoTo QH
    End If
    '--- init HMAC key from second half of the derived key
    If m_uCrypto.hHmacAlg = 0 Then
        hResult = BCryptOpenAlgorithmProvider(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_SHA1_ALGORITHM), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG)
        If hResult < 0 Then
            sApiSource = "BCryptOpenAlgorithmProvider(" & BCRYPT_SHA1_ALGORITHM & ")"
            GoTo QH
        End If
        hResult = BCryptGetProperty(m_uCrypto.hHmacAlg, StrPtr(BCRYPT_HASH_LENGTH), m_uCrypto.HmacLen, 4, lDummy, 0)
        If hResult < 0 Then
            sApiSource = "BCryptGetProperty(" & BCRYPT_HASH_LENGTH & ")"
            GoTo QH
        End If
    End If
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    hResult = BCryptCreateHash(m_uCrypto.hHmacAlg, m_uCrypto.hHmacHash, 0, 0, baDerivedKey(lKeyLen), lKeyLen, 0)
    If hResult < 0 Then
        sApiSource = "BCryptCreateHash"
        GoTo QH
    End If
    '--- success
    pvCryptoAesInit = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
    Exit Function
EH_Unsupported:
    m_sLastError = ERR_UNSUPPORTED_ENCRYPTION
End Function

Private Sub pvCryptoTerminate()
    If m_uCrypto.hHmacHash <> 0 Then
        Call BCryptDestroyHash(m_uCrypto.hHmacHash)
        m_uCrypto.hHmacHash = 0
    End If
    If m_uCrypto.hHmacAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hHmacAlg, 0)
        m_uCrypto.hHmacAlg = 0
    End If
    If m_uCrypto.hAesKey <> 0 Then
        Call BCryptDestroyKey(m_uCrypto.hAesKey)
        m_uCrypto.hAesKey = 0
    End If
    If m_uCrypto.hAesAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(m_uCrypto.hAesAlg, 0)
        m_uCrypto.hAesAlg = 0
    End If
End Sub

Private Function pvCryptoAesCrypt( _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long, _
            Optional ByVal HashBefore As Boolean, _
            Optional ByVal HashAfter As Boolean) As Boolean
    Dim lIdx            As Long
    Dim lPadSize        As Long
    Dim lDummy          As Long '--- discarded
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    If HashBefore Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    With m_uCrypto
        '--- reuse EncrData from prev call until next LNG_AES_BLOCK_SIZE boundary
        For lIdx = Offset To Offset + Size - 1
            If (.EncrPos And (LNG_AES_BLOCK_SIZE - 1)) = 0 Then
                Exit For
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
            .EncrPos = .EncrPos + 1
        Next
        If lIdx < Offset + Size Then
            '--- pad remaining input size to LNG_AES_BLOCK_SIZE
            lPadSize = (Offset + Size - lIdx + LNG_AES_BLOCK_SIZE - 1) And -LNG_AES_BLOCK_SIZE
            If UBound(.EncrData) + 1 < lPadSize Then
                ReDim .EncrData(0 To lPadSize - 1) As Byte
            End If
            '--- encrypt incremental nonces in EncrData
            Call CallWindowProc(m_uRtbl.MemNonce, VarPtr(.EncrData(0)), VarPtr(.Nonce(0)), lPadSize)
            hResult = BCryptEncrypt(.hAesKey, .EncrData(0), lPadSize, 0, 0, 0, .EncrData(0), lPadSize, lDummy, 0)
            If hResult < 0 Then
                sApiSource = "BCryptEncrypt"
                GoTo QH
            End If
            '--- xor remaining input and leave anything extra of EncrData for reuse
            .EncrPos = Offset + Size - lIdx
            Call CallWindowProc(m_uRtbl.MemXor, VarPtr(.EncrData(0)), VarPtr(baData(lIdx)), .EncrPos)
        End If
    End With
    If HashAfter Then
        hResult = BCryptHashData(m_uCrypto.hHmacHash, baData(Offset), Size, 0)
        If hResult < 0 Then
            sApiSource = "BCryptHashData"
            GoTo QH
        End If
    End If
    '--- success
    pvCryptoAesCrypt = True
    Exit Function
QH:
    If hResult < 0 Then
        m_sLastError = GetSystemMessage(hResult) & " [" & sApiSource & "]"
    ElseIf Err.LastDllError <> 0 Then
        m_sLastError = GetSystemMessage(Err.LastDllError)
    End If
End Function

Private Function pvCryptoAesGetFinalHash(ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To m_uCrypto.HmacLen - 1) As Byte
    Call BCryptFinishHash(m_uCrypto.hHmacHash, baResult(0), m_uCrypto.HmacLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    pvCryptoAesGetFinalHash = baResult
End Function

Private Function pvCryptoTradInit(baPass() As Byte) As Boolean
    With m_uCrypto
        .TradKey(0) = &H12345678
        .TradKey(1) = &H23456789
        .TradKey(2) = &H34567890
        .TradKey(3) = m_uRtbl.Crc32Table
        Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(.TradKey(0)), VarPtr(baPass(0)), UBound(baPass) + 1)
    End With
    '--- success
    pvCryptoTradInit = True
End Function

Private Function pvCryptoTradCrypt( _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long = -1, _
            Optional ByVal UseAfter As Boolean) As Boolean
    If Size < 0 Then
        Size = UBound(baData) + 1 - Offset
    End If
    Call CallWindowProc(m_uRtbl.ZipCrypt, VarPtr(m_uCrypto.TradKey(0)), VarPtr(baData(Offset)), Size, IIf(UseAfter, 3, 1))
    '--- success
    pvCryptoTradCrypt = True
End Function

#End If ' ImplCrypto

'= VFS ===================================================================

Private Function pvVfsOpen(File As Variant) As ZipVfsType
    Dim hFind           As Long
    
    If IsArray(File) Then
        pvVfsOpen.FileName = STR_BUFFER
        If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
            Call CopyMemory(pvVfsOpen.BufferArray, File, 16)
        Else
            pvVfsOpen.BufferArray = File
        End If
        pvVfsOpen.BufferBase = pvArrPtr(File)
        pvVfsOpen.BufferPtr = pvVfsOpen.BufferBase
        If pvVfsOpen.BufferBase <> 0 Then
            pvVfsOpen.bufferSize = UBound(File) + 1
        End If
        pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.bufferSize
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
    ElseIf IsObject(File) Then
        pvVfsOpen.FileName = STR_STREAM
        Set pvVfsOpen.Stream = File
        If Not pvVfsOpen.Stream Is Nothing Then
            pvVfsOpen.Data.nFileSizeLow = pvVfsOpen.Stream.VfsSetFilePointer(0, FILE_END)
        End If
        pvVfsOpen.Data.dwFileAttributes = vbArchive
        pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
    Else
        pvVfsOpen.FileName = File
        If Right$(File, 1) = ":" Then
            pvVfsOpen.Data.dwFileAttributes = vbArchive
            pvVfsOpen.Data.ftLastWriteTime = pvToFileTime(VBA.Now)
        ElseIf Right$(File, 1) <> "\" Then
            hFind = FindFirstFile(StrPtr(File), VarPtr(pvVfsOpen.Data))
        Else
            hFind = FindFirstFile(StrPtr(Left$(File, Len(File) - 1)), VarPtr(pvVfsOpen.Data))
        End If
        If hFind <> INVALID_HANDLE_VALUE Then
            Call FindClose(hFind)
        Else
            Err.Raise vbObjectError, "pvVfsOpen", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
        End If
    End If
End Function

Private Function pvCreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Select Case LCase$(lpFileName)
    Case "stdin:"
        pvCreateFile = GetStdHandle(STD_INPUT_HANDLE)
    Case "stdout:"
        pvCreateFile = GetStdHandle(STD_OUTPUT_HANDLE)
    Case "stderr:"
        pvCreateFile = GetStdHandle(STD_ERROR_HANDLE)
    Case Else
        pvCreateFile = CreateFile(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, NoSecurity, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    End Select
End Function

Private Function pvVfsCreate(File As Variant) As ZipVfsType
    Dim sError          As String
    
    If IsArray(File) Then
        pvVfsCreate.FileName = STR_BUFFER
        If (Peek(VarPtr(File)) And VT_BYREF) <> 0 Then
            Call CopyMemory(pvVfsCreate.BufferArray, File, 16)
        Else
            Err.Raise vbObjectError, "pvVfsCreate", ERR_INVALID_BYTEARRAY
        End If
        pvVfsCreate.BufferBase = pvArrPtr(pvVfsCreate.BufferArray)
        pvVfsCreate.BufferPtr = pvVfsCreate.BufferBase
        If pvVfsCreate.BufferBase <> 0 Then
            pvVfsCreate.bufferSize = UBound(pvVfsCreate.BufferArray) + 1
        End If
    ElseIf IsObject(File) Then
        pvVfsCreate.FileName = STR_STREAM
        Set pvVfsCreate.Stream = File
    Else
        pvVfsCreate.FileName = File
        If InStrRev(File, "\") > 0 Then
            If Not MkPath(Left$(File, InStrRev(File, "\") - 1), sError) Then
                Err.Raise vbObjectError, , sError
            End If
        End If
        If Right$(File, 1) <> "\" Then
            pvVfsCreate.Handle = pvCreateFile(File, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If pvVfsCreate.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsCreate", GetSystemMessage(Err.LastDllError) & " (" & File & ")"
            End If
        End If
    End If
End Function

Private Sub pvVfsClose(uFile As ZipVfsType)
    Dim vEmpty          As Variant
    
    If Not IsEmpty(uFile.BufferArray) Then
        If (Peek(VarPtr(uFile.BufferArray)) And VT_BYREF) <> 0 Then
            Call CopyMemory(uFile.BufferArray, vEmpty, 16)
        Else
            uFile.BufferArray = Empty
        End If
    End If
    If uFile.Handle <> 0 And uFile.Handle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(uFile.Handle)
        uFile.Handle = 0
    End If
    Set uFile.Stream = Nothing
End Sub

Private Function pvVfsRead(uFile As ZipVfsType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Not IsEmpty(uFile.BufferArray) Then
        pvVfsRead = Clamp(lSize, 0, uFile.bufferSize - (uFile.BufferPtr - uFile.BufferBase))
        Call CopyMemory(ByVal lPtr, ByVal uFile.BufferPtr, pvVfsRead)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsRead
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsRead = uFile.Stream.VfsReadFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsRead", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If ReadFile(uFile.Handle, ByVal lPtr, lSize, pvVfsRead, 0) = 0 Then
            If Err.LastDllError <> LNG_PIPE_ENDED_ERROR Then
                Err.Raise vbObjectError, "pvVfsRead", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
    End If
End Function

Private Function pvVfsWrite(uFile As ZipVfsType, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Const MAX_STEP      As Long = 2 ^ 24 '--- 16MB
    Dim lOffset         As Long
    
    If Not IsEmpty(uFile.BufferArray) Then
        lOffset = uFile.BufferPtr - uFile.BufferBase
        If lOffset + lSize > uFile.bufferSize Then
            If lOffset + lSize <= MAX_STEP Then
                uFile.bufferSize = Clamp(2 ^ Int(Log(lOffset + lSize) / Log(2) + 1), 8192)
            Else
                uFile.bufferSize = (lOffset + lSize + MAX_STEP - 1) And -MAX_STEP
            End If
            ReDim Preserve uFile.BufferArray(0 To uFile.bufferSize - 1) As Byte
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + lOffset
        End If
        pvVfsWrite = lSize
        Call CopyMemory(ByVal uFile.BufferPtr, ByVal lPtr, pvVfsWrite)
        uFile.BufferPtr = uFile.BufferPtr + pvVfsWrite
    ElseIf Not uFile.Stream Is Nothing Then
        If uFile.Handle = 0 Then
            uFile.Handle = INVALID_HANDLE_VALUE
            uFile.Stream.VfsSetFilePointer 0, FILE_BEGIN
        End If
        pvVfsWrite = uFile.Stream.VfsWriteFile(lPtr, lSize)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            uFile.Handle = pvCreateFile(uFile.FileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive, 0)
            If uFile.Handle = INVALID_HANDLE_VALUE Then
                Err.Raise vbObjectError, "pvVfsWrite", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
            End If
        End If
        If WriteFile(uFile.Handle, ByVal lPtr, lSize, pvVfsWrite, 0) = 0 Then
            Err.Raise vbObjectError, "pvVfsWrite", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Function

Private Function pvVfsSeek(uFile As ZipVfsType, ByVal cPosition As Currency, ByVal lMoveMethod As Long) As Currency
    If Not IsEmpty(uFile.BufferArray) Then
        Select Case lMoveMethod
        Case FILE_BEGIN
            pvVfsSeek = cPosition
        Case FILE_CURRENT
            pvVfsSeek = uFile.BufferPtr - uFile.BufferBase + cPosition
        Case FILE_END
            pvVfsSeek = uFile.bufferSize + cPosition
        End Select
        uFile.BufferPtr = uFile.BufferBase + pvVfsSeek
    ElseIf Not uFile.Stream Is Nothing Then
        pvVfsSeek = uFile.Stream.VfsSetFilePointer(cPosition, lMoveMethod)
    Else
        If uFile.Handle = 0 Or uFile.Handle = INVALID_HANDLE_VALUE Then
            '--- touch to open for read
            pvVfsRead uFile, 0, 0
        End If
        If SetFilePointerEx(uFile.Handle, cPosition / 10000@, pvVfsSeek, lMoveMethod) = 0 Then
            Err.Raise vbObjectError, "pvVfsSeek", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
        pvVfsSeek = pvVfsSeek * 10000@
    End If
End Function

Private Sub pvVfsSetEof(uFile As ZipVfsType, sMetaData As String)
    If Not IsEmpty(uFile.BufferArray) Then
        If uFile.bufferSize <> uFile.BufferPtr - uFile.BufferBase Then
            uFile.bufferSize = uFile.BufferPtr - uFile.BufferBase
            If uFile.bufferSize > 0 Then
                ReDim Preserve uFile.BufferArray(0 To uFile.bufferSize - 1) As Byte
            Else
                Erase uFile.BufferArray
            End If
            uFile.BufferBase = pvArrPtr(uFile.BufferArray)
            uFile.BufferPtr = uFile.BufferBase + uFile.bufferSize
        End If
    ElseIf Not uFile.Stream Is Nothing Then
        uFile.Stream.VfsSetEndOfFile sMetaData
    ElseIf Right$(uFile.FileName, 1) <> ":" Then
        If SetEndOfFile(uFile.Handle) = 0 Then
            Err.Raise vbObjectError, "pvVfsSetEof", GetSystemMessage(Err.LastDllError) & " (" & uFile.FileName & ")"
        End If
    End If
End Sub

Private Function pvArrPtr(vArray As Variant) As Long
    Dim lPtr            As Long
    
    lPtr = Peek(UnsignedAdd(VarPtr(vArray), 8))
    If (Peek(VarPtr(vArray)) And VT_BYREF) <> 0 Then
        lPtr = Peek(lPtr)
    End If
    If lPtr <> 0 Then
        pvArrPtr = Peek(UnsignedAdd(lPtr, 12))
    End If
End Function

'= common ================================================================

Private Function pvFromOemString(baBuffer() As Byte, ByVal lCodePage As Long) As String
    Dim lSize           As Long
    
    If UBound(baBuffer) >= 0 Then
        lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, 0, 0)
        If lSize > 0 Then
            pvFromOemString = String$(lSize, 0)
            lSize = MultiByteToWideChar(lCodePage, 0, baBuffer(0), UBound(baBuffer) + 1, ByVal StrPtr(pvFromOemString), lSize)
        End If
    End If
End Function

Private Function pvFromFileTime(uFileTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME
    
    If FileTimeToLocalFileTime(uFileTime, uLocalTime) <> 0 Then
        If FileTimeToSystemTime(uLocalTime, uSysTime) <> 0 Then
            Call SystemTimeToVariantTime(uSysTime, pvFromFileTime)
        End If
    End If
End Function

Private Function pvToFileTime(dDate As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uLocalTime      As FILETIME

    If VariantTimeToSystemTime(dDate, uSysTime) <> 0 Then
        If SystemTimeToFileTime(uSysTime, uLocalTime) <> 0 Then
            Call LocalFileTimeToFileTime(uLocalTime, pvToFileTime)
        End If
    End If
End Function

Private Function pvToInt64(ByVal lLowPart As Long, Optional ByVal lHiPart As Long) As Currency
    Call CopyMemory(pvToInt64, lLowPart, 4)
    If lHiPart <> 0 Then
        Call CopyMemory(ByVal UnsignedAdd(VarPtr(pvToInt64), 4), lHiPart, 4)
    End If
    pvToInt64 = pvToInt64 * 10000@
End Function

Private Function pvToInt32(ByVal cValue As Currency) As Long
    cValue = cValue / 10000@
    Call CopyMemory(pvToInt32, cValue, 4)
End Function

Private Function pvSetError(Optional Source As String, Optional Description As String, Optional ByVal CanContinue As Boolean) As Boolean
    If StrPtr(Source) = 0 And StrPtr(Description) = 0 Then
        m_sLastError = vbNullString
        m_bCancel = False
    Else
        #If ImplLogging Then
            If m_lCurrentFile >= 0 Then
                Debug.Print "Error in " & Source & ": " & Description, m_uFiles(m_lCurrentFile).FileName, Timer
            Else
                Debug.Print "Error in " & Source & ": " & Description, Timer
            End If
        #End If ' ImplLogging
        m_sLastError = Description
        If Not CanContinue Then
            m_bCancel = True
        End If
        RaiseEvent Error(m_lCurrentFile, Source, m_sLastError, m_bCancel)
        pvSetError = m_bCancel
    End If
End Function

Private Function pvPathGetFolder(sPath As String, Optional sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sPath, "\")
    If lPos > 0 Then
        pvPathGetFolder = Left$(sPath, lPos - 1)
        sFile = Mid$(sPath, lPos + 1)
    Else
        pvPathGetFolder = sPath
        sFile = vbNullString
    End If
End Function

Private Function pvFileGetFName(sFile As String) As String
    Dim lPos            As Long
    
    lPos = InStrRev(sFile, ".")
    If lPos > InStrRev(sFile, "\") Then
        pvFileGetFName = Left$(sFile, lPos - 1)
    Else
        pvFileGetFName = sFile
    End If
End Function

Private Function pvArrayEqual(baFirst() As Byte, baSecond() As Byte) As Boolean
    If UBound(baFirst) = UBound(baSecond) Then
        pvArrayEqual = (InStrB(baFirst, baSecond) = 1)
    End If
End Function

'= zlib thunks ===========================================================

Private Function pvInitRelocTable(uRtbl As ZipRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(zipRelCompressInit)
        .CompressCleanup = lpThunk + vSplit(zipRelCompressCleanup)
        .CompressBlock = lpThunk + vSplit(zipRelCompressBlock)
        .DecompressInit = lpThunk + vSplit(zipRelDecompressInit)
        .DecompressCleanup = lpThunk + vSplit(zipRelDecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(zipRelDecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(zipRelCalcCrc32)
        .MemNonce = lpThunk + vSplit(zipRelMemNonce)
        .MemXor = lpThunk + vSplit(zipRelMemXor)
        .ZipCrypt = lpThunk + vSplit(zipRelZipCrypt)
        .MallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle(StrPtr("ole32")), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(zipRelLenCodes)
        .DistCodes = lpThunk + vSplit(zipRelDistCodes)
        .MirrorBytes = lpThunk + vSplit(zipRelMirrorBytes)
        .LenLenMap = lpThunk + vSplit(zipRelLenLenMap)
        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    Dim sBuffer         As String
    
    If lpThunk = 0 Then
        sBuffer = String$(50, 0)
        Call GetEnvironmentVariable(StrPtr("_ZIP_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(sBuffer), Len(sBuffer) - 1)
        lpThunk = Val(sBuffer)
        If lpThunk = 0 Then
            baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2)
            lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
            Call SetEnvironmentVariable(StrPtr("_ZIP_THUNK_" & GetCurrentProcessId() & "_" & STR_THUNK_BUILDDATE), StrPtr(lpThunk))
        End If
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function pvGetCrc32Table() As Long
    Static aTable()     As Long
    Static bIsInit      As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lReminder       As Long
    Dim lValue          As Long

    If Not bIsInit Then
        '--- table mem allocated: 4KB
        ReDim aTable(0 To &H3FF) As Long
        For lIdx = 0 To &H3FF
            If lIdx < &H100 Then
                lReminder = 0
                lValue = lIdx
            Else
                lReminder = aTable(lIdx - &H100)
                lValue = 0
            End If
            For lJdx = 1 To 8
                If ((lReminder Xor lValue) And 1) <> 0 Then
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF Xor &HEDB88320
                Else
                  lReminder = (lReminder And &HFFFFFFFE) \ 2 And &H7FFFFFFF
                End If
                lValue = (lValue And &HFFFFFFFE) \ 2 And &H7FFFFFFF
            Next
            aTable(lIdx) = lReminder
        Next
        bIsInit = True
    End If
    pvGetCrc32Table = VarPtr(aTable(0))
End Function

Public Function CalcCrc32Array(baData() As Byte) As Long
    CalcCrc32Array = -1
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), VarPtr(baData(0)), UBound(baData) + 1, VarPtr(CalcCrc32Array))
    CalcCrc32Array = CalcCrc32Array Xor -1
End Function

Public Sub CalcCrc32Ptr(ByVal lPtr As Long, ByVal lSize As Long, lCrc32 As Long)
    Call CallWindowProc(m_uRtbl.CalcCrc32, VarPtr(m_uRtbl), lPtr, lSize, VarPtr(lCrc32))
End Sub

'= shared ================================================================

#If ImplUseShared = 0 Then

Private Function FileAttribs(sFile As String) As VbFileAttribute
    FileAttribs = GetFileAttributes(StrPtr(sFile))
    If FileAttribs = -1 Then
        FileAttribs = &H8000&
    End If
End Function

Private Function MkPath(sPath As String, sError As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(StrPtr(sPath))
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1), sError) Then
                Exit Function
            End If
        End If
        If CreateDirectory(StrPtr(sPath), 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError) & " (" & sPath & ")"
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
        sError = "File already exists (" & sPath & ")"
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Private Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = String$(2000, 0)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDllError, 0, StrPtr(GetSystemMessage), Len(GetSystemMessage), 0)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, lSize) & " &H" & Hex$(lLastDllError)
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Private Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
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

Private Function ToBase64String(baInput() As Byte) As String
    Dim lSize As Long
    If (CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, 0&, lSize) <> 0) Then
        ToBase64String = String$(lSize - 1, 0)
        Call CryptBinaryToString(VarPtr(baInput(LBound(baInput))), UBound(baInput) - LBound(baInput) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64String), lSize)
    End If
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

Private Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal Min As Long = -2147483647, _
            Optional ByVal Max As Long = 2147483647) As Long
    If lValue < Min Then
        Clamp = Min
    ElseIf lValue > Max Then
        Clamp = Max
    Else
        Clamp = lValue
    End If
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), VarPtr(baRetVal(0)), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, ByVal StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function C_Obj(Value As Variant) As Object
    On Error Resume Next
    Set C_Obj = Value
    On Error GoTo 0
End Function

Private Function SanitizeFileName(sFileName As String, Optional InvalidChar As String = " ") As String
    Const STR_INVALID   As String = "[\x00-\x1F""<>\|:\*\?\\/]"
    
    With InitRegExp(STR_INVALID)
        SanitizeFileName = .Replace(sFileName, InvalidChar)
    End With
End Function

Private Function InitRegExp(sPattern As String) As Object
    Static oCache       As Object
    
    If Not oCache Is Nothing Then
        If oCache.Pattern = sPattern Then
            GoTo QH
        End If
    End If
    Set oCache = VBA.CreateObject("VBScript.RegExp")
    oCache.Global = True
    oCache.Pattern = sPattern
QH:
    Set InitRegExp = oCache
End Function

#End If ' Not ImplUseShared

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    pvInitRelocTable m_uRtbl
    m_lCurrentFile = -1
    m_lCodePage = CP_OEMCP
End Sub

#If ImplCrypto Then
    Private Sub Class_Terminate()
        pvCryptoTerminate
    End Sub
#End If ' ImplCrypto

'=========================================================================
' CHANGES FOR PHOTODEMON USE ONLY
'=========================================================================

'For use in PD by the OpenRaster importer (which has already validated srcFilename); do *NOT* use elsewhere.
Friend Function PeekSingleFile(ByVal srcFilename As String, ByRef dstOffset As Long, ByRef dstSizeDeflate As Long, ByRef dstSizeInflate As Long, ByRef dstItemIsCompressed As Boolean) As Boolean
    
    If (InStr(1, srcFilename, "/") <> 0) Then srcFilename = Replace(srcFilename, "/", "\")
    
    Dim uArchiveFile    As ZipVfsType
    Dim lIdx            As Long
    Dim uLocal          As ZipLocalHeaderType
    
    m_lCurrentFile = -1
    uArchiveFile = pvVfsOpen(m_vArchiveFile)
    pvVfsRead uArchiveFile, 0, 0
    For lIdx = 0 To m_lFileCount - 1
        m_lCurrentFile = lIdx
        With m_uFiles(lIdx)
            If Strings.StringsEqual(.FileName, srcFilename, True) Then
                dstOffset = .Offset
                dstSizeDeflate = .compressedSize
                dstSizeInflate = .Size
                dstItemIsCompressed = (.Method = LNG_METHOD_DEFLATE)
                PeekSingleFile = True
                
                'We now need to advance the offset by an unpredictable number of bytes.
                ' (The exact length depends on the local header in the zip file, unfortunately.)
                pvVfsSeek uArchiveFile, .Offset, FILE_BEGIN
                
                'Read uLocal members in two segments, to avoid word-aligned padding issues in the struct
                pvVfsRead uArchiveFile, VarPtr(uLocal.Signature), 14
                pvVfsRead uArchiveFile, VarPtr(uLocal.Crc32), 16
                
                'Final offset is "base offset + 30 bytes (local header size 14+16, above) + len(filename) + len(extra field)"
                dstOffset = dstOffset + 30& + uLocal.LenFname + uLocal.LenExt
                
                Exit For
            End If
            
        End With
    Next lIdx
    
    pvVfsClose uArchiveFile
    
End Function
