Imports System.Runtime.InteropServices
Public Class FanucClass
    Public Structure ODBSYS
        Public addinfo As Short
        Public max_axis As Short
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public cnc_type As Char()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public mt_type As Char()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public series As Char()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public version As Char()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public axes As Char()
    End Structure

    ' cnc_rdalmhistry:read alarm history data 
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=4)>
    Public Structure ALM_HIS_data
        Public dummy As Short
        Public alm_grp As Short       ' alarm group 
        Public alm_no As Short        ' alarm number 
        Public axis_no As Byte        ' axis number 
        Public year As Byte           ' year 
        Public month As Byte          ' month 
        Public day As Byte            ' day 
        Public hour As Byte           ' hour 
        Public minute As Byte         ' minute 
        Public second As Byte         ' second 
        Public dummy2 As Byte
        Public len_msg As Short       ' alarm message length 
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public alm_msg As String      ' alarm message 
    End Structure
    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure ALM_HIS1
        Public data1 As ALM_HIS_data
        Public data2 As ALM_HIS_data
        Public data3 As ALM_HIS_data
        Public data4 As ALM_HIS_data
        Public data5 As ALM_HIS_data
        Public data6 As ALM_HIS_data
        Public data7 As ALM_HIS_data
        Public data8 As ALM_HIS_data
        Public data9 As ALM_HIS_data
        Public data10 As ALM_HIS_data
        Public data11 As ALM_HIS_data
        Public data12 As ALM_HIS_data
        Public data13 As ALM_HIS_data
        Public data14 As ALM_HIS_data
        Public data15 As ALM_HIS_data
        Public data16 As ALM_HIS_data
        Public data17 As ALM_HIS_data
        Public data18 As ALM_HIS_data
        Public data19 As ALM_HIS_data
        Public data20 As ALM_HIS_data
        Public data21 As ALM_HIS_data
        Public data22 As ALM_HIS_data
        Public data23 As ALM_HIS_data
        Public data24 As ALM_HIS_data
        Public data25 As ALM_HIS_data
        Public data26 As ALM_HIS_data
        Public data27 As ALM_HIS_data
        Public data28 As ALM_HIS_data
        Public data29 As ALM_HIS_data
        Public data30 As ALM_HIS_data
        Public data31 As ALM_HIS_data
        Public data32 As ALM_HIS_data
        Public data33 As ALM_HIS_data
        Public data34 As ALM_HIS_data
        Public data35 As ALM_HIS_data
        Public data36 As ALM_HIS_data
        Public data37 As ALM_HIS_data
        Public data38 As ALM_HIS_data
        Public data39 As ALM_HIS_data
        Public data40 As ALM_HIS_data
        Public data41 As ALM_HIS_data
        Public data42 As ALM_HIS_data
        Public data43 As ALM_HIS_data
        Public data44 As ALM_HIS_data
        Public data45 As ALM_HIS_data
        Public data46 As ALM_HIS_data
        Public data47 As ALM_HIS_data
        Public data48 As ALM_HIS_data
        Public data49 As ALM_HIS_data
        Public data50 As ALM_HIS_data
    End Structure ' In case that the number of data is 10 
    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure ODBAHIS
        Public s_no As Short  ' start number  C# ushort
        Public type As Short  ' dummy 
        Public e_no As Short  ' end number  C# ushort
        Public alm_his As ALM_HIS1
    End Structure
    Public Structure IODBTIME
        Public minute As Integer
        Public msec As Integer
    End Structure
    ' cnc_statinfo:read CNC status information 
    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure ODBST
        Public dummy As Short     ' dummy 
        Public tmmode As Short    ' T/M mode 
        Public aut As Short       ' selected automatic mode 
        Public run As Short       ' running status 
        Public motion As Short    ' axis, dwell status 
        Public mstb As Short      ' m, s, t, b status 
        Public emergency As Short ' emergency stop status 
        Public alarm As Short     ' alarm status 
        Public edit As Short      ' editting status 
    End Structure
    ' cnc_rdomhistry:read operator message history 
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=4)>
    Public Structure ODBOMHIS_data
        Public om_no As Short ' operator message number 
        Public year As Short  ' year 
        Public month As Short ' month 
        Public day As Short   ' day 
        Public hour As Short  ' hour 
        Public minute As Short ' mimute 
        Public second As Short ' second 
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public om_msg As String
    End Structure
    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure ODBOMHIS
        Public omhis1 As ODBOMHIS_data
        Public omhis2 As ODBOMHIS_data
        Public omhis3 As ODBOMHIS_data
        Public omhis4 As ODBOMHIS_data
        Public omhis5 As ODBOMHIS_data
        Public omhis6 As ODBOMHIS_data
        Public omhis7 As ODBOMHIS_data
        Public omhis8 As ODBOMHIS_data
        Public omhis9 As ODBOMHIS_data
        Public omhis10 As ODBOMHIS_data
    End Structure ' In case that the number of data is 10 
    Function new_coding(ByVal str As String) As String
        Dim charw As Char,
            newstr As String

        For i = 1 To Len(str)
            charw = Mid(str, i, 1)

            Select Case charw
                Case "Ж"
                    newstr = newstr + "Ё"
                Case "З"
                    newstr = newstr + "Ж"
                Case "И"
                    newstr = newstr + "З"
                Case "Й"
                    newstr = newstr + "И"
                Case "К"
                    newstr = newstr + "Й"
                Case "Л"
                    newstr = newstr + "К"
                Case "М"
                    newstr = newstr + "Л"
                Case "Н"
                    newstr = newstr + "М"
                Case "О"
                    newstr = newstr + "Н"
                Case "П"
                    newstr = newstr + "О"
                Case "Р"
                    newstr = newstr + "П"
                Case "С"
                    newstr = newstr + "Р"
                Case "Т"
                    newstr = newstr + "С"
                Case "У"
                    newstr = newstr + "Т"
                Case "Ф"
                    newstr = newstr + "У"
                Case "Х"
                    newstr = newstr + "Ф"
                Case "Ц"
                    newstr = newstr + "Х"
                Case "Ч"
                    newstr = newstr + "Ц"
                Case "Ш"
                    newstr = newstr + "Ч"
                Case "Щ"
                    newstr = newstr + "Ш"
                Case "Ъ"
                    newstr = newstr + "Щ"
                Case "Ы"
                    newstr = newstr + "Ъ"
                Case "Ь"
                    newstr = newstr + "Ы"
                Case "Э"
                    newstr = newstr + "Ь"
                Case "Ю"
                    newstr = newstr + "Э"
                Case "Я"
                    newstr = newstr + "Ю"
                Case "а"
                    newstr = newstr + "Я"
                Case Else
                    newstr = newstr + charw
            End Select
        Next

        Return newstr
    End Function
    Private Declare Function cnc_allclibhndl Lib "FWLIB32.DLL" (ByRef FlibHndl As Short) As Short
    Private Declare Function cnc_sysinfo Lib "FWLIB32.DLL" (ByVal FlibHndl As Short, ByRef odb As ODBSYS) As Short
    ' read alarm history data 
    Declare Function cnc_rdalmhistry Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByVal a As Integer, ByVal b As Integer, ByVal c As Integer, ByRef d As ODBAHIS) As Short
    ' stop logging operation history data 
    Declare Function cnc_stopophis Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer) As Short
    ' restart logging operation history data 
    Declare Function cnc_startophis Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer) As Short
    ' read number of alarm history data 
    Declare Function cnc_rdalmhisno Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByRef a As Integer) As Short
    ' read timer data from cnc 
    Declare Function cnc_rdtimer Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByVal a As Short, ByRef b As IODBTIME) As Short
    ' read CNC status information 
    Declare Function cnc_statinfo Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByRef a As ODBST) As Short
    ' read operator message history 
    Declare Function cnc_rdomhistry Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByVal a As Short, ByRef b As Integer, ByRef c As ODBOMHIS) As Short
    ' stop the sampling for operator message history 
    Declare Function cnc_stopomhis Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer) As Short
    ' start the sampling for operator message history 
    Declare Function cnc_startomhis Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer) As Short
    ' read block counter 
    Declare Function cnc_rdblkcount Lib "FWLIB32.DLL" (ByVal FlibHndl As Integer, ByRef a As Integer) As Short
    Function _cnc_allclibhndl(ByRef err As Short) As Short
        Dim Hndl As Short

        err = cnc_allclibhndl(Hndl)
        If (err = 0) Then
            Return Hndl
        Else
            Return 0
        End If

    End Function
    Function _cnc_rdalmhistry(ByVal Hndl As Short, ByRef err As Short) As ODBAHIS
        Dim e_no, lenght As Integer,
            d As ODBAHIS

        cnc_stopophis(Hndl)
        cnc_rdalmhisno(Hndl, e_no)
        lenght = 54 * e_no

        err = cnc_rdalmhistry(Hndl, 2, e_no, lenght, d)
        cnc_startophis(Hndl)

        Return d
    End Function
    Function _cnc_statinfo(ByVal Hndl As Short, ByRef err As Short) As ODBST
        Dim DataStat As ODBST

        err = cnc_statinfo(Hndl, DataStat)

        Return DataStat
    End Function
    Function _cnc_rdblkcount() As Integer
        Dim Hndl As Short,
            count As Integer

        cnc_allclibhndl(Hndl)
        cnc_rdblkcount(Hndl, count)

        Return count
    End Function
    Function _cnc_cycletimer() As IODBTIME
        Dim Hndl As Short,
           iodb As IODBTIME

        cnc_allclibhndl(Hndl)
        cnc_rdtimer(Hndl, 3, iodb)

        Return iodb
    End Function
End Class
