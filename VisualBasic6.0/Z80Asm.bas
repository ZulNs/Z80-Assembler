Attribute VB_Name = "modMain"

'=========================================================================================='
'                                                                                          '
'              Z80 Assembler for ED-Laboratory's Microprocessor Trainer MPT-1              '
'                                                                                          '
'                Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005                '
'                                                                                          '
'=========================================================================================='


Option Explicit

Const cOpCodesTblFileName As String = "OPCODES.TBL"
Public Const cSrcFileExt As String = "ASM"
Const cOutFileExt As String = "Z80"
Const cBinFileExt As String = "BIN"
Const cErrFileExt As String = "ERR"
Const cLogFileExt As String = "LOG"
Const cErrFlag As String = "$Err"
Const cBinFileStartStr As String = "<Z80_Executable_Codes>"
Const cBinFileEndStr As String = _
    "<ZulNs#05-11-1970#Viva_New_Technology_Protocol#Gorontalo#Feb-2005>"
Const cDefStartAddr As Long = &H1800
Const cDir As Long = 1
Const cInst As Long = 2
Const cReg As Long = 4
Const cRefReg As Long = 8
Const cFlagId As Long = &H10
Const cNum As Long = &H20
Const cOvNum As Long = &H40
Const cBadNum As Long = &H80
Const cBadOvNum As Long = cBadNum Or cOvNum
Const cEmpty As Long = &H100
Const cUndef As Long = &H200
Const cConst As Long = &H400
Const cConstNum As Long = cConst Or cNum
Const cOpIdAbs As Long = 1
Const cOpIdRef As Long = 2
Const cOpIdDisRef As Long = 4

Dim DirsList, InstsList, RegsList, FlagIdsList, RefRegsList
Dim ConstNames() As String, ConstVals() As String, InstsTbl(695, 3) As String
Dim Lbls() As String, Insts() As String, Op1s() As String, Op2s() As String
Dim OpCodes() As String, Args() As String, ErrDesc As String, LblLen As Long
Dim SrcFile As String, PrgPath As String, OutFile As String

Sub Main()
    Dim blnSuccess As Boolean
'    If Not IsFileExist(cOpCodesTblFileName) Then
'        MsgBox "Can't found '" & cOpCodesTblFileName & _
'            "' file. Assembling process aborted.", vbCritical
'        Exit Sub
'    End If
    If Not GetSrcFileName Then Exit Sub
    InitVars
    blnSuccess = BuildMnemonicsList
    blnSuccess = ReplConsts And blnSuccess
    blnSuccess = SetAddrsVal And blnSuccess
    blnSuccess = ReplConsts And blnSuccess
    If blnSuccess Then WrtOutFile Else WrtErrFile
    Shell "Notepad.exe " & OutFile, vbNormalFocus
End Sub

Private Function GetSrcFileName() As Boolean
    Dim CmdTail As String, fso As Object, Path As String, UserRespons As VbMsgBoxResult
    PrgPath = CurDir
    CmdTail = Command()
    If CmdTail = "" Then
        If Not GetSrcFileNameFromDlg Then Exit Function
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(CmdTail) Then _
            If fso.GetExtensionName(CmdTail) = "" Then CmdTail = CmdTail & "." & cSrcFileExt
        If fso.FileExists(CmdTail) Then
            Path = fso.GetParentFolderName(CmdTail)
            If Path = "" Then
                SrcFile = CmdTail
            Else
                ChDrive fso.GetDriveName(Path)
                ChDir Path
                SrcFile = fso.GetFileName(CmdTail)
            End If
            Set fso = Nothing
        Else
            Set fso = Nothing
            UserRespons = MsgBox("Can't found '" & CmdTail & "' file or '" & _
                CmdTail & "' is not a legal file name." & vbCr & _
                "Try to find it or another file by your self?", vbQuestion + vbOKCancel)
            If UserRespons = vbOK Then
                If Not GetSrcFileNameFromDlg Then Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    Select Case UCase(GetFileExt(SrcFile))
    Case "", cErrFileExt, cOutFileExt, cLogFileExt, cBinFileExt
        OutFile = SrcFile
    Case Else
        OutFile = GetFileName(SrcFile)
    End Select
    DelFile OutFile & "." & cOutFileExt
    DelFile OutFile & "." & cErrFileExt
    DelFile OutFile & "." & cLogFileExt
    DelFile OutFile & "." & cBinFileExt
    GetSrcFileName = True
End Function

Private Function GetSrcFileNameFromDlg() As Boolean
    Dim dlg As New frmDlgFileOpen, blnExit As Boolean
    dlg.Show vbModal
    blnExit = dlg.ExitMode
    If blnExit Then SrcFile = dlg.FileName
    Unload dlg
    Set dlg = Nothing
    If blnExit Then GetSrcFileNameFromDlg = True _
    Else MsgBox "No file selected. Assembling process aborted.", vbInformation
End Function

Private Function IsFileExist(FileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsFileExist = fso.FileExists(FileName)
    Set fso = Nothing
End Function

Private Function DelFile(FileName As String)
    Dim fso As Object
    If IsFileExist(FileName) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile FileName, True
        Set fso = Nothing
    End If
End Function

Private Function InitVars()
    Dim I As Long, J As Long
    DirsList = Array("END", "EQU", "DEFB", "DEFS", "DEFW", "ORG")
    InstsList = Array("ADC", "ADD", "AND", "BIT", "CALL", "CCF", "CP", "CPD", "CPDR", _
        "CPI", "CPIR", "CPL", "DAA", "DEC", "DI", "DJNZ", "EI", "EX", "EXX", "HALT", "IM", _
        "IN", "INC", "IND", "INDR", "INI", "INIR", "JP", "JR", "LD", "LDD", "LDDR", "LDI", _
        "LDIR", "NEG", "NOP", "OR", "OTDR", "OTIR", "OUT", "OUTD", "OUTI", "POP", "PUSH", _
        "RES", "RET", "RETI", "RETN", "RL", "RLA", "RLC", "RLCA", "RLD", "RR", "RRA", _
        "RRC", "RRCA", "RRD", "RST", "SBC", "SCF", "SET", "SLA", "SRA", "SRL", "SUB", _
        "XOR")
    ' ZulNs: RegsList = Array("A", "B", "C", "D", "E", "H", "L", "AF", "AF'", "BC", "DE", "HL", _
    ' ZulNs:    "IX", "IY", "SP")
    RegsList = Array("A", "B", "C", "D", "E", "H", "L", "I", "R", "AF", "AF'", "BC", "DE", "HL", _
        "IX", "IY", "SP")
    RefRegsList = Array("C", "BC", "DE", "HL", "IX", "IY", "SP")
    FlagIdsList = Array("Z", "NZ", "C", "NC", "P", "M", "PO", "PE")
    DelFile PrgPath & "\" & cOpCodesTblFileName
    Open PrgPath & "\" & cOpCodesTblFileName For Output As #1
    Print #1, "DEFB,n8,,n8"
    Print #1, "DEFW,n16,,nLnH"
    Print #1, "ADC,A,(HL),8E"
    Print #1, "ADC,A,(IX+d7),DD8Ed7"
    Print #1, "ADC,A,(IY+d7),FD8Ed7"
    Print #1, "ADC,A,A,8F"
    Print #1, "ADC,A,B,88"
    Print #1, "ADC,A,C,89"
    Print #1, "ADC,A,D,8A"
    Print #1, "ADC,A,E,8B"
    Print #1, "ADC,A,H,8C"
    Print #1, "ADC,A,L,8D"
    Print #1, "ADC,A,n8,CEn8"
    Print #1, "ADC,HL,BC,ED4A"
    Print #1, "ADC,HL,DE,ED5A"
    Print #1, "ADC,HL,HL,ED6A"
    Print #1, "ADC,HL,SP,ED7A"
    Print #1, "ADD,A,(HL),86"
    Print #1, "ADD,A,(IX+d7),DD86d7"
    Print #1, "ADD,A,(IY+d7),FD86d7"
    Print #1, "ADD,A,A,87"
    Print #1, "ADD,A,B,80"
    Print #1, "ADD,A,C,81"
    Print #1, "ADD,A,D,82"
    Print #1, "ADD,A,E,83"
    Print #1, "ADD,A,H,84"
    Print #1, "ADD,A,L,85"
    Print #1, "ADD,A,n8,C6n8"
    Print #1, "ADD,HL,BC,09"
    Print #1, "ADD,HL,DE,19"
    Print #1, "ADD,HL,HL,29"
    Print #1, "ADD,HL,SP,39"
    Print #1, "ADD,IX,BC,DD09"
    Print #1, "ADD,IX,DE,DD19"
    Print #1, "ADD,IX,IX,DD29"
    Print #1, "ADD,IX,SP,DD39"
    Print #1, "ADD,IY,BC,FD09"
    Print #1, "ADD,IY,DE,FD19"
    Print #1, "ADD,IY,IY,FD29"
    Print #1, "ADD,IY,SP,FD39"
    Print #1, "AND,(HL),,A6"
    Print #1, "AND,(IX+d7),,DDA6d7"
    Print #1, "AND,(IY+d7),,FDA6d7"
    Print #1, "AND,A,,A7"
    Print #1, "AND,B,,A0"
    Print #1, "AND,C,,A1"
    Print #1, "AND,D,,A2"
    Print #1, "AND,E,,A3"
    Print #1, "AND,H,,A4"
    Print #1, "AND,L,,A5"
    Print #1, "AND,n8,,E6n8"
    Print #1, "BIT,0,(HL),CB46"
    Print #1, "BIT,0,(IX+d7),DDCBd746"
    Print #1, "BIT,0,(IY+d7),FDCBd746"
    Print #1, "BIT,0,A,CB47"
    Print #1, "BIT,0,B,CB40"
    Print #1, "BIT,0,C,CB41"
    Print #1, "BIT,0,D,CB42"
    Print #1, "BIT,0,E,CB43"
    Print #1, "BIT,0,H,CB44"
    Print #1, "BIT,0,L,CB45"
    Print #1, "BIT,1,(HL),CB4E"
    Print #1, "BIT,1,(IX+d7),DDCBd74E"
    Print #1, "BIT,1,(IY+d7),FDCBd74E"
    Print #1, "BIT,1,A,CB4F"
    Print #1, "BIT,1,B,CB48"
    Print #1, "BIT,1,C,CB49"
    Print #1, "BIT,1,D,CB4A"
    Print #1, "BIT,1,E,CB4B"
    Print #1, "BIT,1,H,CB4C"
    Print #1, "BIT,1,L,CB4D"
    Print #1, "BIT,2,(HL),CB56"
    Print #1, "BIT,2,(IX+d7),DDCBd756"
    Print #1, "BIT,2,(IY+d7),FDCBd756"
    Print #1, "BIT,2,A,CB57"
    Print #1, "BIT,2,B,CB50"
    Print #1, "BIT,2,C,CB51"
    Print #1, "BIT,2,D,CB52"
    Print #1, "BIT,2,E,CB53"
    Print #1, "BIT,2,H,CB54"
    Print #1, "BIT,2,L,CB55"
    Print #1, "BIT,3,(HL),CB5E"
    Print #1, "BIT,3,(IX+d7),DDCBd75E"
    Print #1, "BIT,3,(IY+d7),FDCBd75E"
    Print #1, "BIT,3,A,CB5F"
    Print #1, "BIT,3,B,CB58"
    Print #1, "BIT,3,C,CB59"
    Print #1, "BIT,3,D,CB5A"
    Print #1, "BIT,3,E,CB5B"
    Print #1, "BIT,3,H,CB5C"
    Print #1, "BIT,3,L,CB5D"
    Print #1, "BIT,4,(HL),CB66"
    Print #1, "BIT,4,(IX+d7),DDCBd766"
    Print #1, "BIT,4,(IY+d7),FDCBd766"
    Print #1, "BIT,4,A,CB67"
    Print #1, "BIT,4,B,CB60"
    Print #1, "BIT,4,C,CB61"
    Print #1, "BIT,4,D,CB62"
    Print #1, "BIT,4,E,CB63"
    Print #1, "BIT,4,H,CB64"
    Print #1, "BIT,4,L,CB65"
    Print #1, "BIT,5,(HL),CB6E"
    Print #1, "BIT,5,(IX+d7),DDCBd76E"
    Print #1, "BIT,5,(IY+d7),FDCBd76E"
    Print #1, "BIT,5,A,CB6F"
    Print #1, "BIT,5,B,CB68"
    Print #1, "BIT,5,C,CB69"
    Print #1, "BIT,5,D,CB6A"
    Print #1, "BIT,5,E,CB6B"
    Print #1, "BIT,5,H,CB6C"
    Print #1, "BIT,5,L,CB6D"
    Print #1, "BIT,6,(HL),CB76"
    Print #1, "BIT,6,(IX+d7),DDCBd776"
    Print #1, "BIT,6,(IY+d7),FDCBd776"
    Print #1, "BIT,6,A,CB77"
    Print #1, "BIT,6,B,CB70"
    Print #1, "BIT,6,C,CB71"
    Print #1, "BIT,6,D,CB72"
    Print #1, "BIT,6,E,CB73"
    Print #1, "BIT,6,H,CB74"
    Print #1, "BIT,6,L,CB75"
    Print #1, "BIT,7,(HL),CB7E"
    Print #1, "BIT,7,(IX+d7),DDCBd77E"
    Print #1, "BIT,7,(IY+d7),FDCBd77E"
    Print #1, "BIT,7,A,CB7F"
    Print #1, "BIT,7,B,CB78"
    Print #1, "BIT,7,C,CB79"
    Print #1, "BIT,7,D,CB7A"
    Print #1, "BIT,7,E,CB7B"
    Print #1, "BIT,7,H,CB7C"
    Print #1, "BIT,7,L,CB7D"
    Print #1, "CALL,C,n16,DCnLnH"
    Print #1, "CALL,M,n16,FCnLnH"
    Print #1, "CALL,n16,,CDnLnH"
    Print #1, "CALL,NC,n16,D4nLnH"
    Print #1, "CALL,NZ,n16,C4nLnH"
    Print #1, "CALL,P,n16,F4nLnH"
    Print #1, "CALL,PE,n16,ECnLnH"
    Print #1, "CALL,PO,n16,E4nLnH"
    Print #1, "CALL,Z,n16,CCnLnH"
    Print #1, "CCF,,,3F"
    Print #1, "CP,(HL),,BE"
    Print #1, "CP,(IX+d7),,DDBEd7"
    Print #1, "CP,(IY+d7),,FDBEd7"
    Print #1, "CP,A,,BF"
    Print #1, "CP,B,,B8"
    Print #1, "CP,C,,B9"
    Print #1, "CP,D,,BA"
    Print #1, "CP,E,,BB"
    Print #1, "CP,H,,BC"
    Print #1, "CP,L,,BD"
    Print #1, "CP,n8,,FEn8"
    Print #1, "CPD,,,EDA9"
    Print #1, "CPDR,,,EDB9"
    Print #1, "CPI,,,EDA1"
    Print #1, "CPIR,,,EDB1"
    Print #1, "CPL,,,2F"
    Print #1, "DAA,,,27"
    Print #1, "DEC,(HL),,35"
    Print #1, "DEC,(IX+d7),,DD35d7"
    Print #1, "DEC,(IY+d7),,FD35d7"
    Print #1, "DEC,A,,3D"
    Print #1, "DEC,B,,05"
    Print #1, "DEC,BC,,0B"
    Print #1, "DEC,C,,0D"
    Print #1, "DEC,D,,15"
    Print #1, "DEC,DE,,1B"
    Print #1, "DEC,E,,1D"
    Print #1, "DEC,H,,25"
    Print #1, "DEC,HL,,2B"
    Print #1, "DEC,IX,,DD2B"
    Print #1, "DEC,IY,,FD2B"
    Print #1, "DEC,L,,2D"
    Print #1, "DEC,SP,,3B"
    Print #1, "DI,,,F3"
    Print #1, "DJNZ,d7,,10d7"
    Print #1, "EI,,,FB"
    Print #1, "EX,(SP),HL,E3"
    Print #1, "EX,(SP),IX,DDE3"
    Print #1, "EX,(SP),IY,FDE3"
    Print #1, "EX,AF,AF',08"
    Print #1, "EX,DE,HL,EB"
    Print #1, "EXX,,,D9"
    Print #1, "HALT,,,76"
    Print #1, "IM,0,,ED46"
    Print #1, "IM,1,,ED56"
    Print #1, "IM,2,,ED5E"
    Print #1, "IN,A,(C),ED78"
    Print #1, "IN,A,(n8),DBn8"
    Print #1, "IN,B,(C),ED40"
    Print #1, "IN,C,(C),ED48"
    Print #1, "IN,D,(C),ED50"
    Print #1, "IN,E,(C),ED58"
    Print #1, "IN,H,(C),ED60"
    Print #1, "IN,L,(C),ED68"
    Print #1, "INC,(HL),,34"
    Print #1, "INC,(IX+d7),,DD34d7"
    Print #1, "INC,(IY+d7),,FD34d7"
    Print #1, "INC,A,,3C"
    Print #1, "INC,B,,04"
    Print #1, "INC,BC,,03"
    Print #1, "INC,C,,0C"
    Print #1, "INC,D,,14"
    Print #1, "INC,DE,,13"
    Print #1, "INC,E,,1C"
    Print #1, "INC,H,,24"
    Print #1, "INC,HL,,23"
    Print #1, "INC,IX,,DD23"
    Print #1, "INC,IY,,FD23"
    Print #1, "INC,L,,2C"
    Print #1, "INC,SP,,33"
    Print #1, "IND,,,EDAA"
    Print #1, "INDR,,,EDBA"
    Print #1, "INI,,,EDA2"
    Print #1, "INIR,,,EDB2"
    Print #1, "JP,(HL),,E9"
    Print #1, "JP,(IX),,DDE9"
    Print #1, "JP,(IY),,FDE9"
    Print #1, "JP,C,n16,DAnLnH"
    Print #1, "JP,M,n16,FAnLnH"
    Print #1, "JP,n16,,C3nLnH"
    Print #1, "JP,NC,n16,D2nLnH"
    Print #1, "JP,NZ,n16,C2nLnH"
    Print #1, "JP,P,n16,F2nLnH"
    Print #1, "JP,PE,n16,EAnLnH"
    Print #1, "JP,PO,n16,E2nLnH"
    Print #1, "JP,Z,n16,CAnLnH"
    Print #1, "JR,C,d7,38d7"
    Print #1, "JR,d7,,18d7"
    Print #1, "JR,NC,d7,30d7"
    Print #1, "JR,NZ,d7,20d7"
    Print #1, "JR,Z,d7,28d7"
    Print #1, "LD,(BC),A,02"
    Print #1, "LD,(DE),A,12"
    Print #1, "LD,(HL),A,77"
    Print #1, "LD,(HL),B,70"
    Print #1, "LD,(HL),C,71"
    Print #1, "LD,(HL),D,72"
    Print #1, "LD,(HL),E,73"
    Print #1, "LD,(HL),H,74"
    Print #1, "LD,(HL),L,75"
    Print #1, "LD,(HL),n8,36n8"
    Print #1, "LD,(IX+d7),A,DD77d7"
    Print #1, "LD,(IX+d7),B,DD70d7"
    Print #1, "LD,(IX+d7),C,DD71d7"
    Print #1, "LD,(IX+d7),D,DD72d7"
    Print #1, "LD,(IX+d7),E,DD73d7"
    Print #1, "LD,(IX+d7),H,DD74d7"
    Print #1, "LD,(IX+d7),L,DD75d7"
    Print #1, "LD,(IX+d7),n8,DD36d7n8"
    Print #1, "LD,(IY+d7),A,FD77d7"
    Print #1, "LD,(IY+d7),B,FD70d7"
    Print #1, "LD,(IY+d7),C,FD71d7"
    Print #1, "LD,(IY+d7),D,FD72d7"
    Print #1, "LD,(IY+d7),E,FD73d7"
    Print #1, "LD,(IY+d7),H,FD74d7"
    Print #1, "LD,(IY+d7),L,FD75d7"
    Print #1, "LD,(IY+d7),n8,FD36d7n8"
    Print #1, "LD,(n16),A,32nLnH"
    Print #1, "LD,(n16),BC,ED43nLnH"
    Print #1, "LD,(n16),DE,ED53nLnH"
    Print #1, "LD,(n16),HL,22nLnH"
    Print #1, "LD,(n16),IX,DD22nLnH"
    Print #1, "LD,(n16),IY,FD22nLnH"
    Print #1, "LD,(n16),SP,ED73nLnH"
    Print #1, "LD,A,(BC),0A"
    Print #1, "LD,A,(DE),1A"
    Print #1, "LD,A,(HL),7E"
    Print #1, "LD,A,(IX+d7),DD7Ed7"
    Print #1, "LD,A,(IY+d7),FD7Ed7"
    Print #1, "LD,A,(n16),3AnLnH"
    Print #1, "LD,A,A,7F"
    Print #1, "LD,A,B,78"
    Print #1, "LD,A,C,79"
    Print #1, "LD,A,D,7A"
    Print #1, "LD,A,E,7B"
    Print #1, "LD,A,H,7C"
    Print #1, "LD,A,I,ED57"
    Print #1, "LD,A,L,7D"
    Print #1, "LD,A,R,ED5F"
    Print #1, "LD,A,n8,3En8"
    Print #1, "LD,B,(HL),46"
    Print #1, "LD,B,(IX+d7),DD46d7"
    Print #1, "LD,B,(IY+d7),FD46d7"
    Print #1, "LD,B,A,47"
    Print #1, "LD,B,B,40"
    Print #1, "LD,B,C,41"
    Print #1, "LD,B,D,42"
    Print #1, "LD,B,E,43"
    Print #1, "LD,B,H,44"
    Print #1, "LD,B,L,45"
    Print #1, "LD,B,n8,06n8"
    Print #1, "LD,BC,(n16),ED4BnLnH"
    Print #1, "LD,BC,n16,01nLnH"
    Print #1, "LD,C,(HL),4E"
    Print #1, "LD,C,(IX+d7),DD4Ed7"
    Print #1, "LD,C,(IY+d7),FD4Ed7"
    Print #1, "LD,C,A,4F"
    Print #1, "LD,C,B,48"
    Print #1, "LD,C,C,49"
    Print #1, "LD,C,D,4A"
    Print #1, "LD,C,E,4B"
    Print #1, "LD,C,H,4C"
    Print #1, "LD,C,L,4D"
    Print #1, "LD,C,n8,0En8"
    Print #1, "LD,D,(HL),56"
    Print #1, "LD,D,(IX+d7),DD56d7"
    Print #1, "LD,D,(IY+d7),FD56d7"
    Print #1, "LD,D,A,57"
    Print #1, "LD,D,B,50"
    Print #1, "LD,D,C,51"
    Print #1, "LD,D,D,52"
    Print #1, "LD,D,E,53"
    Print #1, "LD,D,H,54"
    Print #1, "LD,D,L,55"
    Print #1, "LD,D,n8,16n8"
    Print #1, "LD,DE,(n16),ED5BnLnH"
    Print #1, "LD,DE,n16,11nLnH"
    Print #1, "LD,E,(HL),5E"
    Print #1, "LD,E,(IX+d7),DD5Ed7"
    Print #1, "LD,E,(IY+d7),FD5Ed7"
    Print #1, "LD,E,A,5F"
    Print #1, "LD,E,B,58"
    Print #1, "LD,E,C,59"
    Print #1, "LD,E,D,5A"
    Print #1, "LD,E,E,5B"
    Print #1, "LD,E,H,5C"
    Print #1, "LD,E,L,5D"
    Print #1, "LD,E,n8,1En8"
    Print #1, "LD,H,(HL),66"
    Print #1, "LD,H,(IX+d7),DD66d7"
    Print #1, "LD,H,(IY+d7),FD66d7"
    Print #1, "LD,H,A,67"
    Print #1, "LD,H,B,60"
    Print #1, "LD,H,C,61"
    Print #1, "LD,H,D,62"
    Print #1, "LD,H,E,63"
    Print #1, "LD,H,H,64"
    Print #1, "LD,H,L,65"
    Print #1, "LD,H,n8,26n8"
    Print #1, "LD,HL,(n16),2AnLnH"
    Print #1, "LD,HL,n16,21nLnH"
    Print #1, "LD,I,A,ED47"
    Print #1, "LD,IX,(n16),DD2AnLnH"
    Print #1, "LD,IX,n16,DD21nLnH"
    Print #1, "LD,IY,(n16),FD2AnLnH"
    Print #1, "LD,IY,n16,FD21nLnH"
    Print #1, "LD,L,(HL),6E"
    Print #1, "LD,L,(IX+d7),DD6Ed7"
    Print #1, "LD,L,(IY+d7),FD6Ed7"
    Print #1, "LD,L,A,6F"
    Print #1, "LD,L,B,68"
    Print #1, "LD,L,C,69"
    Print #1, "LD,L,D,6A"
    Print #1, "LD,L,E,6B"
    Print #1, "LD,L,H,6C"
    Print #1, "LD,L,L,6D"
    Print #1, "LD,L,n8,2En8"
    Print #1, "LD,R,A,ED4F"
    Print #1, "LD,SP,(n16),ED7BnLnH"
    Print #1, "LD,SP,HL,F9"
    Print #1, "LD,SP,IX,DDF9"
    Print #1, "LD,SP,IY,FDF9"
    Print #1, "LD,SP,n16,31nLnH"
    Print #1, "LDD,,,EDA8"
    Print #1, "LDDR,,,EDB8"
    Print #1, "LDI,,,EDA0"
    Print #1, "LDIR,,,EDB0"
    Print #1, "NEG,,,ED44"
    Print #1, "NOP,,,00"
    Print #1, "OR,(HL),,B6"
    Print #1, "OR,(IX+d7),,DDB6d7"
    Print #1, "OR,(IY+d7),,FDB6d7"
    Print #1, "OR,A,,B7"
    Print #1, "OR,B,,B0"
    Print #1, "OR,C,,B1"
    Print #1, "OR,D,,B2"
    Print #1, "OR,E,,B3"
    Print #1, "OR,H,,B4"
    Print #1, "OR,L,,B5"
    Print #1, "OR,n8,,F6n8"
    Print #1, "OTDR,,,EDBB"
    Print #1, "OTIR,,,EDB3"
    Print #1, "OUT,(C),A,ED79"
    Print #1, "OUT,(C),B,ED41"
    Print #1, "OUT,(C),C,ED49"
    Print #1, "OUT,(C),D,ED51"
    Print #1, "OUT,(C),E,ED59"
    Print #1, "OUT,(C),H,ED61"
    Print #1, "OUT,(C),L,ED69"
    Print #1, "OUT,(n8),A,D3n8"
    Print #1, "OUTD,,,EDAB"
    Print #1, "OUTI,,,EDA3"
    Print #1, "POP,AF,,F1"
    Print #1, "POP,BC,,C1"
    Print #1, "POP,DE,,D1"
    Print #1, "POP,HL,,E1"
    Print #1, "POP,IX,,DDE1"
    Print #1, "POP,IY,,FDE1"
    Print #1, "PUSH,AF,,F5"
    Print #1, "PUSH,BC,,C5"
    Print #1, "PUSH,DE,,D5"
    Print #1, "PUSH,HL,,E5"
    Print #1, "PUSH,IX,,DDE5"
    Print #1, "PUSH,IY,,FDE5"
    Print #1, "RES,0,(HL),CB86"
    Print #1, "RES,0,(IX+d7),DDCBd786"
    Print #1, "RES,0,(IY+d7),FDCBd786"
    Print #1, "RES,0,A,CB87"
    Print #1, "RES,0,B,CB80"
    Print #1, "RES,0,C,CB81"
    Print #1, "RES,0,D,CB82"
    Print #1, "RES,0,E,CB83"
    Print #1, "RES,0,H,CB84"
    Print #1, "RES,0,L,CB85"
    Print #1, "RES,1,(HL),CB8E"
    Print #1, "RES,1,(IX+d7),DDCBd78E"
    Print #1, "RES,1,(IY+d7),FDCBd78E"
    Print #1, "RES,1,A,CB8F"
    Print #1, "RES,1,B,CB88"
    Print #1, "RES,1,C,CB89"
    Print #1, "RES,1,D,CB8A"
    Print #1, "RES,1,E,CB8B"
    Print #1, "RES,1,H,CB8C"
    Print #1, "RES,1,L,CB8D"
    Print #1, "RES,2,(HL),CB96"
    Print #1, "RES,2,(IX+d7),DDCBd796"
    Print #1, "RES,2,(IY+d7),FDCBd796"
    Print #1, "RES,2,A,CB97"
    Print #1, "RES,2,B,CB90"
    Print #1, "RES,2,C,CB91"
    Print #1, "RES,2,D,CB92"
    Print #1, "RES,2,E,CB93"
    Print #1, "RES,2,H,CB94"
    Print #1, "RES,2,L,CB95"
    Print #1, "RES,3,(HL),CB9E"
    Print #1, "RES,3,(IX+d7),DDCBd79E"
    Print #1, "RES,3,(IY+d7),FDCBd79E"
    Print #1, "RES,3,A,CB9F"
    Print #1, "RES,3,B,CB98"
    Print #1, "RES,3,C,CB99"
    Print #1, "RES,3,D,CB9A"
    Print #1, "RES,3,E,CB9B"
    Print #1, "RES,3,H,CB9C"
    Print #1, "RES,3,L,CB9D"
    Print #1, "RES,4,(HL),CBA6"
    Print #1, "RES,4,(IX+d7),DDCBd7A6"
    Print #1, "RES,4,(IY+d7),FDCBd7A6"
    Print #1, "RES,4,A,CBA7"
    Print #1, "RES,4,B,CBA0"
    Print #1, "RES,4,C,CBA1"
    Print #1, "RES,4,D,CBA2"
    Print #1, "RES,4,E,CBA3"
    Print #1, "RES,4,H,CBA4"
    Print #1, "RES,4,L,CBA5"
    Print #1, "RES,5,(HL),CBAE"
    Print #1, "RES,5,(IX+d7),DDCBd7AE"
    Print #1, "RES,5,(IY+d7),FDCBd7AE"
    Print #1, "RES,5,A,CBAF"
    Print #1, "RES,5,B,CBA8"
    Print #1, "RES,5,C,CBA9"
    Print #1, "RES,5,D,CBAA"
    Print #1, "RES,5,E,CBAB"
    Print #1, "RES,5,H,CBAC"
    Print #1, "RES,5,L,CBAD"
    Print #1, "RES,6,(HL),CBB6"
    Print #1, "RES,6,(IX+d7),DDCBd7B6"
    Print #1, "RES,6,(IY+d7),FDCBd7B6"
    Print #1, "RES,6,A,CBB7"
    Print #1, "RES,6,B,CBB0"
    Print #1, "RES,6,C,CBB1"
    Print #1, "RES,6,D,CBB2"
    Print #1, "RES,6,E,CBB3"
    Print #1, "RES,6,H,CBB4"
    Print #1, "RES,6,L,CBB5"
    Print #1, "RES,7,(HL),CBBE"
    Print #1, "RES,7,(IX+d7),DDCBd7BE"
    Print #1, "RES,7,(IY+d7),FDCBd7BE"
    Print #1, "RES,7,A,CBBF"
    Print #1, "RES,7,B,CBB8"
    Print #1, "RES,7,C,CBB9"
    Print #1, "RES,7,D,CBBA"
    Print #1, "RES,7,E,CBBB"
    Print #1, "RES,7,H,CBBC"
    Print #1, "RES,7,L,CBBD"
    Print #1, "RET,C,,D8"
    Print #1, "RET,M,,F8"
    Print #1, "RET,NC,,D0"
    Print #1, "RET,NZ,,C0"
    Print #1, "RET,P,,F0"
    Print #1, "RET,PE,,E8"
    Print #1, "RET,PO,,E0"
    Print #1, "RET,Z,,C8"
    Print #1, "RET,,,C9"
    Print #1, "RETI,,,ED4D"
    Print #1, "RETN,,,ED45"
    Print #1, "RL,(HL),,CB16"
    Print #1, "RL,(IX+d7),,DDCBd716"
    Print #1, "RL,(IY+d7),,FDCBd716"
    Print #1, "RL,A,,CB17"
    Print #1, "RL,B,,CB10"
    Print #1, "RL,C,,CB11"
    Print #1, "RL,D,,CB12"
    Print #1, "RL,E,,CB13"
    Print #1, "RL,H,,CB14"
    Print #1, "RL,L,,CB15"
    Print #1, "RLA,,,17"
    Print #1, "RLC,(HL),,CB06"
    Print #1, "RLC,(IX+d7),,DDCBd706"
    Print #1, "RLC,(IY+d7),,FDCBd706"
    Print #1, "RLC,A,,CB07"
    Print #1, "RLC,B,,CB00"
    Print #1, "RLC,C,,CB01"
    Print #1, "RLC,D,,CB02"
    Print #1, "RLC,E,,CB03"
    Print #1, "RLC,H,,CB04"
    Print #1, "RLC,L,,CB05"
    Print #1, "RLCA,,,07"
    Print #1, "RLD,,,ED6F"
    Print #1, "RR,(HL),,CB1E"
    Print #1, "RR,(IX+d7),,DDCBd71E"
    Print #1, "RR,(IY+d7),,FDCBd71E"
    Print #1, "RR,A,,CB1F"
    Print #1, "RR,B,,CB18"
    Print #1, "RR,C,,CB19"
    Print #1, "RR,D,,CB1A"
    Print #1, "RR,E,,CB1B"
    Print #1, "RR,H,,CB1C"
    Print #1, "RR,L,,CB1D"
    Print #1, "RRA,,,1F"
    Print #1, "RRC,(HL),,CB0E"
    Print #1, "RRC,(IX+d7),,DDCBd70E"
    Print #1, "RRC,(IY+d7),,FDCBd70E"
    Print #1, "RRC,A,,CB0F"
    Print #1, "RRC,B,,CB08"
    Print #1, "RRC,C,,CB09"
    Print #1, "RRC,D,,CB0A"
    Print #1, "RRC,E,,CB0B"
    Print #1, "RRC,H,,CB0C"
    Print #1, "RRC,L,,CB0D"
    Print #1, "RRCA,,,0F"
    Print #1, "RRD,,,ED67"
    Print #1, "RST,0,,C7"
    Print #1, "RST,8,,CF"
    Print #1, "RST,16,,D7"
    Print #1, "RST,24,,DF"
    Print #1, "RST,32,,E7"
    Print #1, "RST,40,,EF"
    Print #1, "RST,48,,F7"
    Print #1, "RST,56,,FF"
    Print #1, "SBC,A,(HL),9E"
    Print #1, "SBC,A,(IX+d7),DD9Ed7"
    Print #1, "SBC,A,(IY+d7),FD9Ed7"
    Print #1, "SBC,A,A,9F"
    Print #1, "SBC,A,B,98"
    Print #1, "SBC,A,C,99"
    Print #1, "SBC,A,D,9A"
    Print #1, "SBC,A,E,9B"
    Print #1, "SBC,A,H,9C"
    Print #1, "SBC,A,L,9D"
    Print #1, "SBC,A,n8,DEn8"
    Print #1, "SBC,HL,BC,ED42"
    Print #1, "SBC,HL,DE,ED52"
    Print #1, "SBC,HL,HL,ED62"
    Print #1, "SBC,HL,SP,ED72"
    Print #1, "SCF,,,37"
    Print #1, "SET,0,(HL),CBC6"
    Print #1, "SET,0,(IX+d7),DDCBd7C6"
    Print #1, "SET,0,(IY+d7),FDCBd7C6"
    Print #1, "SET,0,A,CBC7"
    Print #1, "SET,0,B,CBC0"
    Print #1, "SET,0,C,CBC1"
    Print #1, "SET,0,D,CBC2"
    Print #1, "SET,0,E,CBC3"
    Print #1, "SET,0,H,CBC4"
    Print #1, "SET,0,L,CBC5"
    Print #1, "SET,1,(HL),CBCE"
    Print #1, "SET,1,(IX+d7),DDCBd7CE"
    Print #1, "SET,1,(IY+d7),FDCBd7CE"
    Print #1, "SET,1,A,CBCF"
    Print #1, "SET,1,B,CBC8"
    Print #1, "SET,1,C,CBC9"
    Print #1, "SET,1,D,CBCA"
    Print #1, "SET,1,E,CBCB"
    Print #1, "SET,1,H,CBCC"
    Print #1, "SET,1,L,CBCD"
    Print #1, "SET,2,(HL),CBD6"
    Print #1, "SET,2,(IX+d7),DDCBd7D6"
    Print #1, "SET,2,(IY+d7),FDCBd7D6"
    Print #1, "SET,2,A,CBD7"
    Print #1, "SET,2,B,CBD0"
    Print #1, "SET,2,C,CBD1"
    Print #1, "SET,2,D,CBD2"
    Print #1, "SET,2,E,CBD3"
    Print #1, "SET,2,H,CBD4"
    Print #1, "SET,2,L,CBD5"
    Print #1, "SET,3,(HL),CBDE"
    Print #1, "SET,3,(IX+d7),DDCBd7DE"
    Print #1, "SET,3,(IY+d7),FDCBd7DE"
    Print #1, "SET,3,A,CBDF"
    Print #1, "SET,3,B,CBD8"
    Print #1, "SET,3,C,CBD9"
    Print #1, "SET,3,D,CBDA"
    Print #1, "SET,3,E,CBDB"
    Print #1, "SET,3,H,CBDC"
    Print #1, "SET,3,L,CBDD"
    Print #1, "SET,4,(HL),CBE6"
    Print #1, "SET,4,(IX+d7),DDCBd7E6"
    Print #1, "SET,4,(IY+d7),FDCBd7E6"
    Print #1, "SET,4,A,CBE7"
    Print #1, "SET,4,B,CBE0"
    Print #1, "SET,4,C,CBE1"
    Print #1, "SET,4,D,CBE2"
    Print #1, "SET,4,E,CBE3"
    Print #1, "SET,4,H,CBE4"
    Print #1, "SET,4,L,CBE5"
    Print #1, "SET,5,(HL),CBEE"
    Print #1, "SET,5,(IX+d7),DDCBd7EE"
    Print #1, "SET,5,(IY+d7),FDCBd7EE"
    Print #1, "SET,5,A,CBEF"
    Print #1, "SET,5,B,CBE8"
    Print #1, "SET,5,C,CBE9"
    Print #1, "SET,5,D,CBEA"
    Print #1, "SET,5,E,CBEB"
    Print #1, "SET,5,H,CBEC"
    Print #1, "SET,5,L,CBED"
    Print #1, "SET,6,(HL),CBF6"
    Print #1, "SET,6,(IX+d7),DDCBd7F6"
    Print #1, "SET,6,(IY+d7),FDCBd7F6"
    Print #1, "SET,6,A,CBF7"
    Print #1, "SET,6,B,CBF0"
    Print #1, "SET,6,C,CBF1"
    Print #1, "SET,6,D,CBF2"
    Print #1, "SET,6,E,CBF3"
    Print #1, "SET,6,H,CBF4"
    Print #1, "SET,6,L,CBF5"
    Print #1, "SET,7,(HL),CBFE"
    Print #1, "SET,7,(IX+d7),DDCBd7FE"
    Print #1, "SET,7,(IY+d7),FDCBd7FE"
    Print #1, "SET,7,A,CBFF"
    Print #1, "SET,7,B,CBF8"
    Print #1, "SET,7,C,CBF9"
    Print #1, "SET,7,D,CBFA"
    Print #1, "SET,7,E,CBFB"
    Print #1, "SET,7,H,CBFC"
    Print #1, "SET,7,L,CBFD"
    Print #1, "SLA,(HL),,CB26"
    Print #1, "SLA,(IX+d7),,DDCBd726"
    Print #1, "SLA,(IY+d7),,FDCBd726"
    Print #1, "SLA,A,,CB27"
    Print #1, "SLA,B,,CB20"
    Print #1, "SLA,C,,CB21"
    Print #1, "SLA,D,,CB22"
    Print #1, "SLA,E,,CB23"
    Print #1, "SLA,H,,CB24"
    Print #1, "SLA,L,,CB25"
    Print #1, "SRA,(HL),,CB2E"
    Print #1, "SRA,(IX+d7),,DDCBd72E"
    Print #1, "SRA,(IY+d7),,FDCBd72E"
    Print #1, "SRA,A,,CB2F"
    Print #1, "SRA,B,,CB28"
    Print #1, "SRA,C,,CB29"
    Print #1, "SRA,D,,CB2A"
    Print #1, "SRA,E,,CB2B"
    Print #1, "SRA,H,,CB2C"
    Print #1, "SRA,L,,CB2D"
    Print #1, "SRL,(HL),,CB3E"
    Print #1, "SRL,(IX+d7),,DDCBd73E"
    Print #1, "SRL,(IY+d7),,FDCBd73E"
    Print #1, "SRL,A,,CB3F"
    Print #1, "SRL,B,,CB38"
    Print #1, "SRL,C,,CB39"
    Print #1, "SRL,D,,CB3A"
    Print #1, "SRL,E,,CB3B"
    Print #1, "SRL,H,,CB3C"
    Print #1, "SRL,L,,CB3D"
    Print #1, "SUB,(HL),,96"
    Print #1, "SUB,(IX+d7),,DD96d7"
    Print #1, "SUB,(IY+d7),,FD96d7"
    Print #1, "SUB,A,,97"
    Print #1, "SUB,B,,90"
    Print #1, "SUB,C,,91"
    Print #1, "SUB,D,,92"
    Print #1, "SUB,E,,93"
    Print #1, "SUB,H,,94"
    Print #1, "SUB,L,,95"
    Print #1, "SUB,n8,,D6n8"
    Print #1, "XOR,(HL),,AE"
    Print #1, "XOR,(IX+d7),,DDAEd7"
    Print #1, "XOR,(IY+d7),,FDAEd7"
    Print #1, "XOR,A,,AF"
    Print #1, "XOR,B,,A8"
    Print #1, "XOR,C,,A9"
    Print #1, "XOR,D,,AA"
    Print #1, "XOR,E,,AB"
    Print #1, "XOR,H,,AC"
    Print #1, "XOR,L,,AD"
    Print #1, "XOR,n8,,EEn8"
    Close #1
    Open PrgPath & "\" & cOpCodesTblFileName For Input As #1
    For I = 0 To UBound(InstsTbl)
        For J = 0 To UBound(InstsTbl, 2)
            Input #1, InstsTbl(I, J)
        Next
    Next
    Close #1
    DelFile PrgPath & "\" & cOpCodesTblFileName
    ReDim ConstNames(0), ConstVals(0)
    ReDim Lbls(0), Insts(0), Op1s(0), Op2s(0)
    LblLen = 5
End Function

Private Function WrtOutFile() As Boolean
    Dim PC As Long, I As Long, J As Long, Addrs() As String
    Dim LogFlag As Boolean
    WrtOutFile = True
    I = UBound(Insts)
    ReDim OpCodes(I), Addrs(I)
    LogFlag = True
    PC = cDefStartAddr
    For I = 1 To I
        Select Case Insts(I)
        Case ""
            Addrs(I) = FixHexVal(PC, 4)
        Case "END", "EQU"
        Case "ORG"
            PC = Val(Op1s(I))
        Case Else
            If Insts(I) <> "DEFS" Then LogFlag = False
            OpCodes(I) = GetInstCode(I, PC)
            If Insts(I) = cErrFlag Then WrtOutFile = False
            Addrs(I) = FixHexVal(PC, 4)
            IncPrgCtr PC, Len(OpCodes(I)) / 2
            DecToHex Op1s(I)
            DecToHex Op2s(I)
        End Select
    Next
    If Not WrtOutFile Then
        WrtErrFile
        Exit Function
    End If
    If LogFlag Then
        WrtLogFile
        Exit Function
    End If
    WrtBinFile
    OutFile = OutFile & "." & cOutFileExt
    Open OutFile For Output As #1
    Print #1, GetHorLine(LblLen + 53)
    Print #1, "ADDRESS MACHINE-CODE  #   LABEL:"; Tab(LblLen + 32); "OPCODE  OPERAND"
    Print #1, GetHorLine(LblLen + 53)
    Print #1,
    For I = 1 To UBound(Insts)
        Select Case Insts(I)
        Case "END"
            Print #1, Tab(LblLen + 22); "*****     END     *****"
        Case "EQU"
        Case "ORG"
            Print #1,
        Case Else
            Print #1, Addrs(I); ":  ";
            If Insts(I) <> "" And Insts(I) <> "DEFS" Then
                J = 1
                Do
                    Print #1, " "; Mid(OpCodes(I), J, 2);
                    J = J + 2
                Loop While Mid(OpCodes(I), J, 2) <> ""
            End If
            Print #1, Tab(23); "#"; Tab(27); GetCmdLine(I)
        End Select
    Next
    Close #1
    MsgBox "Assembling process successful.", vbInformation
End Function

Private Function WrtBinFile() As Boolean
    Dim I As Long, J As Long, BinPtr As Long, LenPtr As Long, By As Byte
    Dim lngTmp As Long, strTmp As String
    WrtBinFile = True
    Open OutFile & "." & cBinFileExt For Binary Access Write As #1
    Put #1, 1, cBinFileStartStr
    LenPtr = Len(cBinFileStartStr) + 1
    BinPtr = LenPtr + 4
    strTmp = FixHexVal(cDefStartAddr, 4)
    By = Val("&H" & Right(strTmp, 2))
    Put #1, BinPtr - 2, By
    By = Val("&H" & Left(strTmp, 2))
    Put #1, BinPtr - 1, By
    For I = 1 To UBound(Insts)
        Select Case Insts(I)
        Case "", "EQU"
        Case "END"
            Exit For
        Case "ORG"
            strTmp = FixHexVal(Val(Op1s(I)), 4)
            If BinPtr = LenPtr + 4 Then
                By = Val("&H" & Right(strTmp, 2))
                Put #1, BinPtr - 2, By
                By = Val("&H" & Left(strTmp, 2))
                Put #1, BinPtr - 1, By
            Else
                lngTmp = BinPtr - LenPtr - 4
                If lngTmp > 65535 Then
                    WrtBinFile = False
                    Close #1
                    DelFile OutFile & "." & cBinFileExt
                    MsgBox "Can't created '" & OutFile & "." & cBinFileExt & _
                        "' file for more than 64kB instructions.", vbCritical
                    Exit Function
                End If
                By = Val("&H" & Right(strTmp, 2))
                Put #1, BinPtr + 2, By
                By = Val("&H" & Left(strTmp, 2))
                Put #1, BinPtr + 3, By
                strTmp = FixHexVal(lngTmp, 4)
                By = Val("&H" & Right(strTmp, 2))
                Put #1, LenPtr, By
                By = Val("&H" & Left(strTmp, 2))
                Put #1, LenPtr + 1, By
                LenPtr = BinPtr
                BinPtr = BinPtr + 4
            End If
        Case Else
            For J = 1 To Len(OpCodes(I)) Step 2
                By = Val("&H" & Mid(OpCodes(I), J, 2))
                Put #1, BinPtr, By
                BinPtr = BinPtr + 1
            Next
        End Select
    Next
    lngTmp = BinPtr - LenPtr - 4
    If lngTmp > 65535 Then
        WrtBinFile = False
        Close #1
        DelFile OutFile & "." & cBinFileExt
        MsgBox "Can't created '" & OutFile & "." & cBinFileExt & _
            "' file for more than 64kB instructions.", vbCritical
        Exit Function
    End If
    If lngTmp > 0 Then
        strTmp = FixHexVal(lngTmp, 4)
        By = Val("&H" & Right(strTmp, 2))
        Put #1, LenPtr, By
        By = Val("&H" & Left(strTmp, 2))
        Put #1, LenPtr + 1, By
    Else
        BinPtr = BinPtr - 4
    End If
    By = 0
    Put #1, BinPtr, By
    Put #1, BinPtr + 1, By
    Put #1, BinPtr + 2, cBinFileEndStr
    Close #1
    MsgBox "Executable file (" & OutFile & "." & cBinFileExt & _
        ") successful created.", vbInformation
End Function

Private Function WrtErrFile()
    Dim I As Long, ErrCtr As Long, CmdLine As String
    OutFile = OutFile & "." & cErrFileExt
    Open OutFile For Output As #1
    Print #1, GetHorLine(LblLen + 53)
    Print #1, "LABEL:"; Tab(LblLen + 6); "OPCODE  OPERAND           ; ERROR-DESCRIPTION"
    Print #1, GetHorLine(LblLen + 53)
    Print #1,
    For I = 1 To UBound(Insts)
        CmdLine = GetCmdLine(I)
        If Insts(I) = cErrFlag Then
            ErrCtr = ErrCtr + 1
            CmdLine = CmdLine & " (" & Mid(Str(ErrCtr), 2) & ")"
        End If
        Print #1, CmdLine
    Next
    Close #1
    MsgBox Mid(Str(ErrCtr), 2) & " Error(s) found!", vbCritical
End Function

Private Function WrtLogFile()
    OutFile = OutFile & "." & cLogFileExt
    Open OutFile For Output As #1
    Print #1, ";"; GetHorLine(78); ";"
    Print #1, ";"; Tab(80); ";"
    Print #1, ";        Z80 Assembler for ED-Laboratory's Microprocessor Trainer MPT-1"; Tab(80); ";"
    Print #1, ";"; Tab(80); ";"
    Print #1, ";                         Viva New Technology Protocol"; Tab(80); ";"
    Print #1, ";"; Tab(80); ";"
    Print #1, ";          Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005"; Tab(80); ";"
    Print #1, ";"; Tab(80); ";"
    Print #1, ";"; GetHorLine(78); ";"
    Print #1,
    Print #1,
    Print #1, "; There is no instruction to assembling..."
    Close #1
End Function

Private Function GetHorLine(ChrNum As Long) As String
    For ChrNum = 1 To ChrNum
        GetHorLine = GetHorLine & "="
    Next
End Function

Private Function FixHexVal(Value As Long, Optional FixVal As Long = 2) As String
    FixHexVal = Hex(Value)
    Do While Len(FixHexVal) < FixVal
        FixHexVal = "0" & FixHexVal
    Loop
End Function

Private Function DecToHex(Op As String)
    Dim D As String, H As String
    D = ExtractOp(Op)
    If Val(D) > 0 Then
        H = Hex(Val(D))
        If Val(D) > 9 Then H = H & "H"
        If Left(H, 1) > "9" Then H = "0" & H
        CompressOp Op, H
    End If
End Function

Private Function SetAddrsVal() As Boolean
    Dim PC As Long, InstLen As Long, I As Long
    SetAddrsVal = True
    PC = cDefStartAddr
    For I = 1 To UBound(Insts)
        Select Case Insts(I)
        Case "END", "EQU", cErrFlag
        Case "ORG"
            PC = Val(Op1s(I))
        Case Else
            If Lbls(I) <> "" Then SetConstVal Lbls(I), Mid(Str(PC), 2)
            If Insts(I) = cErrFlag Then SetAddrsVal = False
            IncPrgCtr PC, GetInstLen(I)
        End Select
    Next
End Function

Private Function GetInstLen(LinesPtr As Long) As Long
    Dim Ptr As Long
    Select Case Insts(LinesPtr)
    Case "", "END", "EQU", "ORG", cErrFlag
    Case "DEFS"
        GetInstLen = Val(Op1s(LinesPtr))
    Case Else
        If FindInstFromTbl(LinesPtr, Ptr) Then GetInstLen = Len(InstsTbl(Ptr, 3)) / 2
    End Select
End Function

Private Function GetInstCode(LinesPtr As Long, PrgCtr As Long) As String
    Dim Ptr As Long, OpIdFlag As Boolean, Op As String, Num As String
    Select Case Insts(LinesPtr)
    Case "", "END", "EQU", "ORG", cErrFlag
    Case "DEFS"
        For Ptr = 1 To Val(Op1s(LinesPtr))
            GetInstCode = GetInstCode & "FF"
        Next
    Case Else
        If FindInstFromTbl(LinesPtr, Ptr) Then GetInstCode = InstsTbl(Ptr, 3)
        Do
            If OpIdFlag Then Op = Op2s(LinesPtr) Else Op = Op1s(LinesPtr)
            Num = ExtractOp(Op)
            Select Case GetOpEqv(Op, LinesPtr)
            Case "n8", "(n8)"
                GetInstCode = Replace(GetInstCode, "n8", FixHexVal(Val(Num)))
            Case "n16", "(n16)"
                Num = FixHexVal(Val(Num), 4)
                GetInstCode = Replace(GetInstCode, "nH", Left(Num, 2))
                GetInstCode = Replace(GetInstCode, "nL", Right(Num, 2))
            Case "(IX+d7)", "(IY+d7)"
                If Mid(Op, 4, 1) = "-" Then Num = Str(256 - Val(Num))
                GetInstCode = Replace(GetInstCode, "d7", FixHexVal(Val(Num)))
            Case "d7"
                Num = Str(Val(Num) - PrgCtr - Len(GetInstCode) / 2)
                If Val(Num) < -128 Or Val(Num) > 127 Then
                    AddError GetCmdLine(LinesPtr), _
                        "Jump relative address is out of range", LinesPtr
                Else
                    If Val(Num) < 0 Then Num = Str(Val(Num) + 256)
                    GetInstCode = Replace(GetInstCode, "d7", _
                        FixHexVal(Val("&H" & Right(Hex(Val(Num)), 2))))
                End If
            End Select
            OpIdFlag = Not OpIdFlag
        Loop While OpIdFlag
    End Select
End Function

Private Function FindInstFromTbl(LinesPtr As Long, Ptr As Long) As Boolean
    For Ptr = 0 To UBound(InstsTbl)
        If Insts(LinesPtr) = InstsTbl(Ptr, 0) Then
            If GetOpEqv(Op1s(LinesPtr), LinesPtr) = InstsTbl(Ptr, 1) Then
                If GetOpEqv(Op2s(LinesPtr), LinesPtr) = InstsTbl(Ptr, 2) Then
                    FindInstFromTbl = True
                    Exit Function
                End If
            End If
        End If
    Next
    AddError GetCmdLine(LinesPtr), "Illegal instruction", LinesPtr
End Function

Private Function IncPrgCtr(LastVal As Long, IncVal As Long)
    LastVal = LastVal + IncVal
    If LastVal > 65536 Then LastVal = LastVal - 65536
End Function

Private Function GetOpEqv(Op As String, LinesPtr As Long) As String
    Select Case GetOpId(Op)
    Case cOpIdAbs
        If GetOpType(Op) And cConstNum Then
            Select Case Insts(LinesPtr)
            Case "CALL", "JP", "DEFW"
                GetOpEqv = "n16"
            Case "LD"
                If Op = Op2s(LinesPtr) And Len(Op1s(LinesPtr)) = 2 Then
                    GetOpEqv = "n16"
                Else
                    GetOpEqv = "n8"
                End If
            Case "ADC", "ADD", "AND", "CP", "OR", "SBC", "SUB", "XOR", "DEFB"
                GetOpEqv = "n8"
            Case "DJNZ", "JR"
                GetOpEqv = "d7"
            End Select
        End If
    Case cOpIdRef
        If GetOpType(ExtractOp(Op)) And cConstNum Then
            Select Case Insts(LinesPtr)
            Case "LD"
                GetOpEqv = "(n16)"
            Case "IN", "OUT"
                GetOpEqv = "(n8)"
            End Select
        End If
    Case cOpIdDisRef
        GetOpEqv = Left(Op, 3) & "+d7)"
    End Select
    If GetOpEqv = "" Then GetOpEqv = Op
End Function

Private Function ReplConsts() As Boolean
    Dim I As Long, OpFl As Boolean, Op As String, OkFl As Boolean, NumHold As String
    ReplConsts = True
    For I = 1 To UBound(Insts)
        Select Case Insts(I)
        Case "", "END", "EQU"
        Case cErrFlag
            ReplConsts = False
        Case Else
            If ReplChkConst(Op1s(I), I) Then
                If ReplChkConst(Op2s(I), I) Then GoTo NextReplConsts
            End If
            AddError GetCmdLine(I), ErrDesc, I
            ReplConsts = False
        End Select
NextReplConsts:
    Next
End Function

Private Function ReplChkConst(Op As String, LinesPtr As Long) As Boolean
    Dim Arg As String, Value As String
    ReplChkConst = True
    If Op = "" Then Exit Function
    Arg = ExtractOp(Op)
    If GetOpType(Arg) = cConst Then
        If Not GetConstVal(Arg, Value) Then
            ErrDesc = "Undefined constant or label name"
            ReplChkConst = False
            Exit Function
        End If
        If Value <> "" Then
            CompressOp Op, Value
            Arg = Value
        End If
    End If
    Select Case GetOpType(Arg)
    Case cConst
        If Insts(LinesPtr) = "DEFS" Or Insts(LinesPtr) = "ORG" Then
            ErrDesc = "Address variable can't be use as operand"
            ReplChkConst = False
            Exit Function
        End If
    Case cNum
        If Insts(LinesPtr) = "DEFS" And Arg = "0" Then
            ErrDesc = "Operand must be a number that is greater than zero"
            ReplChkConst = False
            Exit Function
        End If
        Select Case GetOpEqv(Op, LinesPtr)
        Case "n8", "(n8)"
            If Val(Arg) > 255 Then
                ErrDesc = "Overflow byte type number"
                ReplChkConst = False
            End If
        Case "(IX+d7)", "(IY+d7)"
            If Mid(Op, 4, 1) = "+" And Val(Arg) > 127 Or _
            Mid(Op, 4, 1) = "-" And Val(Arg) > 128 Then
                ErrDesc = "Overflow displacement type number"
                ReplChkConst = False
            End If
            If Mid(Op, 4, 1) = "-" And Arg = "0" Then Op = Left(Op, 3) & "+0)"
        End Select
    End Select
End Function

Private Function ExtractConstName(Op As String) As String
    ExtractConstName = ExtractOp(Op)
    If GetOpType(ExtractConstName) <> cConst Then ExtractConstName = ""
End Function

Private Function CompressOp(Op As String, Arg As String)
    Select Case GetOpId(Op)
    Case cOpIdAbs
        Op = Arg
    Case cOpIdRef
        Op = "(" & Arg & ")"
    Case cOpIdDisRef
        Op = Left(Op, 4) & Arg & ")"
    End Select
End Function

Private Function ExtractOp(Op As String) As String
    Select Case GetOpId(Op)
    Case cOpIdAbs
        ExtractOp = Op
    Case cOpIdRef
        ExtractOp = Mid(Op, 2, Len(Op) - 2)
    Case cOpIdDisRef
        ExtractOp = Mid(Op, 5, Len(Op) - 5)
    End Select
End Function

Private Function GetOpId(Op As String) As Long
    Dim Prv As String
    If Left(Op, 1) = "(" Then
        Prv = Left(Op, 4)
        If Prv = "(IX+" Or Prv = "(IX-" Or Prv = "(IY+" Or Prv = "(IY-" Then
            GetOpId = cOpIdDisRef
        Else
            GetOpId = cOpIdRef
        End If
    Else
        GetOpId = cOpIdAbs
    End If
End Function

Private Function GetCmdLine(LinesPtr As Long) As String
    Dim Tmp As Long
    If Insts(LinesPtr) = cErrFlag Then
        If Len(Op1s(LinesPtr)) < LblLen + 30 Then _
            GetCmdLine = Space(LblLen + 30 - Len(Op1s(LinesPtr)))
        GetCmdLine = Replace(Op1s(LinesPtr), vbTab, " ") & GetCmdLine & " ; " & _
            Op2s(LinesPtr)
    Else
        If Lbls(LinesPtr) = "" Then
            GetCmdLine = ""
        Else
            GetCmdLine = Lbls(LinesPtr) & ":"
        End If
        If Insts(LinesPtr) <> "" Then
            GetCmdLine = GetCmdLine & Space(LblLen + 5 - Len(GetCmdLine)) & Insts(LinesPtr)
            If Op1s(LinesPtr) <> "" Then
                GetCmdLine = GetCmdLine & Space(8 - Len(Insts(LinesPtr))) & Op1s(LinesPtr)
                If Op2s(LinesPtr) <> "" Then GetCmdLine = GetCmdLine & "," & Op2s(LinesPtr)
            End If
        End If
    End If
End Function

Private Function BuildMnemonicsList() As Boolean
    Dim CmdLine As String, LinesPtr As Long, NumArg As Long, I As Long, ArgType As Long
    BuildMnemonicsList = True
    Open SrcFile For Input As #1
    LinesPtr = 1
    ReDim Lbls(LinesPtr), Insts(LinesPtr), Op1s(LinesPtr), Op2s(LinesPtr)
    Do While Not EOF(1)
        Line Input #1, CmdLine
        If Not GetTrimCmdLine(CmdLine) Then
            AddError CmdLine, ErrDesc, LinesPtr
            BuildMnemonicsList = False
            GoTo IncList
        End If
        If UBound(Args) = 0 Then GoTo SkipLine
        If Not GetMnemonic() Then
            AddError CmdLine, ErrDesc, LinesPtr
            BuildMnemonicsList = False
            GoTo IncList
        End If
        If Args(1) <> "" Then
            If Not AddConst(Args(1)) Then
                AddError CmdLine, "Duplicate label name", LinesPtr
                BuildMnemonicsList = False
                GoTo IncList
            End If
            Lbls(LinesPtr) = Args(1)
        End If
        NumArg = UBound(Args)
        If NumArg = 1 Then GoTo IncList
        Insts(LinesPtr) = Args(2)
        Select Case Args(2)
        Case "DEFB", "DEFW"
            If NumArg = 2 Then
                AddError CmdLine, "Expect one or more operands", LinesPtr
                BuildMnemonicsList = False
            Else
                For I = 3 To NumArg
                    Insts(LinesPtr) = Args(2)
                    Op1s(LinesPtr) = Args(I)
                    LinesPtr = LinesPtr + 1
                    ReDim Preserve Lbls(LinesPtr), Insts(LinesPtr), Op1s(LinesPtr), _
                        Op2s(LinesPtr)
                Next
                GoTo SkipLine
            End If
        Case "DEFS"
            If NumArg = 2 Then
                AddError CmdLine, "Expected a number or constant as operand", LinesPtr
                BuildMnemonicsList = False
            Else
                Op1s(LinesPtr) = Args(3)
            End If
        Case "END"
            If Args(1) = "" Then
                GoTo EndBuildMnemonicsList
            Else
                AddError CmdLine, "Unexpect label name", LinesPtr
                BuildMnemonicsList = False
            End If
        Case "EQU"
            If Args(1) = "" Then
                AddError CmdLine, "Expect label name as identifier", LinesPtr
                BuildMnemonicsList = False
            Else
                If NumArg = 2 Then
                    AddError CmdLine, "Expect a number as operand", LinesPtr
                    BuildMnemonicsList = False
                Else
                    SetConstVal Args(1), Args(3)
                    Op1s(LinesPtr) = Args(3)
                End If
            End If
        Case "ORG"
            If Args(1) <> "" Then
                AddError CmdLine, "Unexpect label name", LinesPtr
                BuildMnemonicsList = False
            Else
                If NumArg = 2 Then
                    AddError CmdLine, "Expect a number as operand", LinesPtr
                    BuildMnemonicsList = False
                Else
                    Op1s(LinesPtr) = Args(3)
                End If
            End If
        Case Else
            If NumArg > 2 Then Op1s(LinesPtr) = Args(3)
            If NumArg > 3 Then Op2s(LinesPtr) = Args(4)
        End Select
IncList:
        LinesPtr = LinesPtr + 1
        ReDim Preserve Lbls(LinesPtr), Insts(LinesPtr), Op1s(LinesPtr), Op2s(LinesPtr)
SkipLine:
    Loop
    LinesPtr = LinesPtr - 1
    ReDim Preserve Lbls(LinesPtr), Insts(LinesPtr), Op1s(LinesPtr), Op2s(LinesPtr)
EndBuildMnemonicsList:
    Close #1
    ReDim Args(0)
End Function

Private Function AddError(CmdLine As String, ErrDescription As String, LinesPtr As Long)
    Insts(LinesPtr) = cErrFlag
    Op1s(LinesPtr) = CmdLine
    Op2s(LinesPtr) = ErrDescription
End Function

Private Function GetFileName(FullName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetBaseName(FullName)
    Set fso = Nothing
End Function

Private Function GetFileExt(FullName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileExt = fso.GetExtensionName(FullName)
    Set fso = Nothing
End Function

Private Function AddConst(Name As String, Optional Value As String) As Boolean
    Dim I As Long
    If Len(Name) > LblLen Then LblLen = Len(Name)
    If UBound(ConstNames) = 0 And ConstNames(0) = "" Then
        ConstNames(0) = Name
        If Not IsMissing(Value) Then ConstVals(0) = Value
    Else
        For I = 0 To UBound(ConstNames)
            If Name = ConstNames(I) Then Exit Function
        Next
        ReDim Preserve ConstNames(I), ConstVals(I)
        ConstNames(I) = Name
        If Not IsMissing(Value) Then ConstVals(I) = Value
    End If
    AddConst = True
End Function

Private Function GetConstVal(Name As String, Value As String) As Boolean
    Dim I As Long
    For I = 0 To UBound(ConstNames)
        If Name = ConstNames(I) Then
            Value = ConstVals(I)
            GetConstVal = True
            Exit Function
        End If
    Next
End Function

Private Function SetConstVal(Name As String, Value As String)
    Dim I As Long
    For I = 0 To UBound(ConstNames)
        If Name = ConstNames(I) Then
            ConstVals(I) = Value
            Exit Function
        End If
    Next
End Function

Private Function GetMnemonic() As Boolean
    Dim Tmp() As String, aChr As String, ArgsPtr As Long, TmpPtr As Long, I As Long
    Dim InParenth As Boolean, InQuo As Boolean, ArgsTop As Long, ArgType As Long
    GetMnemonic = True
    ArgsTop = UBound(Args)
    If ArgsTop = 0 Then Exit Function
    ReDim Tmp(0)
    ArgsPtr = 1
    Select Case GetArgType(Args(ArgsPtr))
    Case cConst
        TmpPtr = TmpPtr + 1
        ReDim Preserve Tmp(TmpPtr)
        Tmp(TmpPtr) = Args(ArgsPtr)
        If ArgsPtr = ArgsTop Then GoTo SuccessGetMnemonic
        ArgsPtr = ArgsPtr + 1
        If Args(ArgsPtr) = ":" Then
            If ArgsPtr = ArgsTop Then GoTo SuccessGetMnemonic
            ArgsPtr = ArgsPtr + 1
        End If
        Select Case GetArgType(Args(ArgsPtr))
        Case cDir, cInst
            TmpPtr = TmpPtr + 1
            ReDim Preserve Tmp(TmpPtr)
            Tmp(TmpPtr) = Args(ArgsPtr)
            If ArgsPtr = ArgsTop Then GoTo SuccessGetMnemonic
            ArgsPtr = ArgsPtr + 1
        Case cNum, cUndef
            ErrDesc = "Syntax error"
            GoTo FailGetMnemonic
        Case cConst
            ErrDesc = "Unrecognized instruction"
            GoTo FailGetMnemonic
        Case Else
            ErrDesc = "Illegal instruction"
            GoTo FailGetMnemonic
        End Select
    Case cDir, cInst
        TmpPtr = TmpPtr + 2
        ReDim Preserve Tmp(TmpPtr)
        Tmp(TmpPtr) = Args(ArgsPtr)
        If ArgsPtr = ArgsTop Then GoTo SuccessGetMnemonic
        ArgsPtr = ArgsPtr + 1
        If Args(ArgsPtr) = ":" Then
            ErrDesc = "Illegal use of label name"
            GoTo FailGetMnemonic
        End If
    Case cNum, cUndef
        ErrDesc = "Syntax error"
        GoTo FailGetMnemonic
    Case Else
        ErrDesc = "Illegal label name or directive or instruction"
        GoTo FailGetMnemonic
    End Select
    If Tmp(2) = "END" Then
        ErrDesc = "Unexpected operand"
        GoTo FailGetMnemonic
    End If
    TmpPtr = TmpPtr + 1
    ReDim Preserve Tmp(TmpPtr)
    For I = ArgsPtr To ArgsTop
        Select Case Args(I)
        Case "("
            If FindDir(Tmp(2)) Then
                ErrDesc = "Illegal use of reference value"
                GoTo FailGetMnemonic
            End If
            If Tmp(TmpPtr) <> "" Then
                ErrDesc = "Expected list separator"
                GoTo FailGetMnemonic
            End If
            InParenth = True
            Tmp(TmpPtr) = Args(I)
        Case ")"
            If Tmp(TmpPtr) = "(" Then
                ErrDesc = "Expected reference register or constant or number"
                GoTo FailGetMnemonic
            End If
            If Right(Tmp(TmpPtr), 1) = "+" Or Right(Tmp(TmpPtr), 1) = "-" Then
                ErrDesc = "Expected constant or number"
                GoTo FailGetMnemonic
            End If
            If Tmp(2) <> "JP" And (Tmp(TmpPtr) = "(IX" Or Tmp(TmpPtr) = "(IY") Then _
                Tmp(TmpPtr) = Left(Tmp(TmpPtr), 3) & "+0"
            InParenth = False
            Tmp(TmpPtr) = Tmp(TmpPtr) & Args(I)
        Case "+", "-"
            If Tmp(TmpPtr) <> "(IX" And Tmp(TmpPtr) <> "(IY" Then
                ErrDesc = "Syntax error"
                GoTo FailGetMnemonic
            End If
            Tmp(TmpPtr) = Tmp(TmpPtr) & Args(I)
        Case "["
            If Tmp(TmpPtr) <> "" Then
                If InParenth Then
                    ErrDesc = "Syntax error"
                Else
                    ErrDesc = "Expected list separator"
                End If
                GoTo FailGetMnemonic
            End If
            InQuo = True
        Case "]"
            If Args(I - 1) = "[" Then
                ErrDesc = "Expected one or more characters"
                GoTo FailGetMnemonic
            End If
            InQuo = False
        Case ","
            If Tmp(TmpPtr) = "" Or I = ArgsTop Or Tmp(2) = "EQU" Or Tmp(2) = "ORG" _
            Or Tmp(2) = "DEFS" Or Tmp(2) <> "DEFB" And Tmp(2) <> "DEFW" And TmpPtr > 3 Then
                ErrDesc = "Syntax Error"
                GoTo FailGetMnemonic
            End If
            TmpPtr = TmpPtr + 1
            ReDim Preserve Tmp(TmpPtr)
        Case ":"
            ErrDesc = "Syntax error"
            GoTo FailGetMnemonic
        Case Else
            If InQuo Then
                If Tmp(2) = "DEFB" Then
                    Tmp(TmpPtr) = Tmp(TmpPtr) & Args(I)
                    If Args(I + 1) <> "]" Then
                        TmpPtr = TmpPtr + 1
                        ReDim Preserve Tmp(TmpPtr)
                    End If
                Else
                    If Args(I + 1) = "]" Then
                        Tmp(TmpPtr) = Tmp(TmpPtr) & Args(I)
                    ElseIf Args(I + 2) = "]" Then
                        Tmp(TmpPtr) = Tmp(TmpPtr) & Mid(Str(Val(Args(I)) + _
                            256 * Val(Args(I + 1))), 2)
                        I = I + 1
                    Else
                        ErrDesc = "Overflow number"
                        GoTo FailGetMnemonic
                    End If
                End If
            Else
                ArgType = GetArgType(Args(I))
                If ArgType = cNum Then ArgType = GetNumType(Args(I))
                If ArgType And cBadOvNum Then GoTo FailGetMnemonic
                If ArgType And (cDir Or cInst) Then
                    ErrDesc = "Attempt to use instruction or directive as operand"
                    GoTo FailGetMnemonic
                End If
                If ArgType = cFlagId Then
                    ' ZulNs: If Tmp(2) <> "CALL" And Tmp(2) <> "JP" And Tmp(2) <> "JR" Then
		    If Tmp(2) <> "CALL" And Tmp(2) <> "JP" And Tmp(2) <> "JR" And Tmp(2) <> "RET" Then
                        ErrDesc = "Invalid use of flag identifier"
                        GoTo FailGetMnemonic
                    Else
                        If TmpPtr > 3 Then
                            ErrDesc = "Invalid use of flag identifier"
                            GoTo FailGetMnemonic
                        End If
                    End If
                End If
                If InParenth Then
                    Select Case Tmp(TmpPtr)
                    Case "("
                        If ArgType = cFlagId Then
                            ErrDesc = "Attempt to use flag identifier as reference"
                            GoTo FailGetMnemonic
                        End If
                        If ArgType = cReg Then
                            ErrDesc = "Operand must be a reference register"
                            GoTo FailGetMnemonic
                        End If
                    Case "(IX+", "(IX-", "(IY+", "(IY-"
                        If ArgType And cFlagId Then
                            ErrDesc = "Illegal use of flag identifier"
                            GoTo FailGetMnemonic
                        End If
                        If ArgType And (cReg Or cRefReg) Then
                            ErrDesc = "Illegal instruction"
                            GoTo FailGetMnemonic
                        End If
                    Case Else
                        ErrDesc = "Syntax error"
                        GoTo FailGetMnemonic
                    End Select
                    Tmp(TmpPtr) = Tmp(TmpPtr) & Args(I)
                Else
                    If Tmp(TmpPtr) <> "" Then
                        ErrDesc = "Expected list separator"
                        GoTo FailGetMnemonic
                    End If
                    If FindDir(Tmp(2)) Then
                        Select Case Tmp(2)
                        Case "EQU"
                            If ArgType <> cNum Then
                                ErrDesc = "Operand must be a valid number"
                                GoTo FailGetMnemonic
                            End If
                        Case Else
                            If ArgType <> cNum And ArgType <> cConst Then
                                ErrDesc = "Operand must be a valid constant or number"
                                GoTo FailGetMnemonic
                            End If
                        End Select
                    End If
                    Tmp(TmpPtr) = Args(I)
                End If
            End If
        End Select
    Next
SuccessGetMnemonic:
    Args() = Tmp()
    ErrDesc = ""
    Exit Function
FailGetMnemonic:
    GetMnemonic = False
    ReDim Args(0)
End Function

Private Function GetNumType(Num As String) As Long
    Dim aChr As String, NumHold As Long, NumLen As Long, I As Long
    NumLen = Len(Num)
    ErrDesc = "Invalid number"
    GetNumType = cBadNum
    If NumLen > 1 Then
        Select Case Right(Num, 1)
        Case "0" To "9"
            For I = 1 To NumLen
                aChr = Mid(Num, I, 1)
                If aChr < "0" Or aChr > "9" Then Exit Function
                If NumHold < 65536 Then
                    NumHold = NumHold + Val(aChr) * 10 ^ (NumLen - I)
                End If
            Next
        Case "D"
            For I = 1 To NumLen - 1
                aChr = Mid(Num, I, 1)
                If aChr < "0" Or aChr > "9" Then Exit Function
                If NumHold < 65536 Then
                    NumHold = NumHold + Val(aChr) * 10 ^ (NumLen - I - 1)
                End If
            Next
        Case "H"
            For I = 1 To NumLen - 1
                aChr = Mid(Num, I, 1)
                If aChr < "0" Or aChr > "9" And aChr < "A" Or aChr > "F" Then _
                    Exit Function
                If NumHold < 65536 Then
                    NumHold = NumHold + Val("&H" & aChr) * 16 ^ (NumLen - I - 1)
                End If
            Next
        Case "B"
            For I = 1 To NumLen - 1
                aChr = Mid(Num, I, 1)
                If aChr < "0" Or aChr > "1" Then Exit Function
                If NumHold < 65536 Then
                    NumHold = NumHold + Val(aChr) * 2 ^ (NumLen - I - 1)
                End If
            Next
        Case Else
            Exit Function
        End Select
        If NumHold > 65535 Then
            ErrDesc = "Overflow number"
            GetNumType = cOvNum
            Exit Function
        End If
    Else
        NumHold = Val(Num)
    End If
    Num = Mid(Str(NumHold), 2)
    ErrDesc = ""
    GetNumType = cNum
End Function

Private Function GetArgType(Arg As String) As Long
    GetArgType = GetOpType(Arg)
    If GetArgType = cConst Then
        Select Case Left(Arg, 1)
        Case "A" To "Z"
            If FindDir(Arg) Then
                GetArgType = cDir
                Exit Function
            End If
            If FindInst(Arg) Then
                GetArgType = cInst
                Exit Function
            End If
        Case "_"
        Case ""
            GetArgType = cEmpty
        Case Else
            GetArgType = cUndef
        End Select
    End If
End Function

Private Function GetOpType(Op As String) As Long
    Select Case Left(Op, 1)
    Case "0" To "9"
        GetOpType = cNum
    Case "A" To "Z", "_"
        If FindReg(Op) Then GetOpType = GetOpType + cReg
        If FindRefReg(Op) Then GetOpType = GetOpType + cRefReg
        If FindFlagId(Op) Then GetOpType = GetOpType + cFlagId
        If GetOpType = 0 Then GetOpType = cConst
    End Select
End Function

Private Function FindDir(Arg As String) As Boolean
    FindDir = FindArg(Arg, DirsList)
End Function

Private Function FindInst(Arg As String) As Boolean
    FindInst = FindArg(Arg, InstsList)
End Function

Private Function FindReg(Arg As String) As Boolean
    FindReg = FindArg(Arg, RegsList)
End Function

Private Function FindRefReg(Arg As String) As Boolean
    FindRefReg = FindArg(Arg, RefRegsList)
End Function

Private Function FindFlagId(Arg As String) As Boolean
    FindFlagId = FindArg(Arg, FlagIdsList)
End Function

Private Function FindArg(Arg As String, List As Variant) As Boolean
    Dim I As Long
    For I = 0 To UBound(List)
        If Arg = List(I) Then
            FindArg = True
            Exit Function
        End If
    Next
End Function

Private Function GetTrimCmdLine(CmdLine As String) As Boolean
    Dim NumArgs As Long, CmdLnLen As Long, I As Long, aChr As String
    Dim InArg As Boolean, InQuo As Boolean, InParenth As Boolean
    CmdLnLen = Len(CmdLine)
    ReDim Args(0)
    If CmdLnLen = 0 Then
        GetTrimCmdLine = True
        Exit Function
    End If
    InArg = False
    InQuo = False
    ErrDesc = "Syntax error"
    For I = 1 To CmdLnLen
        aChr = Mid(CmdLine, I, 1)
        If InQuo Then
            NumArgs = NumArgs + 1
            ReDim Preserve Args(NumArgs)
            If aChr = "'" Then
                Args(NumArgs) = "]"
                InQuo = False
            Else
                Args(NumArgs) = Mid(Str(Asc(aChr)), 2)
            End If
        Else
            Select Case aChr
            Case ";"
                Exit For
            Case " ", vbTab
                InArg = False
            Case ":", ",", "(", ")", "+", "-"
                NumArgs = NumArgs + 1
                ReDim Preserve Args(NumArgs)
                Args(NumArgs) = aChr
                InArg = False
                If aChr = "(" Then
                    If InParenth Then Exit Function
                    InParenth = True
                ElseIf aChr = ")" Then
                    If Not InParenth Then Exit Function
                    InParenth = False
                End If
            Case "'"
                If InArg Then
                    If Args(NumArgs) = "AF" Then
                        Args(NumArgs) = Args(NumArgs) & aChr
                    Else
                        Exit Function
                    End If
                Else
                    NumArgs = NumArgs + 1
                    ReDim Preserve Args(NumArgs)
                    Args(NumArgs) = "["
                    InQuo = True
                End If
            Case Else
                If Not InArg Then
                    NumArgs = NumArgs + 1
                    ReDim Preserve Args(NumArgs)
                    InArg = True
                End If
                If Args(NumArgs) = "AF'" Then Exit Function
                If Asc(aChr) >= &H61 And Asc(aChr) <= &H7A Then _
                    aChr = Chr(Asc(aChr) - &H20)
                If aChr >= "0" And aChr <= "9" Or aChr >= "A" And aChr <= "Z" _
                Or aChr = "_" Then
                    Args(NumArgs) = Args(NumArgs) & aChr
                Else
                    Exit Function
                End If
            End Select
        End If
    Next
    If InQuo Or InParenth Then
        ReDim Args(0)
        Exit Function
    End If
    ErrDesc = ""
    GetTrimCmdLine = True
End Function
