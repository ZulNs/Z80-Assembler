Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modMain
	
	'=========================================================================================='
	'                                                                                          '
	'              Z80 Assembler for ED-Laboratory's Microprocessor Trainer MPT-1              '
	'                                                                                          '
	'                Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005                '
	'                                                                                          '
	'=========================================================================================='
	
	
	
	Const cOpCodesTblFileName As String = "OPCODES.TBL"
	Public Const cSrcFileExt As String = "ASM"
	Const cOutFileExt As String = "Z80"
	Const cBinFileExt As String = "BIN"
	Const cErrFileExt As String = "ERR"
	Const cLogFileExt As String = "LOG"
	Const cErrFlag As String = "$Err"
	Const cBinFileStartStr As String = "<Z80_Executable_Codes>"
	Const cBinFileEndStr As String = "<ZulNs#05-11-1970#Viva_New_Technology_Protocol#Gorontalo#Feb-2005>"
	Const cDefStartAddr As Integer = &H1800s
	Const cDir As Integer = 1
	Const cInst As Integer = 2
	Const cReg As Integer = 4
	Const cRefReg As Integer = 8
	Const cFlagId As Integer = &H10s
	Const cNum As Integer = &H20s
	Const cOvNum As Integer = &H40s
	Const cBadNum As Integer = &H80s
	Const cBadOvNum As Integer = cBadNum Or cOvNum
	Const cEmpty As Integer = &H100s
	Const cUndef As Integer = &H200s
	Const cConst As Integer = &H400s
	Const cConstNum As Integer = cConst Or cNum
	Const cOpIdAbs As Integer = 1
	Const cOpIdRef As Integer = 2
	Const cOpIdDisRef As Integer = 4
	
	Dim FlagIdsList, InstsList, DirsList, RegsList, RefRegsList As Object
	Dim ConstNames() As String
	Dim ConstVals() As String
	Dim InstsTbl(695, 3) As String
	Dim Lbls() As String
	Dim Insts() As String
	Dim Op1s() As String
	Dim Op2s() As String
	Dim OpCodes() As String
	Dim Args() As String
	Dim ErrDesc As String
	Dim LblLen As Integer
	Dim PrgPath, SrcFile, OutFile As String
	
	'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
	Public Sub Main()
		Dim blnSuccess As Boolean
		'    If Not IsFileExist(cOpCodesTblFileName) Then
		'        MsgBox "Can't found '" & cOpCodesTblFileName & _
		''            "' file. Assembling process aborted.", vbCritical
		'        Exit Sub
		'    End If
		If Not GetSrcFileName Then Exit Sub
		InitVars()
		blnSuccess = BuildMnemonicsList
		blnSuccess = ReplConsts And blnSuccess
		blnSuccess = SetAddrsVal And blnSuccess
		blnSuccess = ReplConsts And blnSuccess
		If blnSuccess Then WrtOutFile() Else WrtErrFile()
		Shell("Notepad.exe " & OutFile, AppWinStyle.NormalFocus)
	End Sub
	
	Private Function GetSrcFileName() As Boolean
		Dim CmdTail, Path As String
		Dim fso As Object
		Dim UserRespons As MsgBoxResult
		PrgPath = CurDir()
		CmdTail = VB.Command()
		If CmdTail = "" Then
			If Not GetSrcFileNameFromDlg Then Exit Function
		Else
			fso = CreateObject("Scripting.FileSystemObject")
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetExtensionName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not fso.FileExists(CmdTail) Then If fso.GetExtensionName(CmdTail) = "" Then CmdTail = CmdTail & "." & cSrcFileExt
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If fso.FileExists(CmdTail) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetParentFolderName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Path = fso.GetParentFolderName(CmdTail)
				If Path = "" Then
					SrcFile = CmdTail
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetDriveName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ChDrive(fso.GetDriveName(Path))
					ChDir(Path)
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetFileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SrcFile = fso.GetFileName(CmdTail)
				End If
				'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				fso = Nothing
			Else
				'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				fso = Nothing
				UserRespons = MsgBox("Can't found '" & CmdTail & "' file or '" & CmdTail & "' is not a legal file name." & vbCr & "Try to find it or another file by your self?", MsgBoxStyle.Question + MsgBoxStyle.OKCancel)
				If UserRespons = MsgBoxResult.OK Then
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
		DelFile(OutFile & "." & cOutFileExt)
		DelFile(OutFile & "." & cErrFileExt)
		DelFile(OutFile & "." & cLogFileExt)
		DelFile(OutFile & "." & cBinFileExt)
		GetSrcFileName = True
	End Function
	
	Private Function GetSrcFileNameFromDlg() As Boolean
		Dim dlg As New frmDlgFileOpen
		Dim blnExit As Boolean
		dlg.ShowDialog()
		blnExit = dlg.ExitMode
		If blnExit Then SrcFile = dlg.FileName
		dlg.Close()
		'UPGRADE_NOTE: Object dlg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		dlg = Nothing
		If blnExit Then GetSrcFileNameFromDlg = True Else MsgBox("No file selected. Assembling process aborted.", MsgBoxStyle.Information)
	End Function
	
	Private Function IsFileExist(ByRef FileName As String) As Boolean
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IsFileExist = fso.FileExists(FileName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	Private Function DelFile(ByRef FileName As String) As Object
		Dim fso As Object
		If IsFileExist(FileName) Then
			fso = CreateObject("Scripting.FileSystemObject")
			'UPGRADE_WARNING: Couldn't resolve default property of object fso.DeleteFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fso.DeleteFile(FileName, True)
			'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			fso = Nothing
		End If
	End Function
	
	Private Function InitVars() As Object
		Dim I, J As Integer
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object DirsList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DirsList = New Object(){"END", "EQU", "DEFB", "DEFS", "DEFW", "ORG"}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object InstsList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstsList = New Object(){"ADC", "ADD", "AND", "BIT", "CALL", "CCF", "CP", "CPD", "CPDR", "CPI", "CPIR", "CPL", "DAA", "DEC", "DI", "DJNZ", "EI", "EX", "EXX", "HALT", "IM", "IN", "INC", "IND", "INDR", "INI", "INIR", "JP", "JR", "LD", "LDD", "LDDR", "LDI", "LDIR", "NEG", "NOP", "OR", "OTDR", "OTIR", "OUT", "OUTD", "OUTI", "POP", "PUSH", "RES", "RET", "RETI", "RETN", "RL", "RLA", "RLC", "RLCA", "RLD", "RR", "RRA", "RRC", "RRCA", "RRD", "RST", "SBC", "SCF", "SET", "SLA", "SRA", "SRL", "SUB", "XOR"}
		' ZulNs: RegsList = Array("A", "B", "C", "D", "E", "H", "L", "AF", "AF'", "BC", "DE", "HL", _
		'' ZulNs:    "IX", "IY", "SP")
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object RegsList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RegsList = New Object(){"A", "B", "C", "D", "E", "H", "L", "I", "R", "AF", "AF'", "BC", "DE", "HL", "IX", "IY", "SP"}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object RefRegsList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RefRegsList = New Object(){"C", "BC", "DE", "HL", "IX", "IY", "SP"}
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlagIdsList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FlagIdsList = New Object(){"Z", "NZ", "C", "NC", "P", "M", "PO", "PE"}
		DelFile(PrgPath & "\" & cOpCodesTblFileName)
		FileOpen(1, PrgPath & "\" & cOpCodesTblFileName, OpenMode.Output)
		PrintLine(1, "DEFB,n8,,n8")
		PrintLine(1, "DEFW,n16,,nLnH")
		PrintLine(1, "ADC,A,(HL),8E")
		PrintLine(1, "ADC,A,(IX+d7),DD8Ed7")
		PrintLine(1, "ADC,A,(IY+d7),FD8Ed7")
		PrintLine(1, "ADC,A,A,8F")
		PrintLine(1, "ADC,A,B,88")
		PrintLine(1, "ADC,A,C,89")
		PrintLine(1, "ADC,A,D,8A")
		PrintLine(1, "ADC,A,E,8B")
		PrintLine(1, "ADC,A,H,8C")
		PrintLine(1, "ADC,A,L,8D")
		PrintLine(1, "ADC,A,n8,CEn8")
		PrintLine(1, "ADC,HL,BC,ED4A")
		PrintLine(1, "ADC,HL,DE,ED5A")
		PrintLine(1, "ADC,HL,HL,ED6A")
		PrintLine(1, "ADC,HL,SP,ED7A")
		PrintLine(1, "ADD,A,(HL),86")
		PrintLine(1, "ADD,A,(IX+d7),DD86d7")
		PrintLine(1, "ADD,A,(IY+d7),FD86d7")
		PrintLine(1, "ADD,A,A,87")
		PrintLine(1, "ADD,A,B,80")
		PrintLine(1, "ADD,A,C,81")
		PrintLine(1, "ADD,A,D,82")
		PrintLine(1, "ADD,A,E,83")
		PrintLine(1, "ADD,A,H,84")
		PrintLine(1, "ADD,A,L,85")
		PrintLine(1, "ADD,A,n8,C6n8")
		PrintLine(1, "ADD,HL,BC,09")
		PrintLine(1, "ADD,HL,DE,19")
		PrintLine(1, "ADD,HL,HL,29")
		PrintLine(1, "ADD,HL,SP,39")
		PrintLine(1, "ADD,IX,BC,DD09")
		PrintLine(1, "ADD,IX,DE,DD19")
		PrintLine(1, "ADD,IX,IX,DD29")
		PrintLine(1, "ADD,IX,SP,DD39")
		PrintLine(1, "ADD,IY,BC,FD09")
		PrintLine(1, "ADD,IY,DE,FD19")
		PrintLine(1, "ADD,IY,IY,FD29")
		PrintLine(1, "ADD,IY,SP,FD39")
		PrintLine(1, "AND,(HL),,A6")
		PrintLine(1, "AND,(IX+d7),,DDA6d7")
		PrintLine(1, "AND,(IY+d7),,FDA6d7")
		PrintLine(1, "AND,A,,A7")
		PrintLine(1, "AND,B,,A0")
		PrintLine(1, "AND,C,,A1")
		PrintLine(1, "AND,D,,A2")
		PrintLine(1, "AND,E,,A3")
		PrintLine(1, "AND,H,,A4")
		PrintLine(1, "AND,L,,A5")
		PrintLine(1, "AND,n8,,E6n8")
		PrintLine(1, "BIT,0,(HL),CB46")
		PrintLine(1, "BIT,0,(IX+d7),DDCBd746")
		PrintLine(1, "BIT,0,(IY+d7),FDCBd746")
		PrintLine(1, "BIT,0,A,CB47")
		PrintLine(1, "BIT,0,B,CB40")
		PrintLine(1, "BIT,0,C,CB41")
		PrintLine(1, "BIT,0,D,CB42")
		PrintLine(1, "BIT,0,E,CB43")
		PrintLine(1, "BIT,0,H,CB44")
		PrintLine(1, "BIT,0,L,CB45")
		PrintLine(1, "BIT,1,(HL),CB4E")
		PrintLine(1, "BIT,1,(IX+d7),DDCBd74E")
		PrintLine(1, "BIT,1,(IY+d7),FDCBd74E")
		PrintLine(1, "BIT,1,A,CB4F")
		PrintLine(1, "BIT,1,B,CB48")
		PrintLine(1, "BIT,1,C,CB49")
		PrintLine(1, "BIT,1,D,CB4A")
		PrintLine(1, "BIT,1,E,CB4B")
		PrintLine(1, "BIT,1,H,CB4C")
		PrintLine(1, "BIT,1,L,CB4D")
		PrintLine(1, "BIT,2,(HL),CB56")
		PrintLine(1, "BIT,2,(IX+d7),DDCBd756")
		PrintLine(1, "BIT,2,(IY+d7),FDCBd756")
		PrintLine(1, "BIT,2,A,CB57")
		PrintLine(1, "BIT,2,B,CB50")
		PrintLine(1, "BIT,2,C,CB51")
		PrintLine(1, "BIT,2,D,CB52")
		PrintLine(1, "BIT,2,E,CB53")
		PrintLine(1, "BIT,2,H,CB54")
		PrintLine(1, "BIT,2,L,CB55")
		PrintLine(1, "BIT,3,(HL),CB5E")
		PrintLine(1, "BIT,3,(IX+d7),DDCBd75E")
		PrintLine(1, "BIT,3,(IY+d7),FDCBd75E")
		PrintLine(1, "BIT,3,A,CB5F")
		PrintLine(1, "BIT,3,B,CB58")
		PrintLine(1, "BIT,3,C,CB59")
		PrintLine(1, "BIT,3,D,CB5A")
		PrintLine(1, "BIT,3,E,CB5B")
		PrintLine(1, "BIT,3,H,CB5C")
		PrintLine(1, "BIT,3,L,CB5D")
		PrintLine(1, "BIT,4,(HL),CB66")
		PrintLine(1, "BIT,4,(IX+d7),DDCBd766")
		PrintLine(1, "BIT,4,(IY+d7),FDCBd766")
		PrintLine(1, "BIT,4,A,CB67")
		PrintLine(1, "BIT,4,B,CB60")
		PrintLine(1, "BIT,4,C,CB61")
		PrintLine(1, "BIT,4,D,CB62")
		PrintLine(1, "BIT,4,E,CB63")
		PrintLine(1, "BIT,4,H,CB64")
		PrintLine(1, "BIT,4,L,CB65")
		PrintLine(1, "BIT,5,(HL),CB6E")
		PrintLine(1, "BIT,5,(IX+d7),DDCBd76E")
		PrintLine(1, "BIT,5,(IY+d7),FDCBd76E")
		PrintLine(1, "BIT,5,A,CB6F")
		PrintLine(1, "BIT,5,B,CB68")
		PrintLine(1, "BIT,5,C,CB69")
		PrintLine(1, "BIT,5,D,CB6A")
		PrintLine(1, "BIT,5,E,CB6B")
		PrintLine(1, "BIT,5,H,CB6C")
		PrintLine(1, "BIT,5,L,CB6D")
		PrintLine(1, "BIT,6,(HL),CB76")
		PrintLine(1, "BIT,6,(IX+d7),DDCBd776")
		PrintLine(1, "BIT,6,(IY+d7),FDCBd776")
		PrintLine(1, "BIT,6,A,CB77")
		PrintLine(1, "BIT,6,B,CB70")
		PrintLine(1, "BIT,6,C,CB71")
		PrintLine(1, "BIT,6,D,CB72")
		PrintLine(1, "BIT,6,E,CB73")
		PrintLine(1, "BIT,6,H,CB74")
		PrintLine(1, "BIT,6,L,CB75")
		PrintLine(1, "BIT,7,(HL),CB7E")
		PrintLine(1, "BIT,7,(IX+d7),DDCBd77E")
		PrintLine(1, "BIT,7,(IY+d7),FDCBd77E")
		PrintLine(1, "BIT,7,A,CB7F")
		PrintLine(1, "BIT,7,B,CB78")
		PrintLine(1, "BIT,7,C,CB79")
		PrintLine(1, "BIT,7,D,CB7A")
		PrintLine(1, "BIT,7,E,CB7B")
		PrintLine(1, "BIT,7,H,CB7C")
		PrintLine(1, "BIT,7,L,CB7D")
		PrintLine(1, "CALL,C,n16,DCnLnH")
		PrintLine(1, "CALL,M,n16,FCnLnH")
		PrintLine(1, "CALL,n16,,CDnLnH")
		PrintLine(1, "CALL,NC,n16,D4nLnH")
		PrintLine(1, "CALL,NZ,n16,C4nLnH")
		PrintLine(1, "CALL,P,n16,F4nLnH")
		PrintLine(1, "CALL,PE,n16,ECnLnH")
		PrintLine(1, "CALL,PO,n16,E4nLnH")
		PrintLine(1, "CALL,Z,n16,CCnLnH")
		PrintLine(1, "CCF,,,3F")
		PrintLine(1, "CP,(HL),,BE")
		PrintLine(1, "CP,(IX+d7),,DDBEd7")
		PrintLine(1, "CP,(IY+d7),,FDBEd7")
		PrintLine(1, "CP,A,,BF")
		PrintLine(1, "CP,B,,B8")
		PrintLine(1, "CP,C,,B9")
		PrintLine(1, "CP,D,,BA")
		PrintLine(1, "CP,E,,BB")
		PrintLine(1, "CP,H,,BC")
		PrintLine(1, "CP,L,,BD")
		PrintLine(1, "CP,n8,,FEn8")
		PrintLine(1, "CPD,,,EDA9")
		PrintLine(1, "CPDR,,,EDB9")
		PrintLine(1, "CPI,,,EDA1")
		PrintLine(1, "CPIR,,,EDB1")
		PrintLine(1, "CPL,,,2F")
		PrintLine(1, "DAA,,,27")
		PrintLine(1, "DEC,(HL),,35")
		PrintLine(1, "DEC,(IX+d7),,DD35d7")
		PrintLine(1, "DEC,(IY+d7),,FD35d7")
		PrintLine(1, "DEC,A,,3D")
		PrintLine(1, "DEC,B,,05")
		PrintLine(1, "DEC,BC,,0B")
		PrintLine(1, "DEC,C,,0D")
		PrintLine(1, "DEC,D,,15")
		PrintLine(1, "DEC,DE,,1B")
		PrintLine(1, "DEC,E,,1D")
		PrintLine(1, "DEC,H,,25")
		PrintLine(1, "DEC,HL,,2B")
		PrintLine(1, "DEC,IX,,DD2B")
		PrintLine(1, "DEC,IY,,FD2B")
		PrintLine(1, "DEC,L,,2D")
		PrintLine(1, "DEC,SP,,3B")
		PrintLine(1, "DI,,,F3")
		PrintLine(1, "DJNZ,d7,,10d7")
		PrintLine(1, "EI,,,FB")
		PrintLine(1, "EX,(SP),HL,E3")
		PrintLine(1, "EX,(SP),IX,DDE3")
		PrintLine(1, "EX,(SP),IY,FDE3")
		PrintLine(1, "EX,AF,AF',08")
		PrintLine(1, "EX,DE,HL,EB")
		PrintLine(1, "EXX,,,D9")
		PrintLine(1, "HALT,,,76")
		PrintLine(1, "IM,0,,ED46")
		PrintLine(1, "IM,1,,ED56")
		PrintLine(1, "IM,2,,ED5E")
		PrintLine(1, "IN,A,(C),ED78")
		PrintLine(1, "IN,A,(n8),DBn8")
		PrintLine(1, "IN,B,(C),ED40")
		PrintLine(1, "IN,C,(C),ED48")
		PrintLine(1, "IN,D,(C),ED50")
		PrintLine(1, "IN,E,(C),ED58")
		PrintLine(1, "IN,H,(C),ED60")
		PrintLine(1, "IN,L,(C),ED68")
		PrintLine(1, "INC,(HL),,34")
		PrintLine(1, "INC,(IX+d7),,DD34d7")
		PrintLine(1, "INC,(IY+d7),,FD34d7")
		PrintLine(1, "INC,A,,3C")
		PrintLine(1, "INC,B,,04")
		PrintLine(1, "INC,BC,,03")
		PrintLine(1, "INC,C,,0C")
		PrintLine(1, "INC,D,,14")
		PrintLine(1, "INC,DE,,13")
		PrintLine(1, "INC,E,,1C")
		PrintLine(1, "INC,H,,24")
		PrintLine(1, "INC,HL,,23")
		PrintLine(1, "INC,IX,,DD23")
		PrintLine(1, "INC,IY,,FD23")
		PrintLine(1, "INC,L,,2C")
		PrintLine(1, "INC,SP,,33")
		PrintLine(1, "IND,,,EDAA")
		PrintLine(1, "INDR,,,EDBA")
		PrintLine(1, "INI,,,EDA2")
		PrintLine(1, "INIR,,,EDB2")
		PrintLine(1, "JP,(HL),,E9")
		PrintLine(1, "JP,(IX),,DDE9")
		PrintLine(1, "JP,(IY),,FDE9")
		PrintLine(1, "JP,C,n16,DAnLnH")
		PrintLine(1, "JP,M,n16,FAnLnH")
		PrintLine(1, "JP,n16,,C3nLnH")
		PrintLine(1, "JP,NC,n16,D2nLnH")
		PrintLine(1, "JP,NZ,n16,C2nLnH")
		PrintLine(1, "JP,P,n16,F2nLnH")
		PrintLine(1, "JP,PE,n16,EAnLnH")
		PrintLine(1, "JP,PO,n16,E2nLnH")
		PrintLine(1, "JP,Z,n16,CAnLnH")
		PrintLine(1, "JR,C,d7,38d7")
		PrintLine(1, "JR,d7,,18d7")
		PrintLine(1, "JR,NC,d7,30d7")
		PrintLine(1, "JR,NZ,d7,20d7")
		PrintLine(1, "JR,Z,d7,28d7")
		PrintLine(1, "LD,(BC),A,02")
		PrintLine(1, "LD,(DE),A,12")
		PrintLine(1, "LD,(HL),A,77")
		PrintLine(1, "LD,(HL),B,70")
		PrintLine(1, "LD,(HL),C,71")
		PrintLine(1, "LD,(HL),D,72")
		PrintLine(1, "LD,(HL),E,73")
		PrintLine(1, "LD,(HL),H,74")
		PrintLine(1, "LD,(HL),L,75")
		PrintLine(1, "LD,(HL),n8,36n8")
		PrintLine(1, "LD,(IX+d7),A,DD77d7")
		PrintLine(1, "LD,(IX+d7),B,DD70d7")
		PrintLine(1, "LD,(IX+d7),C,DD71d7")
		PrintLine(1, "LD,(IX+d7),D,DD72d7")
		PrintLine(1, "LD,(IX+d7),E,DD73d7")
		PrintLine(1, "LD,(IX+d7),H,DD74d7")
		PrintLine(1, "LD,(IX+d7),L,DD75d7")
		PrintLine(1, "LD,(IX+d7),n8,DD36d7n8")
		PrintLine(1, "LD,(IY+d7),A,FD77d7")
		PrintLine(1, "LD,(IY+d7),B,FD70d7")
		PrintLine(1, "LD,(IY+d7),C,FD71d7")
		PrintLine(1, "LD,(IY+d7),D,FD72d7")
		PrintLine(1, "LD,(IY+d7),E,FD73d7")
		PrintLine(1, "LD,(IY+d7),H,FD74d7")
		PrintLine(1, "LD,(IY+d7),L,FD75d7")
		PrintLine(1, "LD,(IY+d7),n8,FD36d7n8")
		PrintLine(1, "LD,(n16),A,32nLnH")
		PrintLine(1, "LD,(n16),BC,ED43nLnH")
		PrintLine(1, "LD,(n16),DE,ED53nLnH")
		PrintLine(1, "LD,(n16),HL,22nLnH")
		PrintLine(1, "LD,(n16),IX,DD22nLnH")
		PrintLine(1, "LD,(n16),IY,FD22nLnH")
		PrintLine(1, "LD,(n16),SP,ED73nLnH")
		PrintLine(1, "LD,A,(BC),0A")
		PrintLine(1, "LD,A,(DE),1A")
		PrintLine(1, "LD,A,(HL),7E")
		PrintLine(1, "LD,A,(IX+d7),DD7Ed7")
		PrintLine(1, "LD,A,(IY+d7),FD7Ed7")
		PrintLine(1, "LD,A,(n16),3AnLnH")
		PrintLine(1, "LD,A,A,7F")
		PrintLine(1, "LD,A,B,78")
		PrintLine(1, "LD,A,C,79")
		PrintLine(1, "LD,A,D,7A")
		PrintLine(1, "LD,A,E,7B")
		PrintLine(1, "LD,A,H,7C")
		PrintLine(1, "LD,A,I,ED57")
		PrintLine(1, "LD,A,L,7D")
		PrintLine(1, "LD,A,R,ED5F")
		PrintLine(1, "LD,A,n8,3En8")
		PrintLine(1, "LD,B,(HL),46")
		PrintLine(1, "LD,B,(IX+d7),DD46d7")
		PrintLine(1, "LD,B,(IY+d7),FD46d7")
		PrintLine(1, "LD,B,A,47")
		PrintLine(1, "LD,B,B,40")
		PrintLine(1, "LD,B,C,41")
		PrintLine(1, "LD,B,D,42")
		PrintLine(1, "LD,B,E,43")
		PrintLine(1, "LD,B,H,44")
		PrintLine(1, "LD,B,L,45")
		PrintLine(1, "LD,B,n8,06n8")
		PrintLine(1, "LD,BC,(n16),ED4BnLnH")
		PrintLine(1, "LD,BC,n16,01nLnH")
		PrintLine(1, "LD,C,(HL),4E")
		PrintLine(1, "LD,C,(IX+d7),DD4Ed7")
		PrintLine(1, "LD,C,(IY+d7),FD4Ed7")
		PrintLine(1, "LD,C,A,4F")
		PrintLine(1, "LD,C,B,48")
		PrintLine(1, "LD,C,C,49")
		PrintLine(1, "LD,C,D,4A")
		PrintLine(1, "LD,C,E,4B")
		PrintLine(1, "LD,C,H,4C")
		PrintLine(1, "LD,C,L,4D")
		PrintLine(1, "LD,C,n8,0En8")
		PrintLine(1, "LD,D,(HL),56")
		PrintLine(1, "LD,D,(IX+d7),DD56d7")
		PrintLine(1, "LD,D,(IY+d7),FD56d7")
		PrintLine(1, "LD,D,A,57")
		PrintLine(1, "LD,D,B,50")
		PrintLine(1, "LD,D,C,51")
		PrintLine(1, "LD,D,D,52")
		PrintLine(1, "LD,D,E,53")
		PrintLine(1, "LD,D,H,54")
		PrintLine(1, "LD,D,L,55")
		PrintLine(1, "LD,D,n8,16n8")
		PrintLine(1, "LD,DE,(n16),ED5BnLnH")
		PrintLine(1, "LD,DE,n16,11nLnH")
		PrintLine(1, "LD,E,(HL),5E")
		PrintLine(1, "LD,E,(IX+d7),DD5Ed7")
		PrintLine(1, "LD,E,(IY+d7),FD5Ed7")
		PrintLine(1, "LD,E,A,5F")
		PrintLine(1, "LD,E,B,58")
		PrintLine(1, "LD,E,C,59")
		PrintLine(1, "LD,E,D,5A")
		PrintLine(1, "LD,E,E,5B")
		PrintLine(1, "LD,E,H,5C")
		PrintLine(1, "LD,E,L,5D")
		PrintLine(1, "LD,E,n8,1En8")
		PrintLine(1, "LD,H,(HL),66")
		PrintLine(1, "LD,H,(IX+d7),DD66d7")
		PrintLine(1, "LD,H,(IY+d7),FD66d7")
		PrintLine(1, "LD,H,A,67")
		PrintLine(1, "LD,H,B,60")
		PrintLine(1, "LD,H,C,61")
		PrintLine(1, "LD,H,D,62")
		PrintLine(1, "LD,H,E,63")
		PrintLine(1, "LD,H,H,64")
		PrintLine(1, "LD,H,L,65")
		PrintLine(1, "LD,H,n8,26n8")
		PrintLine(1, "LD,HL,(n16),2AnLnH")
		PrintLine(1, "LD,HL,n16,21nLnH")
		PrintLine(1, "LD,I,A,ED47")
		PrintLine(1, "LD,IX,(n16),DD2AnLnH")
		PrintLine(1, "LD,IX,n16,DD21nLnH")
		PrintLine(1, "LD,IY,(n16),FD2AnLnH")
		PrintLine(1, "LD,IY,n16,FD21nLnH")
		PrintLine(1, "LD,L,(HL),6E")
		PrintLine(1, "LD,L,(IX+d7),DD6Ed7")
		PrintLine(1, "LD,L,(IY+d7),FD6Ed7")
		PrintLine(1, "LD,L,A,6F")
		PrintLine(1, "LD,L,B,68")
		PrintLine(1, "LD,L,C,69")
		PrintLine(1, "LD,L,D,6A")
		PrintLine(1, "LD,L,E,6B")
		PrintLine(1, "LD,L,H,6C")
		PrintLine(1, "LD,L,L,6D")
		PrintLine(1, "LD,L,n8,2En8")
		PrintLine(1, "LD,R,A,ED4F")
		PrintLine(1, "LD,SP,(n16),ED7BnLnH")
		PrintLine(1, "LD,SP,HL,F9")
		PrintLine(1, "LD,SP,IX,DDF9")
		PrintLine(1, "LD,SP,IY,FDF9")
		PrintLine(1, "LD,SP,n16,31nLnH")
		PrintLine(1, "LDD,,,EDA8")
		PrintLine(1, "LDDR,,,EDB8")
		PrintLine(1, "LDI,,,EDA0")
		PrintLine(1, "LDIR,,,EDB0")
		PrintLine(1, "NEG,,,ED44")
		PrintLine(1, "NOP,,,00")
		PrintLine(1, "OR,(HL),,B6")
		PrintLine(1, "OR,(IX+d7),,DDB6d7")
		PrintLine(1, "OR,(IY+d7),,FDB6d7")
		PrintLine(1, "OR,A,,B7")
		PrintLine(1, "OR,B,,B0")
		PrintLine(1, "OR,C,,B1")
		PrintLine(1, "OR,D,,B2")
		PrintLine(1, "OR,E,,B3")
		PrintLine(1, "OR,H,,B4")
		PrintLine(1, "OR,L,,B5")
		PrintLine(1, "OR,n8,,F6n8")
		PrintLine(1, "OTDR,,,EDBB")
		PrintLine(1, "OTIR,,,EDB3")
		PrintLine(1, "OUT,(C),A,ED79")
		PrintLine(1, "OUT,(C),B,ED41")
		PrintLine(1, "OUT,(C),C,ED49")
		PrintLine(1, "OUT,(C),D,ED51")
		PrintLine(1, "OUT,(C),E,ED59")
		PrintLine(1, "OUT,(C),H,ED61")
		PrintLine(1, "OUT,(C),L,ED69")
		PrintLine(1, "OUT,(n8),A,D3n8")
		PrintLine(1, "OUTD,,,EDAB")
		PrintLine(1, "OUTI,,,EDA3")
		PrintLine(1, "POP,AF,,F1")
		PrintLine(1, "POP,BC,,C1")
		PrintLine(1, "POP,DE,,D1")
		PrintLine(1, "POP,HL,,E1")
		PrintLine(1, "POP,IX,,DDE1")
		PrintLine(1, "POP,IY,,FDE1")
		PrintLine(1, "PUSH,AF,,F5")
		PrintLine(1, "PUSH,BC,,C5")
		PrintLine(1, "PUSH,DE,,D5")
		PrintLine(1, "PUSH,HL,,E5")
		PrintLine(1, "PUSH,IX,,DDE5")
		PrintLine(1, "PUSH,IY,,FDE5")
		PrintLine(1, "RES,0,(HL),CB86")
		PrintLine(1, "RES,0,(IX+d7),DDCBd786")
		PrintLine(1, "RES,0,(IY+d7),FDCBd786")
		PrintLine(1, "RES,0,A,CB87")
		PrintLine(1, "RES,0,B,CB80")
		PrintLine(1, "RES,0,C,CB81")
		PrintLine(1, "RES,0,D,CB82")
		PrintLine(1, "RES,0,E,CB83")
		PrintLine(1, "RES,0,H,CB84")
		PrintLine(1, "RES,0,L,CB85")
		PrintLine(1, "RES,1,(HL),CB8E")
		PrintLine(1, "RES,1,(IX+d7),DDCBd78E")
		PrintLine(1, "RES,1,(IY+d7),FDCBd78E")
		PrintLine(1, "RES,1,A,CB8F")
		PrintLine(1, "RES,1,B,CB88")
		PrintLine(1, "RES,1,C,CB89")
		PrintLine(1, "RES,1,D,CB8A")
		PrintLine(1, "RES,1,E,CB8B")
		PrintLine(1, "RES,1,H,CB8C")
		PrintLine(1, "RES,1,L,CB8D")
		PrintLine(1, "RES,2,(HL),CB96")
		PrintLine(1, "RES,2,(IX+d7),DDCBd796")
		PrintLine(1, "RES,2,(IY+d7),FDCBd796")
		PrintLine(1, "RES,2,A,CB97")
		PrintLine(1, "RES,2,B,CB90")
		PrintLine(1, "RES,2,C,CB91")
		PrintLine(1, "RES,2,D,CB92")
		PrintLine(1, "RES,2,E,CB93")
		PrintLine(1, "RES,2,H,CB94")
		PrintLine(1, "RES,2,L,CB95")
		PrintLine(1, "RES,3,(HL),CB9E")
		PrintLine(1, "RES,3,(IX+d7),DDCBd79E")
		PrintLine(1, "RES,3,(IY+d7),FDCBd79E")
		PrintLine(1, "RES,3,A,CB9F")
		PrintLine(1, "RES,3,B,CB98")
		PrintLine(1, "RES,3,C,CB99")
		PrintLine(1, "RES,3,D,CB9A")
		PrintLine(1, "RES,3,E,CB9B")
		PrintLine(1, "RES,3,H,CB9C")
		PrintLine(1, "RES,3,L,CB9D")
		PrintLine(1, "RES,4,(HL),CBA6")
		PrintLine(1, "RES,4,(IX+d7),DDCBd7A6")
		PrintLine(1, "RES,4,(IY+d7),FDCBd7A6")
		PrintLine(1, "RES,4,A,CBA7")
		PrintLine(1, "RES,4,B,CBA0")
		PrintLine(1, "RES,4,C,CBA1")
		PrintLine(1, "RES,4,D,CBA2")
		PrintLine(1, "RES,4,E,CBA3")
		PrintLine(1, "RES,4,H,CBA4")
		PrintLine(1, "RES,4,L,CBA5")
		PrintLine(1, "RES,5,(HL),CBAE")
		PrintLine(1, "RES,5,(IX+d7),DDCBd7AE")
		PrintLine(1, "RES,5,(IY+d7),FDCBd7AE")
		PrintLine(1, "RES,5,A,CBAF")
		PrintLine(1, "RES,5,B,CBA8")
		PrintLine(1, "RES,5,C,CBA9")
		PrintLine(1, "RES,5,D,CBAA")
		PrintLine(1, "RES,5,E,CBAB")
		PrintLine(1, "RES,5,H,CBAC")
		PrintLine(1, "RES,5,L,CBAD")
		PrintLine(1, "RES,6,(HL),CBB6")
		PrintLine(1, "RES,6,(IX+d7),DDCBd7B6")
		PrintLine(1, "RES,6,(IY+d7),FDCBd7B6")
		PrintLine(1, "RES,6,A,CBB7")
		PrintLine(1, "RES,6,B,CBB0")
		PrintLine(1, "RES,6,C,CBB1")
		PrintLine(1, "RES,6,D,CBB2")
		PrintLine(1, "RES,6,E,CBB3")
		PrintLine(1, "RES,6,H,CBB4")
		PrintLine(1, "RES,6,L,CBB5")
		PrintLine(1, "RES,7,(HL),CBBE")
		PrintLine(1, "RES,7,(IX+d7),DDCBd7BE")
		PrintLine(1, "RES,7,(IY+d7),FDCBd7BE")
		PrintLine(1, "RES,7,A,CBBF")
		PrintLine(1, "RES,7,B,CBB8")
		PrintLine(1, "RES,7,C,CBB9")
		PrintLine(1, "RES,7,D,CBBA")
		PrintLine(1, "RES,7,E,CBBB")
		PrintLine(1, "RES,7,H,CBBC")
		PrintLine(1, "RES,7,L,CBBD")
		PrintLine(1, "RET,C,,D8")
		PrintLine(1, "RET,M,,F8")
		PrintLine(1, "RET,NC,,D0")
		PrintLine(1, "RET,NZ,,C0")
		PrintLine(1, "RET,P,,F0")
		PrintLine(1, "RET,PE,,E8")
		PrintLine(1, "RET,PO,,E0")
		PrintLine(1, "RET,Z,,C8")
		PrintLine(1, "RET,,,C9")
		PrintLine(1, "RETI,,,ED4D")
		PrintLine(1, "RETN,,,ED45")
		PrintLine(1, "RL,(HL),,CB16")
		PrintLine(1, "RL,(IX+d7),,DDCBd716")
		PrintLine(1, "RL,(IY+d7),,FDCBd716")
		PrintLine(1, "RL,A,,CB17")
		PrintLine(1, "RL,B,,CB10")
		PrintLine(1, "RL,C,,CB11")
		PrintLine(1, "RL,D,,CB12")
		PrintLine(1, "RL,E,,CB13")
		PrintLine(1, "RL,H,,CB14")
		PrintLine(1, "RL,L,,CB15")
		PrintLine(1, "RLA,,,17")
		PrintLine(1, "RLC,(HL),,CB06")
		PrintLine(1, "RLC,(IX+d7),,DDCBd706")
		PrintLine(1, "RLC,(IY+d7),,FDCBd706")
		PrintLine(1, "RLC,A,,CB07")
		PrintLine(1, "RLC,B,,CB00")
		PrintLine(1, "RLC,C,,CB01")
		PrintLine(1, "RLC,D,,CB02")
		PrintLine(1, "RLC,E,,CB03")
		PrintLine(1, "RLC,H,,CB04")
		PrintLine(1, "RLC,L,,CB05")
		PrintLine(1, "RLCA,,,07")
		PrintLine(1, "RLD,,,ED6F")
		PrintLine(1, "RR,(HL),,CB1E")
		PrintLine(1, "RR,(IX+d7),,DDCBd71E")
		PrintLine(1, "RR,(IY+d7),,FDCBd71E")
		PrintLine(1, "RR,A,,CB1F")
		PrintLine(1, "RR,B,,CB18")
		PrintLine(1, "RR,C,,CB19")
		PrintLine(1, "RR,D,,CB1A")
		PrintLine(1, "RR,E,,CB1B")
		PrintLine(1, "RR,H,,CB1C")
		PrintLine(1, "RR,L,,CB1D")
		PrintLine(1, "RRA,,,1F")
		PrintLine(1, "RRC,(HL),,CB0E")
		PrintLine(1, "RRC,(IX+d7),,DDCBd70E")
		PrintLine(1, "RRC,(IY+d7),,FDCBd70E")
		PrintLine(1, "RRC,A,,CB0F")
		PrintLine(1, "RRC,B,,CB08")
		PrintLine(1, "RRC,C,,CB09")
		PrintLine(1, "RRC,D,,CB0A")
		PrintLine(1, "RRC,E,,CB0B")
		PrintLine(1, "RRC,H,,CB0C")
		PrintLine(1, "RRC,L,,CB0D")
		PrintLine(1, "RRCA,,,0F")
		PrintLine(1, "RRD,,,ED67")
		PrintLine(1, "RST,0,,C7")
		PrintLine(1, "RST,8,,CF")
		PrintLine(1, "RST,16,,D7")
		PrintLine(1, "RST,24,,DF")
		PrintLine(1, "RST,32,,E7")
		PrintLine(1, "RST,40,,EF")
		PrintLine(1, "RST,48,,F7")
		PrintLine(1, "RST,56,,FF")
		PrintLine(1, "SBC,A,(HL),9E")
		PrintLine(1, "SBC,A,(IX+d7),DD9Ed7")
		PrintLine(1, "SBC,A,(IY+d7),FD9Ed7")
		PrintLine(1, "SBC,A,A,9F")
		PrintLine(1, "SBC,A,B,98")
		PrintLine(1, "SBC,A,C,99")
		PrintLine(1, "SBC,A,D,9A")
		PrintLine(1, "SBC,A,E,9B")
		PrintLine(1, "SBC,A,H,9C")
		PrintLine(1, "SBC,A,L,9D")
		PrintLine(1, "SBC,A,n8,DEn8")
		PrintLine(1, "SBC,HL,BC,ED42")
		PrintLine(1, "SBC,HL,DE,ED52")
		PrintLine(1, "SBC,HL,HL,ED62")
		PrintLine(1, "SBC,HL,SP,ED72")
		PrintLine(1, "SCF,,,37")
		PrintLine(1, "SET,0,(HL),CBC6")
		PrintLine(1, "SET,0,(IX+d7),DDCBd7C6")
		PrintLine(1, "SET,0,(IY+d7),FDCBd7C6")
		PrintLine(1, "SET,0,A,CBC7")
		PrintLine(1, "SET,0,B,CBC0")
		PrintLine(1, "SET,0,C,CBC1")
		PrintLine(1, "SET,0,D,CBC2")
		PrintLine(1, "SET,0,E,CBC3")
		PrintLine(1, "SET,0,H,CBC4")
		PrintLine(1, "SET,0,L,CBC5")
		PrintLine(1, "SET,1,(HL),CBCE")
		PrintLine(1, "SET,1,(IX+d7),DDCBd7CE")
		PrintLine(1, "SET,1,(IY+d7),FDCBd7CE")
		PrintLine(1, "SET,1,A,CBCF")
		PrintLine(1, "SET,1,B,CBC8")
		PrintLine(1, "SET,1,C,CBC9")
		PrintLine(1, "SET,1,D,CBCA")
		PrintLine(1, "SET,1,E,CBCB")
		PrintLine(1, "SET,1,H,CBCC")
		PrintLine(1, "SET,1,L,CBCD")
		PrintLine(1, "SET,2,(HL),CBD6")
		PrintLine(1, "SET,2,(IX+d7),DDCBd7D6")
		PrintLine(1, "SET,2,(IY+d7),FDCBd7D6")
		PrintLine(1, "SET,2,A,CBD7")
		PrintLine(1, "SET,2,B,CBD0")
		PrintLine(1, "SET,2,C,CBD1")
		PrintLine(1, "SET,2,D,CBD2")
		PrintLine(1, "SET,2,E,CBD3")
		PrintLine(1, "SET,2,H,CBD4")
		PrintLine(1, "SET,2,L,CBD5")
		PrintLine(1, "SET,3,(HL),CBDE")
		PrintLine(1, "SET,3,(IX+d7),DDCBd7DE")
		PrintLine(1, "SET,3,(IY+d7),FDCBd7DE")
		PrintLine(1, "SET,3,A,CBDF")
		PrintLine(1, "SET,3,B,CBD8")
		PrintLine(1, "SET,3,C,CBD9")
		PrintLine(1, "SET,3,D,CBDA")
		PrintLine(1, "SET,3,E,CBDB")
		PrintLine(1, "SET,3,H,CBDC")
		PrintLine(1, "SET,3,L,CBDD")
		PrintLine(1, "SET,4,(HL),CBE6")
		PrintLine(1, "SET,4,(IX+d7),DDCBd7E6")
		PrintLine(1, "SET,4,(IY+d7),FDCBd7E6")
		PrintLine(1, "SET,4,A,CBE7")
		PrintLine(1, "SET,4,B,CBE0")
		PrintLine(1, "SET,4,C,CBE1")
		PrintLine(1, "SET,4,D,CBE2")
		PrintLine(1, "SET,4,E,CBE3")
		PrintLine(1, "SET,4,H,CBE4")
		PrintLine(1, "SET,4,L,CBE5")
		PrintLine(1, "SET,5,(HL),CBEE")
		PrintLine(1, "SET,5,(IX+d7),DDCBd7EE")
		PrintLine(1, "SET,5,(IY+d7),FDCBd7EE")
		PrintLine(1, "SET,5,A,CBEF")
		PrintLine(1, "SET,5,B,CBE8")
		PrintLine(1, "SET,5,C,CBE9")
		PrintLine(1, "SET,5,D,CBEA")
		PrintLine(1, "SET,5,E,CBEB")
		PrintLine(1, "SET,5,H,CBEC")
		PrintLine(1, "SET,5,L,CBED")
		PrintLine(1, "SET,6,(HL),CBF6")
		PrintLine(1, "SET,6,(IX+d7),DDCBd7F6")
		PrintLine(1, "SET,6,(IY+d7),FDCBd7F6")
		PrintLine(1, "SET,6,A,CBF7")
		PrintLine(1, "SET,6,B,CBF0")
		PrintLine(1, "SET,6,C,CBF1")
		PrintLine(1, "SET,6,D,CBF2")
		PrintLine(1, "SET,6,E,CBF3")
		PrintLine(1, "SET,6,H,CBF4")
		PrintLine(1, "SET,6,L,CBF5")
		PrintLine(1, "SET,7,(HL),CBFE")
		PrintLine(1, "SET,7,(IX+d7),DDCBd7FE")
		PrintLine(1, "SET,7,(IY+d7),FDCBd7FE")
		PrintLine(1, "SET,7,A,CBFF")
		PrintLine(1, "SET,7,B,CBF8")
		PrintLine(1, "SET,7,C,CBF9")
		PrintLine(1, "SET,7,D,CBFA")
		PrintLine(1, "SET,7,E,CBFB")
		PrintLine(1, "SET,7,H,CBFC")
		PrintLine(1, "SET,7,L,CBFD")
		PrintLine(1, "SLA,(HL),,CB26")
		PrintLine(1, "SLA,(IX+d7),,DDCBd726")
		PrintLine(1, "SLA,(IY+d7),,FDCBd726")
		PrintLine(1, "SLA,A,,CB27")
		PrintLine(1, "SLA,B,,CB20")
		PrintLine(1, "SLA,C,,CB21")
		PrintLine(1, "SLA,D,,CB22")
		PrintLine(1, "SLA,E,,CB23")
		PrintLine(1, "SLA,H,,CB24")
		PrintLine(1, "SLA,L,,CB25")
		PrintLine(1, "SRA,(HL),,CB2E")
		PrintLine(1, "SRA,(IX+d7),,DDCBd72E")
		PrintLine(1, "SRA,(IY+d7),,FDCBd72E")
		PrintLine(1, "SRA,A,,CB2F")
		PrintLine(1, "SRA,B,,CB28")
		PrintLine(1, "SRA,C,,CB29")
		PrintLine(1, "SRA,D,,CB2A")
		PrintLine(1, "SRA,E,,CB2B")
		PrintLine(1, "SRA,H,,CB2C")
		PrintLine(1, "SRA,L,,CB2D")
		PrintLine(1, "SRL,(HL),,CB3E")
		PrintLine(1, "SRL,(IX+d7),,DDCBd73E")
		PrintLine(1, "SRL,(IY+d7),,FDCBd73E")
		PrintLine(1, "SRL,A,,CB3F")
		PrintLine(1, "SRL,B,,CB38")
		PrintLine(1, "SRL,C,,CB39")
		PrintLine(1, "SRL,D,,CB3A")
		PrintLine(1, "SRL,E,,CB3B")
		PrintLine(1, "SRL,H,,CB3C")
		PrintLine(1, "SRL,L,,CB3D")
		PrintLine(1, "SUB,(HL),,96")
		PrintLine(1, "SUB,(IX+d7),,DD96d7")
		PrintLine(1, "SUB,(IY+d7),,FD96d7")
		PrintLine(1, "SUB,A,,97")
		PrintLine(1, "SUB,B,,90")
		PrintLine(1, "SUB,C,,91")
		PrintLine(1, "SUB,D,,92")
		PrintLine(1, "SUB,E,,93")
		PrintLine(1, "SUB,H,,94")
		PrintLine(1, "SUB,L,,95")
		PrintLine(1, "SUB,n8,,D6n8")
		PrintLine(1, "XOR,(HL),,AE")
		PrintLine(1, "XOR,(IX+d7),,DDAEd7")
		PrintLine(1, "XOR,(IY+d7),,FDAEd7")
		PrintLine(1, "XOR,A,,AF")
		PrintLine(1, "XOR,B,,A8")
		PrintLine(1, "XOR,C,,A9")
		PrintLine(1, "XOR,D,,AA")
		PrintLine(1, "XOR,E,,AB")
		PrintLine(1, "XOR,H,,AC")
		PrintLine(1, "XOR,L,,AD")
		PrintLine(1, "XOR,n8,,EEn8")
		FileClose(1)
		FileOpen(1, PrgPath & "\" & cOpCodesTblFileName, OpenMode.Input)
		For I = 0 To UBound(InstsTbl)
			For J = 0 To UBound(InstsTbl, 2)
				Input(1, InstsTbl(I, J))
			Next 
		Next 
		FileClose(1)
		DelFile(PrgPath & "\" & cOpCodesTblFileName)
		ReDim ConstNames(0)
		ReDim ConstVals(0)
		ReDim Lbls(0)
		ReDim Insts(0)
		ReDim Op1s(0)
		ReDim Op2s(0)
		LblLen = 5
	End Function
	
	Private Function WrtOutFile() As Boolean
		Dim I, PC, J As Integer
		Dim Addrs() As String
		Dim LogFlag As Boolean
		WrtOutFile = True
		I = UBound(Insts)
		ReDim OpCodes(I)
		ReDim Addrs(I)
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
					IncPrgCtr(PC, Len(OpCodes(I)) / 2)
					DecToHex(Op1s(I))
					DecToHex(Op2s(I))
			End Select
		Next 
		If Not WrtOutFile Then
			WrtErrFile()
			Exit Function
		End If
		If LogFlag Then
			WrtLogFile()
			Exit Function
		End If
		WrtBinFile()
		OutFile = OutFile & "." & cOutFileExt
		FileOpen(1, OutFile, OpenMode.Output)
		PrintLine(1, GetHorLine(LblLen + 53))
		PrintLine(1, "ADDRESS MACHINE-CODE  #   LABEL:", TAB(LblLen + 32), "OPCODE  OPERAND")
		PrintLine(1, GetHorLine(LblLen + 53))
		PrintLine(1)
		For I = 1 To UBound(Insts)
			Select Case Insts(I)
				Case "END"
					PrintLine(1, TAB(LblLen + 22), "*****     END     *****")
				Case "EQU"
				Case "ORG"
					PrintLine(1)
				Case Else
					Print(1, Addrs(I) & ":  ")
					If Insts(I) <> "" And Insts(I) <> "DEFS" Then
						J = 1
						Do 
							Print(1, " " & Mid(OpCodes(I), J, 2))
							J = J + 2
						Loop While Mid(OpCodes(I), J, 2) <> ""
					End If
					PrintLine(1, TAB(23), "#", TAB(27), GetCmdLine(I))
			End Select
		Next 
		FileClose(1)
		MsgBox("Assembling process successful.", MsgBoxStyle.Information)
	End Function
	
	Private Function WrtBinFile() As Boolean
		Dim BinPtr, I, J, LenPtr As Integer
		Dim By As Byte
		Dim lngTmp As Integer
		Dim strTmp As String
		WrtBinFile = True
		FileOpen(1, OutFile & "." & cBinFileExt, OpenMode.Binary, OpenAccess.Write)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, cBinFileStartStr, 1)
		LenPtr = Len(cBinFileStartStr) + 1
		BinPtr = LenPtr + 4
		strTmp = FixHexVal(cDefStartAddr, 4)
		By = Val("&H" & Right(strTmp, 2))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, By, BinPtr - 2)
		By = Val("&H" & Left(strTmp, 2))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, By, BinPtr - 1)
		For I = 1 To UBound(Insts)
			Select Case Insts(I)
				Case "", "EQU"
				Case "END"
					Exit For
				Case "ORG"
					strTmp = FixHexVal(Val(Op1s(I)), 4)
					If BinPtr = LenPtr + 4 Then
						By = Val("&H" & Right(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, BinPtr - 2)
						By = Val("&H" & Left(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, BinPtr - 1)
					Else
						lngTmp = BinPtr - LenPtr - 4
						If lngTmp > 65535 Then
							WrtBinFile = False
							FileClose(1)
							DelFile(OutFile & "." & cBinFileExt)
							MsgBox("Can't created '" & OutFile & "." & cBinFileExt & "' file for more than 64kB instructions.", MsgBoxStyle.Critical)
							Exit Function
						End If
						By = Val("&H" & Right(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, BinPtr + 2)
						By = Val("&H" & Left(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, BinPtr + 3)
						strTmp = FixHexVal(lngTmp, 4)
						By = Val("&H" & Right(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, LenPtr)
						By = Val("&H" & Left(strTmp, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, LenPtr + 1)
						LenPtr = BinPtr
						BinPtr = BinPtr + 4
					End If
				Case Else
					For J = 1 To Len(OpCodes(I)) Step 2
						By = Val("&H" & Mid(OpCodes(I), J, 2))
						'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FilePut(1, By, BinPtr)
						BinPtr = BinPtr + 1
					Next 
			End Select
		Next 
		lngTmp = BinPtr - LenPtr - 4
		If lngTmp > 65535 Then
			WrtBinFile = False
			FileClose(1)
			DelFile(OutFile & "." & cBinFileExt)
			MsgBox("Can't created '" & OutFile & "." & cBinFileExt & "' file for more than 64kB instructions.", MsgBoxStyle.Critical)
			Exit Function
		End If
		If lngTmp > 0 Then
			strTmp = FixHexVal(lngTmp, 4)
			By = Val("&H" & Right(strTmp, 2))
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(1, By, LenPtr)
			By = Val("&H" & Left(strTmp, 2))
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(1, By, LenPtr + 1)
		Else
			BinPtr = BinPtr - 4
		End If
		By = 0
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, By, BinPtr)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, By, BinPtr + 1)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(1, cBinFileEndStr, BinPtr + 2)
		FileClose(1)
		MsgBox("Executable file (" & OutFile & "." & cBinFileExt & ") successful created.", MsgBoxStyle.Information)
	End Function
	
	Private Function WrtErrFile() As Object
		Dim I, ErrCtr As Integer
		Dim CmdLine As String
		OutFile = OutFile & "." & cErrFileExt
		FileOpen(1, OutFile, OpenMode.Output)
		PrintLine(1, GetHorLine(LblLen + 53))
		PrintLine(1, "LABEL:", TAB(LblLen + 6), "OPCODE  OPERAND           ; ERROR-DESCRIPTION")
		PrintLine(1, GetHorLine(LblLen + 53))
		PrintLine(1)
		For I = 1 To UBound(Insts)
			CmdLine = GetCmdLine(I)
			If Insts(I) = cErrFlag Then
				ErrCtr = ErrCtr + 1
				CmdLine = CmdLine & " (" & Mid(Str(ErrCtr), 2) & ")"
			End If
			PrintLine(1, CmdLine)
		Next 
		FileClose(1)
		MsgBox(Mid(Str(ErrCtr), 2) & " Error(s) found!", MsgBoxStyle.Critical)
	End Function
	
	Private Function WrtLogFile() As Object
		OutFile = OutFile & "." & cLogFileExt
		FileOpen(1, OutFile, OpenMode.Output)
		PrintLine(1, ";" & GetHorLine(78) & ";")
		PrintLine(1, ";", TAB(80), ";")
		PrintLine(1, ";        Z80 Assembler for ED-Laboratory's Microprocessor Trainer MPT-1", TAB(80), ";")
		PrintLine(1, ";", TAB(80), ";")
		PrintLine(1, ";                         Viva New Technology Protocol", TAB(80), ";")
		PrintLine(1, ";", TAB(80), ";")
		PrintLine(1, ";          Copyright (C) Iskandar Z. Nasibu, Gorontalo, February 2005", TAB(80), ";")
		PrintLine(1, ";", TAB(80), ";")
		PrintLine(1, ";" & GetHorLine(78) & ";")
		PrintLine(1)
		PrintLine(1)
		PrintLine(1, "; There is no instruction to assembling...")
		FileClose(1)
	End Function
	
	Private Function GetHorLine(ByRef ChrNum As Integer) As String
		For ChrNum = 1 To ChrNum
			GetHorLine = GetHorLine & "="
		Next 
	End Function
	
	Private Function FixHexVal(ByRef Value As Integer, Optional ByRef FixVal As Integer = 2) As String
		FixHexVal = Hex(Value)
		Do While Len(FixHexVal) < FixVal
			FixHexVal = "0" & FixHexVal
		Loop 
	End Function
	
	Private Function DecToHex(ByRef Op As String) As Object
		Dim D, H As String
		D = ExtractOp(Op)
		If Val(D) > 0 Then
			H = Hex(Val(D))
			If Val(D) > 9 Then H = H & "H"
			If Left(H, 1) > "9" Then H = "0" & H
			CompressOp(Op, H)
		End If
	End Function
	
	Private Function SetAddrsVal() As Boolean
		Dim InstLen, PC, I As Integer
		SetAddrsVal = True
		PC = cDefStartAddr
		For I = 1 To UBound(Insts)
			Select Case Insts(I)
				Case "END", "EQU", cErrFlag
				Case "ORG"
					PC = Val(Op1s(I))
				Case Else
					If Lbls(I) <> "" Then SetConstVal(Lbls(I), Mid(Str(PC), 2))
					If Insts(I) = cErrFlag Then SetAddrsVal = False
					IncPrgCtr(PC, GetInstLen(I))
			End Select
		Next 
	End Function
	
	Private Function GetInstLen(ByRef LinesPtr As Integer) As Integer
		Dim Ptr As Integer
		Select Case Insts(LinesPtr)
			Case "", "END", "EQU", "ORG", cErrFlag
			Case "DEFS"
				GetInstLen = Val(Op1s(LinesPtr))
			Case Else
				If FindInstFromTbl(LinesPtr, Ptr) Then GetInstLen = Len(InstsTbl(Ptr, 3)) / 2
		End Select
	End Function
	
	Private Function GetInstCode(ByRef LinesPtr As Integer, ByRef PrgCtr As Integer) As String
		Dim Ptr As Integer
		Dim OpIdFlag As Boolean
		Dim Op, Num As String
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
								AddError(GetCmdLine(LinesPtr), "Jump relative address is out of range", LinesPtr)
							Else
								If Val(Num) < 0 Then Num = Str(Val(Num) + 256)
								GetInstCode = Replace(GetInstCode, "d7", FixHexVal(Val("&H" & Right(Hex(Val(Num)), 2))))
							End If
					End Select
					OpIdFlag = Not OpIdFlag
				Loop While OpIdFlag
		End Select
	End Function
	
	Private Function FindInstFromTbl(ByRef LinesPtr As Integer, ByRef Ptr As Integer) As Boolean
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
		AddError(GetCmdLine(LinesPtr), "Illegal instruction", LinesPtr)
	End Function
	
	Private Function IncPrgCtr(ByRef LastVal As Integer, ByRef IncVal As Integer) As Object
		LastVal = LastVal + IncVal
		If LastVal > 65536 Then LastVal = LastVal - 65536
	End Function
	
	Private Function GetOpEqv(ByRef Op As String, ByRef LinesPtr As Integer) As String
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
		Dim I As Integer
		Dim OpFl, OkFl As Boolean
		Dim Op, NumHold As String
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
					AddError(GetCmdLine(I), ErrDesc, I)
					ReplConsts = False
			End Select
NextReplConsts: 
		Next 
	End Function
	
	Private Function ReplChkConst(ByRef Op As String, ByRef LinesPtr As Integer) As Boolean
		Dim Arg, Value As String
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
				CompressOp(Op, Value)
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
						If Mid(Op, 4, 1) = "+" And Val(Arg) > 127 Or Mid(Op, 4, 1) = "-" And Val(Arg) > 128 Then
							ErrDesc = "Overflow displacement type number"
							ReplChkConst = False
						End If
						If Mid(Op, 4, 1) = "-" And Arg = "0" Then Op = Left(Op, 3) & "+0)"
				End Select
		End Select
	End Function
	
	Private Function ExtractConstName(ByRef Op As String) As String
		ExtractConstName = ExtractOp(Op)
		If GetOpType(ExtractConstName) <> cConst Then ExtractConstName = ""
	End Function
	
	Private Function CompressOp(ByRef Op As String, ByRef Arg As String) As Object
		Select Case GetOpId(Op)
			Case cOpIdAbs
				Op = Arg
			Case cOpIdRef
				Op = "(" & Arg & ")"
			Case cOpIdDisRef
				Op = Left(Op, 4) & Arg & ")"
		End Select
	End Function
	
	Private Function ExtractOp(ByRef Op As String) As String
		Select Case GetOpId(Op)
			Case cOpIdAbs
				ExtractOp = Op
			Case cOpIdRef
				ExtractOp = Mid(Op, 2, Len(Op) - 2)
			Case cOpIdDisRef
				ExtractOp = Mid(Op, 5, Len(Op) - 5)
		End Select
	End Function
	
	Private Function GetOpId(ByRef Op As String) As Integer
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
	
	Private Function GetCmdLine(ByRef LinesPtr As Integer) As String
		Dim Tmp As Integer
		If Insts(LinesPtr) = cErrFlag Then
			If Len(Op1s(LinesPtr)) < LblLen + 30 Then GetCmdLine = Space(LblLen + 30 - Len(Op1s(LinesPtr)))
			GetCmdLine = Replace(Op1s(LinesPtr), vbTab, " ") & GetCmdLine & " ; " & Op2s(LinesPtr)
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
		Dim CmdLine As String
		Dim I, LinesPtr, NumArg, ArgType As Integer
		BuildMnemonicsList = True
		FileOpen(1, SrcFile, OpenMode.Input)
		LinesPtr = 1
		ReDim Lbls(LinesPtr)
		ReDim Insts(LinesPtr)
		ReDim Op1s(LinesPtr)
		ReDim Op2s(LinesPtr)
		Do While Not EOF(1)
			CmdLine = LineInput(1)
			If Not GetTrimCmdLine(CmdLine) Then
				AddError(CmdLine, ErrDesc, LinesPtr)
				BuildMnemonicsList = False
				GoTo IncList
			End If
			If UBound(Args) = 0 Then GoTo SkipLine
			If Not GetMnemonic() Then
				AddError(CmdLine, ErrDesc, LinesPtr)
				BuildMnemonicsList = False
				GoTo IncList
			End If
			If Args(1) <> "" Then
				If Not AddConst(Args(1)) Then
					AddError(CmdLine, "Duplicate label name", LinesPtr)
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
						AddError(CmdLine, "Expect one or more operands", LinesPtr)
						BuildMnemonicsList = False
					Else
						For I = 3 To NumArg
							Insts(LinesPtr) = Args(2)
							Op1s(LinesPtr) = Args(I)
							LinesPtr = LinesPtr + 1
							ReDim Preserve Lbls(LinesPtr)
							ReDim Preserve Insts(LinesPtr)
							ReDim Preserve Op1s(LinesPtr)
							ReDim Preserve Op2s(LinesPtr)
						Next 
						GoTo SkipLine
					End If
				Case "DEFS"
					If NumArg = 2 Then
						AddError(CmdLine, "Expected a number or constant as operand", LinesPtr)
						BuildMnemonicsList = False
					Else
						Op1s(LinesPtr) = Args(3)
					End If
				Case "END"
					If Args(1) = "" Then
						GoTo EndBuildMnemonicsList
					Else
						AddError(CmdLine, "Unexpect label name", LinesPtr)
						BuildMnemonicsList = False
					End If
				Case "EQU"
					If Args(1) = "" Then
						AddError(CmdLine, "Expect label name as identifier", LinesPtr)
						BuildMnemonicsList = False
					Else
						If NumArg = 2 Then
							AddError(CmdLine, "Expect a number as operand", LinesPtr)
							BuildMnemonicsList = False
						Else
							SetConstVal(Args(1), Args(3))
							Op1s(LinesPtr) = Args(3)
						End If
					End If
				Case "ORG"
					If Args(1) <> "" Then
						AddError(CmdLine, "Unexpect label name", LinesPtr)
						BuildMnemonicsList = False
					Else
						If NumArg = 2 Then
							AddError(CmdLine, "Expect a number as operand", LinesPtr)
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
			ReDim Preserve Lbls(LinesPtr)
			ReDim Preserve Insts(LinesPtr)
			ReDim Preserve Op1s(LinesPtr)
			ReDim Preserve Op2s(LinesPtr)
SkipLine: 
		Loop 
		LinesPtr = LinesPtr - 1
		ReDim Preserve Lbls(LinesPtr)
		ReDim Preserve Insts(LinesPtr)
		ReDim Preserve Op1s(LinesPtr)
		ReDim Preserve Op2s(LinesPtr)
EndBuildMnemonicsList: 
		FileClose(1)
		ReDim Args(0)
	End Function
	
	Private Function AddError(ByRef CmdLine As String, ByRef ErrDescription As String, ByRef LinesPtr As Integer) As Object
		Insts(LinesPtr) = cErrFlag
		Op1s(LinesPtr) = CmdLine
		Op2s(LinesPtr) = ErrDescription
	End Function
	
	Private Function GetFileName(ByRef FullName As String) As String
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetBaseName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFileName = fso.GetBaseName(FullName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	Private Function GetFileExt(ByRef FullName As String) As String
		Dim fso As Object
		fso = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.GetExtensionName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFileExt = fso.GetExtensionName(FullName)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing
	End Function
	
	Private Function AddConst(ByRef Name As String, Optional ByRef Value As String = "") As Boolean
		Dim I As Integer
		If Len(Name) > LblLen Then LblLen = Len(Name)
		If UBound(ConstNames) = 0 And ConstNames(0) = "" Then
			ConstNames(0) = Name
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(Value) Then ConstVals(0) = Value
		Else
			For I = 0 To UBound(ConstNames)
				If Name = ConstNames(I) Then Exit Function
			Next 
			ReDim Preserve ConstNames(I)
			ReDim Preserve ConstVals(I)
			ConstNames(I) = Name
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(Value) Then ConstVals(I) = Value
		End If
		AddConst = True
	End Function
	
	Private Function GetConstVal(ByRef Name As String, ByRef Value As String) As Boolean
		Dim I As Integer
		For I = 0 To UBound(ConstNames)
			If Name = ConstNames(I) Then
				Value = ConstVals(I)
				GetConstVal = True
				Exit Function
			End If
		Next 
	End Function
	
	Private Function SetConstVal(ByRef Name As String, ByRef Value As String) As Object
		Dim I As Integer
		For I = 0 To UBound(ConstNames)
			If Name = ConstNames(I) Then
				ConstVals(I) = Value
				Exit Function
			End If
		Next 
	End Function
	
	Private Function GetMnemonic() As Boolean
		Dim Tmp() As String
		Dim aChr As String
		Dim TmpPtr, ArgsPtr, I As Integer
		Dim InParenth, InQuo As Boolean
		Dim ArgsTop, ArgType As Integer
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
					If Tmp(2) <> "JP" And (Tmp(TmpPtr) = "(IX" Or Tmp(TmpPtr) = "(IY") Then Tmp(TmpPtr) = Left(Tmp(TmpPtr), 3) & "+0"
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
					If Tmp(TmpPtr) = "" Or I = ArgsTop Or Tmp(2) = "EQU" Or Tmp(2) = "ORG" Or Tmp(2) = "DEFS" Or Tmp(2) <> "DEFB" And Tmp(2) <> "DEFW" And TmpPtr > 3 Then
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
								Tmp(TmpPtr) = Tmp(TmpPtr) & Mid(Str(Val(Args(I)) + 256 * Val(Args(I + 1))), 2)
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
		Args = VB6.CopyArray(Tmp)
		ErrDesc = ""
		Exit Function
FailGetMnemonic: 
		GetMnemonic = False
		ReDim Args(0)
	End Function
	
	Private Function GetNumType(ByRef Num As String) As Integer
		Dim aChr As String
		Dim NumLen, NumHold, I As Integer
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
						If aChr < "0" Or aChr > "9" And aChr < "A" Or aChr > "F" Then Exit Function
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
	
	Private Function GetArgType(ByRef Arg As String) As Integer
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
	
	Private Function GetOpType(ByRef Op As String) As Integer
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
	
	Private Function FindDir(ByRef Arg As String) As Boolean
		FindDir = FindArg(Arg, DirsList)
	End Function
	
	Private Function FindInst(ByRef Arg As String) As Boolean
		FindInst = FindArg(Arg, InstsList)
	End Function
	
	Private Function FindReg(ByRef Arg As String) As Boolean
		FindReg = FindArg(Arg, RegsList)
	End Function
	
	Private Function FindRefReg(ByRef Arg As String) As Boolean
		FindRefReg = FindArg(Arg, RefRegsList)
	End Function
	
	Private Function FindFlagId(ByRef Arg As String) As Boolean
		FindFlagId = FindArg(Arg, FlagIdsList)
	End Function
	
	Private Function FindArg(ByRef Arg As String, ByRef List As Object) As Boolean
		Dim I As Integer
		For I = 0 To UBound(List)
			'UPGRADE_WARNING: Couldn't resolve default property of object List(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Arg = List(I) Then
				FindArg = True
				Exit Function
			End If
		Next 
	End Function
	
	Private Function GetTrimCmdLine(ByRef CmdLine As String) As Boolean
		Dim CmdLnLen, NumArgs, I As Integer
		Dim aChr As String
		Dim InQuo, InArg, InParenth As Boolean
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
						If Asc(aChr) >= &H61s And Asc(aChr) <= &H7As Then aChr = Chr(Asc(aChr) - &H20s)
						If aChr >= "0" And aChr <= "9" Or aChr >= "A" And aChr <= "Z" Or aChr = "_" Then
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
End Module