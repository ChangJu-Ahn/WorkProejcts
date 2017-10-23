<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1411ma1
'*  4. Program Name         : 발주형태구성 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2005/08/08
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kim Duk Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003/05/20
'*        
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'=======================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'=======================================================================================================================
Const BIZ_PGM_ID = "m1411mb1.asp"     

Dim C_PotypeCd
Dim C_PotypeNm
'2005.01.24	Multicompany 추가 
Dim C_InterComFlg

Dim C_StoCond		'STO 여부 
Dim C_ImportCond
Dim C_BlCond
Dim C_CcCond
Dim C_GmCond
Dim C_IvCond
Dim C_RetCond
Dim C_SubCond		'외주가공여부(<- 사급여부)
Dim C_GmTypeCd		'입출고형태(<- 입고형태)
Dim C_GmTypePop
Dim C_GmTypeNm
Dim C_RitypeCd		'사급형태(<- 출고형태)
Dim C_RitypePop
Dim C_RitypeNm
Dim C_IvtypeCd
Dim C_IvtypePop
Dim C_IvtypeNm
Dim C_SoTypeCd		'수주형태 
Dim C_SoTypePop		'수주형태 팝업 
Dim C_SoTypeNm		'수주형태명 
Dim C_Useflg

Dim Actionflg
Dim lgQuery
Dim lgCopyRow
Dim IsOpenPop          

'=======================================================================================================================
Sub initSpreadPosVariables()  
	C_PotypeCd	  = 1
	C_PotypeNm    = 2
	'멀티컴퍼니거래여부 추가 
	C_InterComFlg = 3
	C_StoCond	  = 4
	C_ImportCond  = 5
	C_BlCond      = 6
	C_CcCond      = 7
	C_GmCond      = 8
	C_IvCond      = 9 
	C_RetCond     = 10
	C_SubCond     = 11
	C_GmTypeCd    = 12
	C_GmTypePop   = 13
	C_GmTypeNm    = 14
	C_RitypeCd    = 15 
	C_RitypePop   = 16
	C_RitypeNm    = 17
	C_IvtypeCd    = 18
	C_IvtypePop   = 19
	C_IvtypeNm    = 20
	C_SoTypeCd		= 21
	C_SoTypePop		= 22
	C_SoTypeNm		= 23
	C_Useflg		= 24
End Sub
'=======================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE    
    lgBlnFlgChgValue = False     
    lgIntGrpCount = 0            
    lgStrPrevKey = ""            
    lgLngCurRows = 0             
End Sub
'=======================================================================================================================
 Sub SetDefaultVal()
	frm1.rdoUseflg(0).Checked = true
	frm1.txtPotypeCd.focus
	Set gActiveElement = document.activeElement
	Call SetToolbar("1110110100101111")      
End Sub
'=======================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'=======================================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20051201",,parent.gAllowDragDropSpread    
		
	With frm1.vspdData
 
		.ReDraw = false
		.MaxCols = C_Useflg+1
		.Col = .MaxCols 
    	.MaxRows = 0
    	
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_PotypeCd, "발주형태", 10,,,5,2
		ggoSpread.SSSetEdit		C_PotypeNm, "발주형태명", 20,,,50
		'멀티컴퍼니거래여부 추가 
		ggoSpread.SSSetCheck	C_InterComFlg, "멀티컴퍼니거래여부", 20,,,true
		ggoSpread.SSSetCheck	C_StoCond,"STO여부",10,,,true
		ggoSpread.SSSetCheck	C_ImportCond,"수입여부",10,,,true
		ggoSpread.SSSetCheck	C_BlCond,"선적여부",10,,,true
		ggoSpread.SSSetCheck	C_CcCond,"통관여부",10,,,true
		ggoSpread.SSSetCheck	C_GmCond,"입고여부",10,,,true
		ggoSpread.SSSetCheck	C_IvCond,"매입여부",10,,,true
		ggoSpread.SSSetCheck	C_RetCond,"반품여부",10,,,true
		ggoSpread.SSSetCheck	C_SubCond,"외주가공여부",15,,,true
		ggoSpread.SSSetEdit		C_GmtypeCd, "입출고형태", 15,,,5,2
		ggoSpread.SSSetButton	C_GmtypePop
		ggoSpread.SSSetEdit		C_GmtypeNm, "입출고형태명", 20
		ggoSpread.SSSetEdit		C_RitypeCd, "사급형태", 10,,,5,2
		ggoSpread.SSSetButton	C_RitypePop
		ggoSpread.SSSetEdit		C_RitypeNm, "사급형태명", 20
		ggoSpread.SSSetEdit		C_IvtypeCd, "매입형태",10,,,5,2
		ggoSpread.SSSetButton	C_IvtypePop
		ggoSpread.SSSetEdit		C_IvtypeNm, "매입형태명",20
		ggoSpread.SSSetEdit		C_SoTypeCd, "수주형태",10,,,5,2
		ggoSpread.SSSetButton	C_SoTypePop
		ggoSpread.SSSetEdit		C_SoTypeNm, "수주형태명",20
		ggoSpread.SSSetCheck	C_Useflg,"사용여부",10,,,true
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		
		Call ggoSpread.MakePairsColumn(C_GmtypeCd, C_GmtypePop)
		Call ggoSpread.MakePairsColumn(C_RitypeCd, C_RitypePop)
		Call ggoSpread.MakePairsColumn(C_IvtypeCd, C_IvtypePop)
		Call ggoSpread.MakePairsColumn(C_SoTypeCd, C_SoTypePop)
		
		'Call ggoSpread.SSSetColHidden(C_StoCond, C_StoCond, True)
		'Call ggoSpread.SSSetColHidden(C_SoTypeCd, C_SoTypeNm, True)
		
		Call ggoSpread.SSSetSplit2(2)
		Call SetSpreadLock 
    
		.ReDraw = true
 
    End With
    
End Sub
'=======================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PotypeCd	  = iCurColumnPos(1)
			C_PotypeNm    = iCurColumnPos(2)
			C_InterComFlg = iCurColumnPos(3)
			C_StoCond	  = iCurColumnPos(4)
			C_ImportCond  = iCurColumnPos(5)
			C_BlCond      = iCurColumnPos(6)
			C_CcCond      = iCurColumnPos(7)
			C_GmCond      = iCurColumnPos(8)
			C_IvCond      = iCurColumnPos(9) 
			C_RetCond     = iCurColumnPos(10)
			C_SubCond     = iCurColumnPos(11)
			C_GmTypeCd    = iCurColumnPos(12)
			C_GmTypePop   = iCurColumnPos(13)
			C_GmTypeNm    = iCurColumnPos(14)
			C_RitypeCd    = iCurColumnPos(15) 
			C_RitypePop   = iCurColumnPos(16)
			C_RitypeNm    = iCurColumnPos(17)
			C_IvtypeCd    = iCurColumnPos(18)
			C_IvtypePop   = iCurColumnPos(19)
			C_IvtypeNm    = iCurColumnPos(20)
			C_SoTypeCd	  = iCurColumnPos(21)
			C_SoTypePop   = iCurColumnPos(22)
			C_SoTypeNm	  = iCurColumnPos(23)
			C_Useflg      = iCurColumnPos(24)
    End Select    
End Sub
'=======================================================================================================================
Sub SetSpreadLock()
    
    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
	With ggoSpread
    	.spreadunlock  C_PotypeCd, -1
		.sssetrequired C_PotypeCd, -1
		.sssetrequired C_PotypeNm, -1
		.spreadunlock  C_ImportCond, -1
		.sssetrequired C_GmtypeCd, -1, -1
		.spreadlock    C_GmtypeNm, -1, C_GmtypeNm, -1
		.spreadunlock  C_RitypeCd, -1
'		.sssetrequired C_RitypeCd, -1, -1
		.spreadlock    C_RitypeNm, -1
		.spreadunlock  C_IvtypeCd, -1
		.sssetrequired C_IvtypeCd, -1, -1
		.spreadlock    C_IvtypeNm, -1,  C_IvtypeNm, -1
		
		.SSSetProtected frm1.vspdData.MaxCols, -1
		
    End With
    frm1.vspdData.ReDraw = True
End Sub
'=======================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, Byval pvEndRow)
	ggoSpread.Source = frm1.vspdData
    
    With frm1.vspdData
		.ReDraw = False
		
		ggoSpread.spreadunlock   C_PotypeCd, pvStartRow, C_PotypeCd,pvEndRow
		ggoSpread.sssetrequired  C_PotypeCd, pvStartRow,  pvEndRow
		ggoSpread.sssetrequired  C_PotypeNm, pvStartRow,  pvEndRow
		ggoSpread.sssetrequired  C_GmtypeCd, pvStartRow,  pvEndRow
		ggoSpread.spreadunlock	 C_SubCond, pvStartRow, C_SubCond, pvEndRow
    
		'ggoSpread.SSSetProtected C_RitypeCd, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_RitypePop, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RitypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BlCond, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CcCond, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_SubCond, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_GmtypeCd, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected C_GmtypePop, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_GmtypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IvtypeCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IvtypePop, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_IvtypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoTypeCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoTypePop, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoTypeNm, pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStartRow, pvEndRow
		
		.ReDraw = True
    End With
    
End Sub
'=======================================================================================================================
Function OpenPotype()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
			 
	arrParam(0) = "발주형태" 
	arrParam(1) = "M_CONFIG_PROCESS"    
	arrParam(2) = UCase(Trim(frm1.txtPotypeCd.Value))
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "발주형태" 
			 
	arrField(0) = "po_type_cd" 
	arrField(1) = "po_type_Nm" 
				   
	arrHeader(0) = "발주형태" 
	arrHeader(1) = "발주형태명"  
				   
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPoTypeCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoTypeCd.Value = arrRet(0)
		frm1.txtPoTypeNm.Value = arrRet(1)
		frm1.txtPoTypeCd.focus	
		Set gActiveElement = document.activeElement
	End If 
 
End Function
'=======================================================================================================================
Function OpenGmtype()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = "입출고형태" 
	arrParam(1) = "m_mvmt_type"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_GmTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = ""
	'arrParam(4) = "rcpt_flg=" & FilterVar("Y", "''", "S") & " "
	If Trim(GetSpreadText(frm1.vspdData,C_GmCond,iCurRow,"X","X")) = "1" then
		arrParam(4) = "rcpt_flg=" & FilterVar("Y", "''", "S") & " "
	Else
		arrParam(4) = "rcpt_flg=" & FilterVar("N", "''", "S") & " "
	End if
	
	If Trim(GetSpreadText(frm1.vspdData,C_ImportCond,iCurRow,"X","X")) = "1" then
		arrParam(4) = arrParam(4) & "And import_flg=" & FilterVar("Y", "''", "S") & " "
	Else
		arrParam(4) = arrParam(4) & "And import_flg=" & FilterVar("N", "''", "S") & " "
	End if
	 
	If Trim(GetSpreadText(frm1.vspdData,C_RetCond,iCurRow,"X","X")) = "1" then 
		arrParam(4) = arrParam(4) & "And ret_flg=" & FilterVar("Y", "''", "S") & " "
	Else
		arrParam(4) = arrParam(4) & "And ret_flg=" & FilterVar("N", "''", "S") & " "
	End if
	 
	If Trim(GetSpreadText(frm1.vspdData,C_SubCond,iCurRow,"X","X")) = "1" then 
		arrParam(4) = arrParam(4) & "And subcontra2_flg=" & FilterVar("Y", "''", "S") & " "
	Else
		arrParam(4) = arrParam(4) & "And subcontra2_flg=" & FilterVar("N", "''", "S") & " "
	End if

	arrParam(5) =  "입출고형태" 
	 
	arrField(0) = "io_type_cd" 
	arrField(1) = "io_type_nm" 
		   
	arrHeader(0) =  "입출고형태" 
	arrHeader(1) =  "입출고형태명" 
		   
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_GmTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_GmTypeNm,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If 
 
End Function
'=======================================================================================================================
Function OpenRitype()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = "사급형태" 
	arrParam(1) = "m_mvmt_type"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_RitypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = "subcontra_flg=" & FilterVar("Y", "''", "S") & " "
	 
	arrParam(5) = "사급형태" 
	 
	arrField(0) = "io_type_cd" 
	arrField(1) = "io_type_nm" 
	    
	arrHeader(0) = "사급형태" 
	arrHeader(1) = "사급형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_RitypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_RitypeNm,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If 
 
End Function
'=======================================================================================================================
Function OpenIvtype()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)
 Dim iCurRow
 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = "매입형태" 
	arrParam(1) = "m_iv_type"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_IvTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	 
	If Trim(GetSpreadText(frm1.vspdData,C_ImportCond,iCurRow,"X","X")) = "1" then 
		arrParam(4) = "import_flg=" & FilterVar("Y", "''", "S") & " "
	Else
		arrParam(4) = "import_flg=" & FilterVar("N", "''", "S") & " "
	End if
	 
	arrParam(5) = "매입형태" 
	 
	arrField(0) = "iv_type_cd" 
	arrField(1) = "iv_type_nm" 
	    
	arrHeader(0) = "매입형태" 
	arrHeader(1) = "매입형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_IvTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_IvTypeNm,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If 
 
End Function
'=======================================================================================================================
Function OpenChdGmType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	 
	arrParam(0) = "자품목출고형태" 
	arrParam(1) = "b_minor"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_ChdGmTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = "major_cd = " & FilterVar("m5101", "''", "S") & "" 
	arrParam(5) = "자품목출고형태"
	 
	arrField(0) = "minor_cd" 
	arrField(1) = "minor_nm" 
	    
	arrHeader(0) = "자품목출고형태"
	arrHeader(1) = "자품목출고형태명"
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_ChdGmTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_ChdGmTypeNm,	iCurRow, arrRet(1))
	End If 
 
End Function
'======================================================================================================================='---------------------------------------------------------------------------------------------------------
Function OpenSoType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	iCurRow = frm1.vspdData.ActiveRow
	 
	arrParam(0) = "수주형태" 
	arrParam(1) = "s_so_type_config"    
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_SoTypeCd,iCurRow,"X","X")))
	arrParam(3) = ""
	arrParam(4) = "sto_flag = " & FilterVar("Y", "''", "S") & " and export_flag=" & FilterVar("N", "''", "S") & " and usage_flag=" & FilterVar("Y", "''", "S") & "  "
 
	arrParam(5) = "수주형태" 
	 
	arrField(0) = "so_type" 
	arrField(1) = "so_type_nm" 
	    
	arrHeader(0) = "수주형태" 
	arrHeader(1) = "수주형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_SoTypeCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_SoTypeNm,	iCurRow, arrRet(1))
		Call vspdData_Change(0, iCurRow)
	End If 
 
End Function
'=======================================================================================================================
'	Name : changeImport()
'	Description : 수입여부 Check Event
'=======================================================================================================================
 sub changeImport(ByVal curRow, ByVal Actionflg)

	With frm1
		
		If Trim(GetSpreadText(frm1.vspdData,C_ImportCond,curRow,"X","X")) = "1" Then
			ggoSpread.spreadunlock  C_BlCond, curRow, C_CcCond, curRow

			If Actionflg = True Then
				Call .vspdData.SetText(C_BlCond,	curRow, "1")
				Call .vspdData.SetText(C_CcCond,	curRow, "1")
			End If
		else
			ggoSpread.SSSetProtected C_BlCond, curRow, curRow
			ggoSpread.SSSetProtected C_CcCond, curRow, curRow
			
			If Actionflg = True Then
				Call .vspdData.SetText(C_BlCond,	curRow, "0")
				Call .vspdData.SetText(C_CcCond,	curRow, "0")
			End If
		End If		
  	End With
End Sub
'=======================================================================================================================
'	Name : changeIv()
'	Description : 매입여부 Check Event
'=======================================================================================================================
sub changeIv(ByVal curRow, ByVal Actionflg)
	With frm1
		If Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "1" Then
			ggoSpread.spreadunlock  C_IvtypeCd, curRow, C_IvtypePop, curRow
			ggoSpread.sssetrequired C_IvtypeCd, curRow, curRow
		else
			ggoSpread.SSSetProtected C_IvtypeCd, curRow, curRow
			ggoSpread.SSSetProtected C_IvtypePop, curRow, curRow
			
			If Actionflg = True then
				Call .vspdData.SetText(C_IvtypeCd,	curRow, "")
				Call .vspdData.SetText(C_IvtypeNm,	curRow, "")
			End If
		End if
		
		'Call changeRet(curRow, Actionflg)
		
	End With
End Sub
'=======================================================================================================================
'	Name : changeGm()
'	Description : 입고여부 Check Event (입출고 통합으로 삭제함)
'=======================================================================================================================
'Sub changeGm(ByVal curRow, ByVal Actionflg)
'
'	With frm1
'		if Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" then
'	
'			ggoSpread.spreadunlock  C_SubCond, curRow, C_SubCond, curRow
'			ggoSpread.spreadunlock  C_GmTypeCd, curRow, C_GmTypePop, curRow
'			ggoSpread.sssetrequired C_GmTypeCd, curRow, curRow
'			
'			if Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "1" then
'				ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow
'				ggoSpread.sssetrequired C_RitypeCd, curRow, curRow
'			end if
'		else
'			ggoSpread.SSSetProtected C_SubCond, curRow, curRow
'			ggoSpread.SSSetProtected C_GmTypeCd, curRow, curRow
'			ggoSpread.SSSetProtected C_GmTypePop, curRow, curRow
'			
'			'ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow 20040226 김지현 
'			'ggoSpread.sssetrequired C_RitypeCd, curRow, curRow 20040226 김지현 
'			ggoSpread.SSSetProtected C_RitypeCd, curRow, curRow
'			ggoSpread.SSSetProtected C_RitypePop, curRow, curRow
'			
'			if Actionflg = true then
'				Call .vspdData.SetText(C_GmTypeCd,	curRow, "")
'				Call .vspdData.SetText(C_GmTypeNm,	curRow, "")
'				Call .vspdData.SetText(C_RitypeCd,	curRow, "")
'				Call .vspdData.SetText(C_RitypeNm,	curRow, "")
'				Call .vspdData.SetText(C_SubCond,	curRow, "0")
'			end if
'		end if
'	End with
'End Sub
'=======================================================================================================================
'	Name : changeRet()
'	Description : 반품여부 Check Event (삭제함)
'=======================================================================================================================
'Sub changeRet(ByVal curRow, ByVal Actionflg)
'
'	with frm1
'		if Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "1" then
'	
'			ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow
'			
'			if Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "0" then
'				ggoSpread.sssetrequired C_RitypeCd, curRow, curRow    'refund & return : issue type-> optional
'			end if
'			      
'			ggoSpread.SSSetProtected C_SubCond, curRow, curRow 
'
'			if Actionflg = true then
'				Call .vspdData.SetText(C_SubCond,	curRow, "0")
'			end if
'		
'		elseif Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "0" then
'			
'			If Trim(GetSpreadText(frm1.vspdData,C_StoCond,curRow,"X","X")) = "1" then	'Sto가 체크되면 반품여부 미체크되더라도 사급여부는 무조건 Protect
'				ggoSpread.spreadlock  C_SubCond, curRow, C_SubCond, curRow
'				ggoSpread.SSSetProtected C_SubCond, curRow, curRow
'			Else
'				ggoSpread.spreadunlock  C_SubCond, curRow, C_SubCond, curRow
'			End if
'			
'			ggoSpread.SSSetProtected C_RitypeCd, curRow, curRow
'			ggoSpread.SSSetProtected C_RitypePop, curRow, curRow
'
'			if Actionflg = true then
'				Call .vspdData.SetText(C_RitypeCd,	curRow, "")
'				Call .vspdData.SetText(C_RitypeNm,	curRow, "")
'			end if
'		end if
'	End with
'End Sub
'=======================================================================================================================
'	Name : changeSub()
'	Description : 사급여부 Check Event (삭제함)
'=======================================================================================================================
'Sub changeSub(ByVal curRow, ByVal Actionflg)
'
'	With frm1
'		if Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "1" then
'			ggoSpread.SSSetProtected C_RetCond, curRow, curRow
'			
'			If Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" then
'				ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow
'				ggoSpread.sssetrequired C_RitypeCd, curRow, curRow
'			end if
'			
'			if Actionflg = true then
'				Call .vspdData.SetText(C_RetCond,	curRow, "")
'			end if
'		
'		elseif Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "0" then
'			ggoSpread.spreadunlock  C_RetCond, curRow, C_RetCond, curRow
'			
'			if Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" then
'	
'				ggoSpread.SSSetProtected C_RitypeCd, curRow, curRow
'				ggoSpread.SSSetProtected C_RitypePop, curRow, curRow
'				
'				if Actionflg = true then
'
'					Call .vspdData.SetText(C_RitypeCd,	curRow, "")
'					Call .vspdData.SetText(C_RitypeNm,	curRow, "")
'				end if
'				ggoSpread.SSSetProtected C_RitypePop, curRow, curRow
'			       
'			else
'				if Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "1" then '20040226 C_SubCond ->C_RetCond
'					ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow
'					
'					if Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "0" then
'						ggoSpread.sssetrequired C_RitypeCd, curRow, curRow    'refund & return : issue type-> optional
'					end if
'				end if 
'			end if
'		end if
'	End with
' 
'End Sub
'=======================================================================================================================
'	Name : changeSto()
'	Description : STO여부 Check Event
'=======================================================================================================================
Sub changeSto(ByVal curRow, ByVal Actionflg)

	With frm1
		if Trim(GetSpreadText(frm1.vspdData,C_StoCond,curRow,"X","X")) = "1" then
		
			ggoSpread.SSSetProtected C_ImportCond, curRow, curRow					'수입여부 
			Call .vspdData.SetText(C_ImportCond,	curRow, "0")
			ggoSpread.SSSetProtected C_BlCond, curRow, curRow						'선적여부 
			Call .vspdData.SetText(C_BlCond,	curRow, "0")						
			ggoSpread.SSSetProtected C_CcCond, curRow, curRow						'통관여부 
			Call .vspdData.SetText(C_CcCond,	curRow, "0")
			ggoSpread.SSSetProtected C_IvCond, curRow, curRow						'매입여부 
			Call .vspdData.SetText(C_IvCond,	curRow, "0")
			ggoSpread.SSSetProtected C_SubCond, curRow, curRow						'외주가공여부(사급여부)
			Call .vspdData.SetText(C_SubCond,	curRow, "0")
			'멀티컴퍼니거래여부 추가 
'			ggoSpread.SSSetProtected C_InterComFlg, curRow, curRow						'멀티컴퍼니거래여부 
'			Call .vspdData.SetText(C_InterComFlg,	curRow, "0")
			If Trim(GetSpreadText(frm1.vspdData, C_InterComFlg, curRow, "X","x")) = "1" then
				Call DisplayMsgBox("17A013","x",frm1.vspdData.Row,"멀티컴퍼니거래여부")  
				Call .vspdData.SetText(C_StoCond,	curRow, "0")  
			End If					
				
			ggoSpread.SpreadUnLock  C_GmCond, curRow, C_GmCond, curRow			'입고여부 
			ggoSpread.SpreadUnLock  C_RetCond, curRow, C_RetCond, curRow		'반품여부 
			ggoSpread.SpreadUnLock  C_SoTypeCd, curRow, C_SoTypePop, curRow		'수주형태 
			ggoSpread.SSSetRequired C_SoTypeCd, curRow, curRow
		Else
			ggoSpread.SSSetProtected C_SoTypeCd, curRow, curRow
			ggoSpread.SSSetProtected C_SoTypePop, curRow, curRow
			ggoSpread.SpreadUnLock  C_ImportCond, curRow, C_ImportCond, curRow	'수입여부 
			ggoSpread.SpreadUnLock  C_IvCond, curRow, C_IvCond, curRow			'매입여부 
			ggoSpread.SpreadUnLock  C_SubCond, curRow, C_SubCond, curRow		'외주가공여부(사급여부)
			'멀티컴퍼니거래여부 추가			
			'ggoSpread.SpreadUnLock  C_InterComFlg, curRow, C_IvCond, curRow
			ggoSpread.SpreadUnLock  C_InterComFlg, curRow, C_InterComFlg, curRow
						
			Call .vspdData.SetText(C_SoTypeCd,	curRow, "")
			Call .vspdData.SetText(C_SoTypeNm,	curRow, "")
		
		'	Call changeGm(curRow, Actionflg)
		'    Call changeRet(curRow, Actionflg)
		End if
		ggoSpread.SSSetProtected C_SoTypeNm, curRow, curRow
	End With
End Sub
'=======================================================================================================================
'	Name : changeInterComFlg()
'	Description : 멀티컴퍼니여부 Check Event
'=======================================================================================================================
Sub changeInterComFlg(ByVal curRow, ByVal Actionflg)
'	With frm1                                                                           
'		if Trim(GetSpreadText(frm1.vspdData,C_InterComFlg,curRow,"X","X")) = "1" then   
'			ggoSpread.SSSetProtected C_StoCond, curRow, curRow					'STO여부 
'			Call .vspdData.SetText(C_StoCond,	curRow, "0")                            
'		Else                                                                            
'			ggoSpread.SpreadUnLock  C_StoCond, curRow, C_ImportCond, curRow				
'							                                                            
'		End if                                                                          
'	End With                                                                            
	With frm1                                                                           
		if Trim(GetSpreadText(frm1.vspdData,C_InterComFlg,curRow,"X","X")) = "1" then   
'			ggoSpread.SSSetProtected C_StoCond, curRow, curRow					'STO여부 
			If Trim(GetSpreadText(frm1.vspdData, C_StoCond, curRow, "X","x")) = "1" then
				Call DisplayMsgBox("17A013","x",frm1.vspdData.Row,"STO여부")  
				Call .vspdData.SetText(C_InterComFlg,	curRow, "0")  
			End If			
		Else                                                                            
			'ggoSpread.SpreadUnLock  C_StoCond, curRow, C_ImportCond, curRow
							                                                            
		End if                                                                          
	End With
End Sub                                                     

'=======================================================================================================================
'	Name : changeGmType()
'	Description : 입출고 유형 Check Event
'=======================================================================================================================
Sub changeGmType(ByVal Row)
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strSubcontra_flg

	With frm1.vspdData
		.Row = Row
		.Col = C_GmTypeCd

		Call CommonQueryRs(" IO_TYPE_NM, SUBCONTRA2_FLG ", " M_MVMT_TYPE ", " IO_TYPE_CD = " & FilterVar(.Text, "''", "S") & " AND USAGE_FLG = " & FilterVar("Y", "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If Err.number <> 0 Then
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
	
		If Len(lgF0) > 0 Then
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			If Trim(GetSpreadText(frm1.vspdData,C_SubCond,Row,"X","X")) = "1" Then
				strSubcontra_flg = "Y"
			Else
				strSubcontra_flg = "N"
			End If

			If strSubcontra_flg = Trim(lgF1(0)) Then
				.Col 	= C_GmTypeNm
				.Text 	= Trim(lgF0(0))
			Else
				Call DisplayMsgBox("171930", "X", "X", "X")
				.Col 	= C_GmTypeCd
				.Text 	= ""
				.Col 	= C_GmTypeNm
				.Text 	= ""
			End If
		Else
			Call DisplayMsgBox("171900", "X", .Text, "X")
			.Text 	= ""
			.Col 	= C_GmTypeNm
			.Text 	= ""
		End If
		
	End With
End Sub

'=======================================================================================================================
'	Name : setGmCell()
'	Description : ????????
'=======================================================================================================================
'Sub setGmCell(ByVal curRow, ByVal Actionflg)
'
'	With frm1
'	  
'		If (Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) <> "1" And Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "1") Then   '입고 = N and 반품 = Y 인 경우 출고형태를 Enable시킨다.
'			ggoSpread.spreadunlock  C_RitypeCd, curRow, C_RitypePop, curRow
'			ggoSpread.sssetrequired C_RitypeCd, curRow, curRow
'		Else
'			ggoSpread.SSSetProtected C_RitypeCd, curRow, curRow
'			ggoSpread.SSSetProtected C_RitypePop, curRow, curRow
'			If Actionflg = True then
'				Call .vspdData.SetText(C_RitypeCd,	curRow, "")
'				Call .vspdData.SetText(C_RitypeNm,	curRow, "")
'			End If
'		End If
'	End with
'End Sub
'=======================================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")                       
    
    Call InitSpreadSheet                                        
    Call SetDefaultVal
    Call InitVariables                                    

End Sub
'=======================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
	   Exit Sub
	End If
	   	    
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If
End Sub
'=======================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'=======================================================================================================================
Sub vspdData_MouseDown(ByVal Button , ByVal Shift , ByVal x , ByVal y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'=======================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'=======================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=======================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	ggoSpread.SSSetProtected C_PotypeCd , -1	
	Call ggoSpread.ReOrderingSpreadData()
	Call RestoreColor()
End Sub
'=======================================================================================================================
Sub RestoreColor()
	Dim i
	With frm1
		.vspdData.ReDraw = False
		For i = 1 To .vspdData.MaxRows
			.vspdData.Row = i
			
			'changeSto --
			If Trim(GetSpreadText(frm1.vspdData,C_StoCond,curRow,"X","X")) = "1" Then
				
				ggoSpread.SSSetProtected C_ImportCond, i, i						'수입여부 
				Call .vspdData.SetText(C_ImportCond,	i, "0")
				ggoSpread.SSSetProtected C_BlCond, i, i							'선적여부 
				Call .vspdData.SetText(C_BlCond,	i, "0")
				ggoSpread.SSSetProtected C_CcCond, i, i							'통관여부 
				Call .vspdData.SetText(C_CcCond,	i, "0")
				ggoSpread.SSSetProtected C_IvCond, i, i							'매입여부 
				Call .vspdData.SetText(C_IvCond,	i, "0")
				ggoSpread.SSSetProtected C_SubCond, i, i						'외주가공여부(사급여부)
				call .vspdData.SetText(C_SubCond,	i, "0")
					
				ggoSpread.SpreadUnLock  C_GmCond, i, C_GmCond, i		'입고여부 
				ggoSpread.SpreadUnLock  C_RetCond, i, C_RetCond, i		'반품여부 
				ggoSpread.SpreadUnLock  C_SoTypeCd, i, C_SoTypePop, i	'수주형태 
				ggoSpread.SSSetRequired C_SoTypeCd, i, i
			Else
				ggoSpread.SSSetProtected C_SoTypeCd, i, i
				ggoSpread.SSSetProtected C_SoTypePop, i, i
				ggoSpread.SpreadUnLock  C_ImportCond, i, C_ImportCond, i	
				ggoSpread.SpreadUnLock  C_IvCond, i, C_IvCond, i
				ggoSpread.SpreadUnLock  C_SubCond, i, C_SubCond, i
				
				Call .vspdData.SetText(C_SoTypeCd,	i, "")
				Call .vspdData.SetText(C_SoTypeNm,	i, "")
				
			End if
			ggoSpread.SSSetProtected C_SoTypeNm, i, i
			
			'changeImport --
			If Trim(GetSpreadText(frm1.vspdData,C_ImportCond,curRow,"X","X")) = "1" Then
				ggoSpread.spreadunlock  C_BlCond, i, C_CcCond, i
						
			Else
				ggoSpread.SSSetProtected C_BlCond, i, i
				ggoSpread.SSSetProtected C_CcCond, i, i
				If Actionflg = True Then
					Call .vspdData.SetText(C_BlCond,	i, "0")
					Call .vspdData.SetText(C_CcCond,	i, "0")
				End If
			End If
			
			'changeIv --
			If Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "1" Then
				ggoSpread.spreadunlock  C_IvtypeCd, i, C_IvtypePop, i
				ggoSpread.sssetrequired C_IvtypeCd, i, i
			Else
				ggoSpread.SSSetProtected C_IvtypeCd, i, i
				ggoSpread.SSSetProtected C_IvtypePop, i, i
				
				If Actionflg = True then
					Call .vspdData.SetText(C_IvtypeCd,	i, "")
					Call .vspdData.SetText(C_IvtypeNm,	i, "")
				End If
			End if
			
			'changeRet --
'			If Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "1" Then
'				ggoSpread.spreadunlock  C_RitypeCd, i, C_RitypePop, i
'				
'				If Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "0" Then
'					ggoSpread.sssetrequired C_RitypeCd, i, i    'refund & return : issue type-> optional
'				End If
'					      
'				ggoSpread.SSSetProtected C_SubCond, i, i    
'				If Actionflg = True then
'					Call .vspdData.SetText(C_SubCond,	i, "0")
'				End if
'			Elseif Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "0" Then
'				
'				If Trim(GetSpreadText(frm1.vspdData,C_StoCond,curRow,"X","X")) = "1" Then	'Sto가 체크되면 반품여부 미체크되더라도 사급여부는 무조건 Protect
'					ggoSpread.spreadlock  C_SubCond, i, C_SubCond, i
'					ggoSpread.SSSetProtected C_SubCond, i, i
'				Else
'					ggoSpread.spreadunlock  C_SubCond, i, C_SubCond, i
'				End If
'					
'				ggoSpread.SSSetProtected C_RitypeCd, i, i
'				ggoSpread.SSSetProtected C_RitypePop, i, i
'				If Actionflg = True Then
'					Call .vspdData.SetText(C_RitypeCd,	i, "")
'					Call .vspdData.SetText(C_RitypeNm,	i, "")
'				End if
'			End if
			
			'changeGm --
'			If Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" Then
'				ggoSpread.spreadunlock  C_SubCond, i, C_SubCond, i
'				ggoSpread.spreadunlock  C_GmTypeCd, i, C_GmTypePop, i
'				ggoSpread.sssetrequired C_GmTypeCd, i, i
'				
'				If Trim(GetSpreadText(frm1.vspdData,C_StoCond,curRow,"X","X")) = "1" Then
'					ggoSpread.spreadunlock  C_RitypeCd, i, C_RitypePop, i
'					ggoSpread.sssetrequired C_RitypeCd, i, i
'				End If
'			Else
'				ggoSpread.SSSetProtected C_SubCond, i, i
'				ggoSpread.SSSetProtected C_GmTypeCd, i, i
'				ggoSpread.SSSetProtected C_GmTypePop, i, i
'							
'				ggoSpread.spreadunlock  C_RitypeCd, i, C_RitypePop, i
'				ggoSpread.sssetrequired C_RitypeCd, i, i
'				ggoSpread.SSSetProtected C_RitypeCd, i, i
'				ggoSpread.SSSetProtected C_RitypePop, i, i
'				
'				If Actionflg = True Then
'					Call .vspdData.SetText(C_GmTypeCd,	i, "")
'					Call .vspdData.SetText(C_GmTypeNm,	i, "")
'					Call .vspdData.SetText(C_RitypeCd,	i, "")
'					Call .vspdData.SetText(C_RitypeNm,	i, "")
'					Call .vspdData.SetText(C_SubCond,	i, "0")
'				End If
'				
'			End If
			
			'changeSub --
'			If Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "1" Then
'				ggoSpread.SSSetProtected C_RetCond, i, i
'				
'				If Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" Then
'					ggoSpread.spreadunlock  C_RitypeCd, i, C_RitypePop, i
'					ggoSpread.sssetrequired C_RitypeCd, i, i
'				End If
'				
'				If Actionflg = True Then
'					Call .vspdData.SetText(C_RetCond,	i, "")
'				End If
'		
'			Elseif Trim(GetSpreadText(frm1.vspdData,C_SubCond,curRow,"X","X")) = "0" Then
'				ggoSpread.spreadunlock  C_RetCond, i, C_RetCond, i
'				
'				If Trim(GetSpreadText(frm1.vspdData,C_GmCond,curRow,"X","X")) = "1" Then
'					ggoSpread.SSSetProtected C_RitypeCd, i, i
'					ggoSpread.SSSetProtected C_RitypePop, i, i
'					
'					If Actionflg = True Then
'						Call .vspdData.SetText(C_RitypeCd,	i, "")
'						Call .vspdData.SetText(C_RitypeNm,	i, "")
'					End If
'					ggoSpread.SSSetProtected C_RitypePop, i, i
'				       
'				Else
'					If Trim(GetSpreadText(frm1.vspdData,C_RetCond,curRow,"X","X")) = "1" Then
'						ggoSpread.spreadunlock  C_RitypeCd, i, C_RitypePop, i
'						
'						If Trim(GetSpreadText(frm1.vspdData,C_IvCond,curRow,"X","X")) = "0" Then
'							ggoSpread.sssetrequired C_RitypeCd, i, i    'refund & return : issue type-> optional
'						End if
'
'					End if 
'				End if
'			End if
		Next
	.vspdData.ReDraw = True
	End With
	
End Sub
'=======================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'=======================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.source= frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

 	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        
 
 	frm1.vspdData.ReDraw = False
 
	Select Case Col
		Case C_ImportCond 
			Call changeImport(Row, true)
		Case C_IvCond
			Call changeIv(frm1.vspdData.ActiveRow, true)
		Case C_StoCond
			Call changeSto(frm1.vspdData.ActiveRow, true)
		'멀티컴퍼니거래여부 추가			
		Case C_InterComFlg
			Call changeInterComFlg(frm1.vspdData.ActiveRow, true)
		Case C_GmTypeCd
			Call changeGmType(Row)
	End Select 

	frm1.vspdData.ReDraw = True

End Sub

'=======================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
	If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End if
End Sub
'=======================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	If lgQuery = True then Exit Sub
	If lgCopyRow = True then Exit Sub
 
	frm1.vspdData.ReDraw = False
 
	Select Case Col
		Case C_ImportCond 
			Call changeImport(Row, true)
		Case C_IvCond
			Call changeIv(frm1.vspdData.ActiveRow, true)
		' 입출고 통합으로 삭제함 
		'Case C_GmCond
		'	Call changeGm(frm1.vspdData.ActiveRow, true)
		' 반품여부 Event 삭제함 
		'Case C_RetCond
		'	Call changeRet(frm1.vspdData.ActiveRow, true)
		' 사급여부 Event 삭제함 
		'Case C_SubCond
		'	Call changeSub(frm1.vspdData.ActiveRow, true)
		Case C_StoCond
			Call changeSto(frm1.vspdData.ActiveRow, true)
		'멀티컴퍼니거래여부 추가			
		Case C_InterComFlg
			Call changeInterComFlg(frm1.vspdData.ActiveRow, true)			
	End Select 
 
	frm1.vspdData.ReDraw = True
	 
	if Col = C_GmTypePop then
		Call OpenGmType()
	elseif Col = C_RitypePop then
		Call OpenRitype()
	elseif Col = C_IvTypePop then
		Call OpenIvType()
	elseif Col = C_SoTypePop then
		Call OpenSoType()
	End if
 
End Sub
'=======================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'=======================================================================================================================
Function FncQuery()
	Dim IntRetCD 
	    
	FncQuery = False                                        
	    
	Err.Clear                                               
	 
	ggoSpread.Source = frm1.vspdData
	    
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	Call ggoOper.ClearField(Document, "2")     
	Call InitVariables
	
	If Not ChkField(Document, "1") Then      
		Exit Function
	End If
	        
	If DbQuery = False Then Exit Function
	       
	FncQuery = True           
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    On Error Resume Next                                   
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call SetDefaultVal
    Call InitVariables
        
    FncNew = True                                           
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear                                               
    On Error Resume Next                              
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    
    If Not ggoSpread.SSDefaultCheck Then 
       Exit Function
    End If

    If DbSave = False Then Exit Function
    
    FncSave = True                                                       
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncCopy()  
	frm1.vspdData.ReDraw = False
	If frm1.vspdData.Maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData 
	lgCopyRow = true
	ggoSpread.CopyRow
	    
	Call changeImport(frm1.vspdData.ActiveRow, false)
	Call changeIv(frm1.vspdData.ActiveRow, false)
	'입출고 통합으로 삭제함 
	'Call changeGm(frm1.vspdData.ActiveRow, false)
	'사급여부 삭제함 
	'Call changeSub(frm1.vspdData.ActiveRow, false)  
	Call changeSto(frm1.vspdData.ActiveRow, false)
	'멀티컴퍼니거래여부 추가 
	Call changeInterComFlg(frm1.vspdData.ActiveRow, false)
	 
	ggoSpread.spreadunlock  C_PotypeCd, frm1.vspdData.ActiveRow, C_PotypeNm,frm1.vspdData.ActiveRow
	ggoSpread.sssetrequired C_PotypeCd, frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow
	ggoSpread.sssetrequired C_PotypeNm, frm1.vspdData.ActiveRow,  frm1.vspdData.ActiveRow
	ggoSpread.SSSetProtected C_GmtypeNm, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	ggoSpread.SSSetProtected C_RitypeNm, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	ggoSpread.SSSetProtected C_IvtypeNm, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	 
	Call frm1.vspdData.SetText(C_PotypeCd,	frm1.vspdData.ActiveRow, "")
	    
	frm1.vspdData.ReDraw = True
	lgCopyRow = False
	Set gActiveElement = document.ActiveElement         
End Function
'=======================================================================================================================
Function FncCancel() 
	If frm1.vspdData.Maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo 
	Set gActiveElement = document.ActiveElement                                                    
End Function
'=======================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	
	Dim IntRetCD
    Dim imRow, iRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End if
    End If
    
    With frm1
		.vspdData.ReDraw = False
 		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			Call .vspdData.SetText(C_Useflg,	iRow, "1")
		Next
		.vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
    
    If Err.number = 0 Then 
		FncInsertRow = True                                                          '☜: Processing is OK
    End If
End Function
'=======================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    if frm1.vspdData.Maxrows < 1 then exit function
    
    With frm1.vspdData 
    	.focus
		ggoSpread.Source = frm1.vspdData 
        lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncPrev() 
	ggoSpread.Source = frm1.vspdData
    On Error Resume Next                                                
End Function
'=======================================================================================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(parent.C_MULTI)          
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(parent.C_MULTI , False)   
    Set gActiveElement = document.ActiveElement                                
End Function
'=======================================================================================================================
Function FncExit()

 Dim IntRetCD
 
	FncExit = False
 
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")           
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal
    
    Err.Clear
    
    DbQuery = False
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    With frm1
	    
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPotypeCd=" & .hdnPotype.value
			strVal = strVal & "&txtUseflg=" & .hdnUseflg.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPotypeCd=" & Trim(.txtPotypeCd.value)
			
			If .rdoUseflg(0).checked = True then
				strVal = strVal & "&txtUseflg=" & ""
			Elseif .rdoUseflg(1).checked = True then
				strVal = strVal & "&txtUseflg=" & "Y"
			Else
				strVal = strVal & "&txtUseflg=" & "N"
			End if
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		end if 
		    
		Call RunMyBizASP(MyBizASP, strVal)
	        
	End With
    
    DbQuery = True
    
End Function
'=======================================================================================================================
Function DbQueryOk()             
 
 Dim index
	
	lgIntFlgMode = parent.OPMD_UMODE           
	        
	Call ggoOper.LockField(Document, "Q")        
	Call SetToolbar("1110111100111111")
	
	Call RemovedivTextArea    
	
	frm1.vspdData.ReDraw = False
	ggoSpread.spreadlock  C_PotypeCd, 1,C_PotypeCd,frm1.vspdData.MaxRows
	 
	For index = 1 To frm1.vspdData.MaxRows
		Call changeImport(index, false)
		Call changeIv(index, false)
		'입출고 통합으로 삭제함 
		'Call changeGm(index, false)
		'사급여부 이벤트 삭제함 
		'Call changeSub(index, false)
		Call changeSto(index, false)	'added for STO
		'멀티컴퍼니거래여부 추가 
		Call changeInterComFlg(index, false)
	Next
	
	lgQuery = False
	frm1.vspdData.ReDraw = True
      
End Function
'=======================================================================================================================
Function DbSave() 
	Dim lRow        
	Dim strVal, strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
	
	DbSave = False                                                      
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If 
    
    iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
    
	With frm1
		.txtMode.value = parent.UID_M0002
		
		strVal = ""
	    strDel = ""
	    
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0
        	Select Case .vspdData.Text
				Case ggoSpread.InsertFlag
					strVal = "C"																				& iColSep				
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_PotypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_PotypeNm,lRow, "X","X"))				& iColSep
					'멀티컴퍼니거래여부 추가 
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_InterComFlg,lRow, "X","X"))			& iColSep
					
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_StoCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_ImportCond,lRow, "X","X"))			& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_BlCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_CcCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_GmCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_IvCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_RetCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_SubCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_GmTypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_RitypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_IvtypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_SoTypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_Useflg,lRow, "X","X"))				& iColSep
					strVal = strVal & lRow & iRowSep
				Case ggoSpread.UpdateFlag	
					strVal = "U"																				& iColSep				
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_PotypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_PotypeNm,lRow, "X","X"))				& iColSep
					'멀티컴퍼니거래여부 추가 
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_InterComFlg,lRow, "X","X"))			& iColSep					
					
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_StoCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_ImportCond,lRow, "X","X"))			& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_BlCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_CcCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_GmCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_IvCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_RetCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_SubCond,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_GmTypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_RitypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_IvtypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_SoTypeCd,lRow, "X","X"))				& iColSep
					strVal = strVal & UCase(GetSpreadText(frm1.vspdData,C_Useflg,lRow, "X","X"))				& iColSep
					strVal = strVal & lRow & iRowSep
				Case ggoSpread.DeleteFlag
					strDel = "D"																				& iColSep				
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_PotypeCd,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_PotypeNm,lRow, "X","X"))				& iColSep
					'멀티컴퍼니거래여부 추가 
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_InterComFlg,lRow, "X","X"))			& iColSep					
					
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_StoCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_ImportCond,lRow, "X","X"))			& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_BlCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_CcCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_GmCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_IvCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_RetCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_SubCond,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_GmTypeCd,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_RitypeCd,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_IvtypeCd,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_SoTypeCd,lRow, "X","X"))				& iColSep
					strDel = strDel & UCase(GetSpreadText(frm1.vspdData,C_Useflg,lRow, "X","X"))				& iColSep
					strDel = strDel & lRow & iRowSep
			End Select
			
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 

			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
				 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
				       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 

			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
				          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
				       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
				         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select
        Next
	End With

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   
	'------ Developer Coding part (End ) -------------------------------------------------------------- 

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)
				
    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement 
   
End Function
'=======================================================================================================================
Function DbSaveOk()          

	Call InitVariables
	frm1.vspdData.MaxRows = 0
	
	Call MainQuery()
 
End Function
'=======================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'=======================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" --> 
</HEAD>
<!-- '#########################################################################################################
'            6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주형태</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPotypeCd" ALT="발주형태" SIZE=10 MAXLENGTH=5  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPotype()">
										<INPUT TYPE=TEXT ID="txtPotypeNm" ALT="발주형태" NAME="arrCond" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>사용여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg1" Value="A" checked tag="1X"><label for="rdoUseflg1">&nbsp;전체&nbsp;</label>
										<INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg2" Value="Y" tag="1X"><label for="rdoUseflg2">&nbsp;사용&nbsp;</label>
										<INPUT TYPE=radio Class="Radio" ALT="사용여부" NAME="rdoUseflg" id = "rdoUseflg3" Value="N" tag="1X"><label for="rdoUseflg3">&nbsp;미사용&nbsp;</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUseflg" tag="24">
<P ID="divTextArea"></P>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
