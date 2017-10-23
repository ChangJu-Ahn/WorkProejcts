<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s1911ma1.asp																*
'*  4. Program Name         : 수주형태등록																*
'*  5. Program Desc         : 수주형태등록																*
'*  6. Comproxy List        :  																			*
'*  7. Modified date(First) : 2000/08/25																*
'*  8. Modified date(Last)  : 2005/01/24																*
'*  9. Modifier (First)     : Juvenile	 																*
'* 10. Modifier (Last)      : Sim Hae Young															*
'* 11. Comment              : 수주형태명 뒤에 멀티컴퍼니거래여부 칼럼추가(2005/01/24)																			*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s1911mb1.asp"
Const gstrDNTypeMajor = "I0001"

Dim S_SoType
Dim S_SoTypeNm

Dim S_InterComFlg

Dim S_StoFlg
Dim S_ExportFlg
Dim S_RetItemFlg
Dim S_CiFlg
Dim S_RelDnFlg
Dim S_AutoDnFlg
Dim S_RelBillFlg
Dim S_SpStkFlg
Dim S_PostDnType
Dim S_PostDNTypePopup
Dim S_PostDNTypeNm
Dim S_PostBillType
Dim S_PostBillTypePopup
Dim S_PostBillTypeNm
Dim S_DlvyLt
Dim S_SoMgmtFlg
Dim S_CreditChkFlg
Dim S_DepositChkFlg
Dim S_UsageFlg
Dim S_IndReqmtNo

Dim lsBtnClickFlag
Dim lsBtnProtectedRow

Dim lsQueryMode
Dim lsFncCopyFlag

Dim IsOpenPop

'=================================================================================================================
Sub initSpreadPosVariables()

	S_SoType			= 1		'수주형태 
	S_SoTypeNm			= 2		'수주형태명 

	S_InterComFlg			= 3		'멀티컴퍼니거래여부 

	S_StoFlg			= 4		'STO여부 
	S_ExportFlg			= 5		'수출여부 
	S_RetItemFlg		= 6		'반품여부 
	S_CiFlg				= 7		'통관여부 
	S_RelDnFlg			= 8		'출하여부 
	S_AutoDnFlg			= 9		'자동출하여부 
	S_RelBillFlg		= 10		'매출여부 
	S_SpStkFlg			= 11	'위탁여부 
	S_PostDnType		= 12	'출하타입 
	S_PostDNTypePopup	= 13	'출하타입팝업버튼 
	S_PostDNTypeNm		= 14	'출하타입명 
	S_PostBillType		= 15	'매출타입 
	S_PostBillTypePopup	= 16	'매출타입팝업버튼 
	S_PostBillTypeNm	= 17	'매출타입명 
	S_DlvyLt			= 18	'납기일수 
	S_SoMgmtFlg			= 19    '수주관리여부 
	S_CreditChkFlg		= 20    '여신관리여부 
	S_DepositChkFlg		= 21    '적립금적용여부 
	S_UsageFlg			= 22	'사용여부 
	S_IndReqmtNo		= 23	'Sort시 필요한 히든 칼럼 
End Sub

'=================================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lsFncCopyFlag = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
End Sub

'=================================================================================================================
Sub SetDefaultVal()
	frm1.txtSOType.focus
	Set gActiveElement = document.activeElement
End Sub

'=================================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'=================================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20050419",,parent.gAllowDragDropSpread

		.ReDraw = false

	    .MaxCols = S_IndReqmtNo									'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols											'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
	    .MaxRows = 0

	    Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit S_SOType, "수주형태", 12,,, 4, 2
	    	ggoSpread.SSSetEdit S_SoTypeNm, "수주형태명", 30,,, 50

		ggoSpread.SSSetCheck S_InterComFlg, "멀티컴퍼니거래여부", 15,,,true

		ggoSpread.SSSetCheck S_StoFlg, "STO여부", 12,,,true
		ggoSpread.SSSetCheck S_RetItemFlg, "반품여부", 12,,,true
		ggoSpread.SSSetCheck S_ExportFlg, "수출여부", 12,,,true
		ggoSpread.SSSetCheck S_CIFlg, "통관여부", 12,,,true
	    ggoSpread.SSSetCheck S_RelDNFlg, "출하여부", 12,,,true
	    ggoSpread.SSSetCheck S_AutoDNFlg, "자동출하생성여부", 20,,,true
	    ggoSpread.SSSetEdit S_PostDNType, "출하형태", 12,,,3, 2
		ggoSpread.SSSetButton S_PostDNTypePopup
		ggoSpread.SSSetEdit S_PostDNTypeNm, "출하형태명", 20, 0
	    ggoSpread.SSSetCheck S_RelBillFlg, "매출여부", 12,,,true
	    ggoSpread.SSSetEdit S_PostBillType, "매출채권형태", 12,,, 20, 2
		ggoSpread.SSSetButton S_PostBillTypePopup
		ggoSpread.SSSetEdit S_PostBillTypeNm, "매출채권형태명", 20, 0
		ggoSpread.SSSetCheck S_SpStkFlg, "위탁여부", 12,,,true
		ggoSpread.SSSetCheck S_SoMgmtFlg, "수주관리여부", 12,,,true
		ggoSpread.SSSetCheck S_CreditChkFlg, "여신관리여부", 12,,,true
		ggoSpread.SSSetCheck S_DepositChkFlg, "적립금적용여부", 18,,,true
		ggoSpread.SSSetCheck S_UsageFlg, "사용여부", 12,,,true
		Call AppendNumberPlace("7","3","0")
		ggoSpread.SSSetFloat S_DlvyLt,"납기일수" ,12,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

		Call ggoSpread.MakePairsColumn(S_PostDNType,S_PostDNTypePopup)
		Call ggoSpread.MakePairsColumn(S_PostBillType,S_PostBillTypePopup)
		Call ggoSpread.SSSetColHidden(S_SpStkFlg,S_SpStkFlg,True)
	    Call SetSpreadLock("", 0, -1, "")
	    .ReDraw = True
    End With
End Sub

'=================================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False

		ggoSpread.SSSetrequired S_SoType, lRow, -1
		ggoSpread.SSSetrequired S_SoTypeNm, lRow, -1
		ggoSpread.SpreadUnLock S_RelDNFlg, lRow, -1
		ggoSpread.SpreadLock S_PostDNTypeNm,lRow, -1
		ggoSpread.SpreadUnLock S_PostBillType, lRow, -1
		ggoSpread.SpreadLock S_PostBillTypeNm, lRow, -1

		ggoSpread.SpreadUnLock S_InterComFlg, lRow, -1

		ggoSpread.SpreadUnLock S_StoFlg, lRow, -1
		ggoSpread.SpreadUnLock S_UsageFlg, lRow, -1
		ggoSpread.SpreadUnLock S_SoMgmtFlg, lRow, -1
		ggoSpread.SpreadUnLock S_CreditChkFlg, lRow, -1
		ggoSpread.SpreadUnLock S_DepositChkFlg, lRow, -1
		ggoSpread.SpreadUnLock S_DlvyLt, lRow, -1

		.vspdData.ReDraw = True
    End With
End Sub

'=================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		ggoSpread.SSSetProtected S_SoType, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired S_SoTypeNm, pvStartRow,pvEndRow
		ggoSpread.SSSetProtected S_PostDNTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostBillTypeNm, pvStartRow, pvEndRow
    End With
End Sub

'=================================================================================================================
Sub SetInsertSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		ggoSpread.SSSetRequired S_SoType, pvStartRow,pvEndRow
		ggoSpread.SSSetRequired S_SoTypeNm, pvStartRow,pvEndRow
		ggoSpread.SSSetProtected S_PostBillType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostBillTypePopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostBillTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostDNType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostDNTypePopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_PostDNTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_AutoDNFlg, pvStartRow, pvEndRow

		ggoSpread.SpreadUnLock S_CIFlg, pvStartRow, S_CIFlg, pvEndRow
		ggoSpread.SpreadUnLock S_SoMgmtFlg, pvStartRow, S_SoMgmtFlg, pvEndRow
		ggoSpread.SpreadUnLock S_CreditChkFlg, pvStartRow, S_CreditChkFlg, pvEndRow
		ggoSpread.SpreadUnLock S_DepositChkFlg, pvStartRow, S_DepositChkFlg, pvEndRow
		ggoSpread.SpreadUnLock S_DlvyLt, pvStartRow, S_DlvyLt, pvEndRow
    End With
End Sub

'=================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				S_SoType			= iCurColumnPos(1)
				S_SoTypeNm			= iCurColumnPos(2)		'수주형태명 

				S_InterComFlg  			= iCurColumnPos(3)		'멀티컴퍼니거래여부 

				S_StoFlg  			= iCurColumnPos(4)		'Sto여부 
				S_ExportFlg			= iCurColumnPos(5)		'수출여부 
				S_RetItemFlg		= iCurColumnPos(6)		'반품여부 
				S_CiFlg				= iCurColumnPos(7)		'통관여부 
				S_RelDnFlg			= iCurColumnPos(8)		'출하여부 
				S_AutoDnFlg			= iCurColumnPos(9)		'자동출하여부 
				S_RelBillFlg		= iCurColumnPos(10)		'매출여부 
				S_SpStkFlg			= iCurColumnPos(11)		'위탁여부 
				S_PostDnType		= iCurColumnPos(12)	'출하타입 
				S_PostDNTypePopup	= iCurColumnPos(13)	'출하타입팝업버튼 
				S_PostDNTypeNm		= iCurColumnPos(14)	'출하타입명 
				S_PostBillType		= iCurColumnPos(15)	'매출타입 
				S_PostBillTypePopup	= iCurColumnPos(16)	'매출타입팝업버튼 
				S_PostBillTypeNm	= iCurColumnPos(17)	'매출타입명 
				S_DlvyLt			= iCurColumnPos(18)	'납기일수 
				S_SoMgmtFlg			= iCurColumnPos(19)    '수주관리여부 
				S_CreditChkFlg		= iCurColumnPos(20)    '여신관리여부 
				S_DepositChkFlg		= iCurColumnPos(21)    '적립금적용여부 
				S_UsageFlg			= iCurColumnPos(22)	'사용여부 
				S_IndReqmtNo		= iCurColumnPos(23)	'Sort시 필요한 히든 칼럼 
    End Select
End Sub

'=================================================================================================================
Sub SetQuerySpreadColor()
	With frm1
	    ggoSpread.SSSetProtected S_SoType
		ggoSpread.SSSetProtected S_SoTypeNm
    End With
End Sub

'=================================================================================================================
Function OpenCondtionPopup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "수주형태"
	arrParam(1) = "S_SO_TYPE_CONFIG"
	arrParam(2) = Trim(frm1.txtSoType.value)
	arrParam(4) = ""
	arrParam(5) = "수주형태"

	arrField(0) = "SO_type"
	arrField(1) = "SO_TYPE_NM"

	arrHeader(0) = "수주형태"
	arrHeader(1) = "수주형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCondtionPopup(arrRet)
	End If

End Function

'=================================================================================================================
Function OpenMinorCd(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "출하형태"
	arrParam(1) = "b_minor a, I_MOVETYPE_CONFIGURATION  b"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & "  OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & "  AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & "  "	' Where Condition

	arrParam(5) = "출하형태"

	arrField(0) = "a.Minor_CD"
	arrField(1) = "a.Minor_NM"

	arrHeader(0) = "출하형태"
	arrHeader(1) = "출하형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCD(arrRet)
	End If

End Function

'=================================================================================================================
Function OpenTypePopup(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim ExportFlg
	If IsOpenPop = True Then Exit Function


    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = S_ExportFlg

	If frm1.vspdData.Text = "1" Then
		ExportFlg = "Y"
	Else
		ExportFlg = "N"
	End If

	IsOpenPop = True

	arrParam(0) = "매출채권형태"
	arrParam(1) = "S_BILL_TYPE_CONFIG"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "EXCEPT_FLAG = " & FilterVar("N", "''", "S") & "  and EXPORT_FLAG = " & FilterVar(ExportFlg, "''", "S") & ""
	arrParam(5) = "매출채권형태"

	arrField(0) = "BILL_TYPE"
	arrField(1) = "BILL_TYPE_NM"

	arrHeader(0) = "매출채권형태"
	arrHeader(1) = "매출채권형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTypePopup(arrRet)
	End If

End Function

'=================================================================================================================
Function SetCondtionPopup(Byval arrRet)
	frm1.txtSOType.value = arrRet(0)
	frm1.txtSOTypeNm.value = arrRet(1)
	frm1.txtSOType.focus
End Function

'=================================================================================================================
Function SetTypePopup(Byval arrRet)
	With frm1
		.vspdData.Col = S_PostBillType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = S_PostBillTypeNm
		.vspdData.Text = arrRet(1)

		Call vspdData_Change(.vspdData.Col, .vspdData.Row)
	End With

	lgBlnFlgChgValue = True
End Function

'=================================================================================================================
Function SetMinorCD(Byval arrRet)
	With frm1
		.vspdData.Col = S_PostDNType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = S_PostDNTypeNm
		.vspdData.Text = arrRet(1)

		Call vspdData_Change(.vspdData.Col, .vspdData.Row)
	End With

	lgBlnFlgChgValue = True
End Function

'=================================================================================================================
Function SetRadio()
	frm1.rdoUsageFlgAll.checked = True
End Function

'=================================================================================================================
Function SetDnRelGrid(Byval Col, Byval Row)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.Col = Col
		.Row = Row

		If .TEXT = "0" then
			.Col = S_AutoDnFlg
			.Row = Row
			.Text = "0"
			ggoSpread.SpreadLock S_AutoDnFlg, Row, Col ,Row
		Else
			ggoSpread.SpreadUnLock S_AutoDnFlg, Row, Col ,Row
		End If

	End With
End Function

'=================================================================================================================
Function SetCCRelGrid(Byval Col, Byval Row)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.Col = Col
		.Row = Row

		If .TEXT = "1" then
			.Col = S_RelDnFlg
			.Row = Row
			.Text = "1"
			ggoSpread.SpreadLock S_RelDnFlg, Row, Col ,Row
		Else
			.Col = S_ExportFlg
			If .Text = "1" Then
				.Col = S_RelDnFlg
				.Text = "0"
				ggoSpread.SpreadLock S_RelDnFlg, Row, Col ,Row
			Else
				ggoSpread.SpreadUnLock S_RelDnFlg, Row, Col ,Row
			End If
		End If

	End With
End Function

'=================================================================================================================
Function OnChgRelBillFlg(strCheck, ByVal lRow)
	OnChgRelBillFlg = False

	With frm1.vspdData

		Select Case strCheck
		Case "BillUnCheck"

			.Col = S_PostBillType
			.Row = lRow
			.Text = ""

			.Col = S_PostBillTypeNm
			.Row = lRow
			.Text = ""

			ggoSpread.SSSetProtected	S_PostBillType, lRow, lRow
			ggoSpread.SSSetProtected	S_PostBillTypePopup, lRow, lRow
			ggoSpread.SSSetProtected	S_PostBillTypeNm, lRow, lRow

		Case "BillCheck"
			ggoSpread.Spreadunlock 		S_PostBillType, lRow, S_PostBillTypePopup, lRow
			ggoSpread.SSSetrequired		S_PostBillType, lRow, lRow

		Case "DnUnCheck"
			.Col = S_PostDnType
			.Row = lRow
			.Text = ""

			.Col = S_PostDNTypeNm
			.Row = lRow
			.Text = ""
			ggoSpread.SSSetProtected	S_PostDnType, lRow, lRow
			ggoSpread.SSSetProtected	S_PostDNTypePopup, lRow, lRow
			ggoSpread.SSSetProtected	S_PostDNTypeNm, lRow, lRow

		Case "DnCheck"
			ggoSpread.Spreadunlock 		S_PostDnType, lRow, S_PostDNTypePopup, lRow
			ggoSpread.SSSetrequired		S_PostDnType, lRow, lRow

		End Select

	End With
	OnChgRelBillFlg = True

End Function

'=================================================================================================================
Function SetInterRelGrid(Byval Col, Byval Row)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.Col = Col
		.Row = Row

		If .Text = "" Then .Text = "0"

		If .TEXT = "0" Then	'해제시 

		Else			'체크시 //STO여부 해제후 체크가능합니다.
			.Col = S_StoFlg
			.Row = Row
			If .TEXT = "1" Then
		        	Call DisplayMsgBox("17A013","x",frm1.vspdData.Row,"STO여부")
		        	.Col = Col
				.Row = Row
		        	.Text = "0"

			        Exit Function
			End If
		End If
	End With
End Function
'=================================================================================================================
Function SetStoRelGrid(Byval Col, Byval Row)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.Col = Col
		.Row = Row

		If .Text = "" Then .Text = "0"

		If .TEXT = "0" then
			.Col = S_ExportFlg
			.Row = Row
			.Text = "0"
			.Col = S_CiFlg
			.Row = Row
			.Text = "0"
			.Col = S_RelDnFlg
			.Row = Row
			.Text = "0"
			.Col = S_RelBillFlg
			.Row = Row
			.Text = "0"
			.Col = S_SoMgmtFlg
			.Row = Row
			.Text = "0"
			.Col = S_CreditChkFlg
			.Row = Row
			.Text = "0"
			.Col = S_DepositChkFlg
			.Row = Row
			.Text = "0"

			ggoSpread.SpreadUnLock S_ExportFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_CiFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_RelDnFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_RelBillFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_SoMgmtFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_CreditChkFlg, Row, Col ,Row
			ggoSpread.SpreadUnLock S_DepositChkFlg, Row, Col ,Row

		Else
			.Col = S_InterComFlg
			.Row = Row
			If .TEXT = "1" Then	'체크시 //멀티컴퍼니거래여부 해제후 체크가능합니다.
		        	Call DisplayMsgBox("17A013","x",frm1.vspdData.Row,"멀티컴퍼니거래여부")
		        	.Col = Col
				.Row = Row
		        	.Text = "0"
			        Exit Function
			End If

			.Col = S_ExportFlg
			.Row = Row
			.Text = "0"
			.Col = S_CiFlg
			.Row = Row
			.Text = "0"
			.Col = S_RelDnFlg
			.Row = Row
			.Text = "1"
			.Col = S_RelBillFlg
			.Row = Row
			.Text = "0"
			.Col = S_SoMgmtFlg
			.Row = Row
			.Text = "1"
			.Col = S_CreditChkFlg
			.Row = Row
			.Text = "0"
			.Col = S_DepositChkFlg
			.Row = Row
			.Text = "0"

			ggoSpread.SSSetProtected S_ExportFlg, Row, Row
			ggoSpread.SSSetProtected S_CiFlg, Row, Row
			ggoSpread.SSSetProtected S_RelDnFlg, Row, Row
			ggoSpread.SSSetProtected S_RelBillFlg, Row, Row
			ggoSpread.SSSetProtected S_SoMgmtFlg, Row, Row
			ggoSpread.SSSetProtected S_CreditChkFlg, Row, Row
			ggoSpread.SSSetProtected S_DepositChkFlg, Row, Row
		End If
	End With
End Function

'=================================================================================================================
Function SetExportRelGrid(Byval Col, Byval Row)
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.Col = Col
		.Row = Row

		If .TEXT = "0" then

			.Col = S_CiFlg
			If .Text = "" Then .Text = "0"
			If .Text = "0" Then
				ggoSpread.SpreadUnLock S_RelDnFlg, Row, Col ,Row
			End If
			ggoSpread.SpreadUnLock S_RetItemFlg, Row, Col ,Row
		Else
			.Col = S_RetItemFlg
			.Text = "0"

			.Col = S_RelDnFlg
			If .Text = "1" Then
				.Col = S_CiFlg
				.Text = "1"
			End If

			ggoSpread.SpreadLock S_RetItemFlg, Row, Col ,Row
			ggoSpread.SpreadLock S_RelDnFlg, Row, Col ,Row
		End If
	End With
End Function


'=================================================================================================================
Function SetCopyRow()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetRequired S_SoType, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
	ggoSpread.SSSetRequired S_SoTypeNm, frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
	ggoSpread.SSSetProtected S_PostDNTypeNm, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	ggoSpread.SSSetProtected S_PostBillTypeNm, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	With frm1.vspdData

		Call SetInterRelGrid(S_InterComFlg, .ActiveRow)	'멀티컴퍼니거래여부 

		Call SetStoRelGrid(S_StoFlg, .ActiveRow)
		Call SetExportRelGrid(S_ExportFlg, .ActiveRow)
		Call SetDnRelGrid(S_RelDnFlg, .ActiveRow)
		.Col = S_RelDnFlg
		.Row = .ActiveRow

		If .Text = "1" Then
			.Col = S_PostBillType
			.Row = .ActiveRow
			.Text = ""

			.Col = S_PostBillTypeNm
			.Row = .ActiveRow
			.Text = ""

			ggoSpread.SSSetProtected	S_PostBillType, .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected	S_PostBillTypePopup, .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected	S_PostBillTypeNm, .ActiveRow, .ActiveRow
		Else
			ggoSpread.Spreadunlock 	S_PostBillType, .ActiveRow, S_PostBillTypePopup, .ActiveRow
			ggoSpread.SSSetrequired S_PostBillType, .ActiveRow, .ActiveRow
		End If

		.Col = S_RelBillFlg
		.Row = .ActiveRow

		If .Text = "1" Then
			.Col = S_PostDnType
			.Row = .ActiveRow
			.Text = ""

			.Col = S_PostDNTypeNm
			.Row = .ActiveRow
			.Text = ""

			ggoSpread.SSSetProtected	S_PostDnType, .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected	S_PostDNTypePopup, .ActiveRow, .ActiveRow
			ggoSpread.SSSetProtected	S_PostDNTypeNm, .ActiveRow, .ActiveRow

		Else
			ggoSpread.Spreadunlock 	S_PostDnType, .ActiveRow, S_PostDNTypePopup, .ActiveRow
			ggoSpread.SSSetrequired S_PostDnType, .ActiveRow, .ActiveRow
		End If

	End With
	frm1.vspdData.ReDraw = True
	lsFncCopyFlag = False
End Function

'=================================================================================================================
Function ResetRelGrid()
	Dim lngRow

	Call SetSpreadColor (-1 , -1)
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False

	lngRow = 1

	For lngRow = 1 To frm1.vspdData.MaxRows

		frm1.vspdData.Row = lngRow
		frm1.vspdData.Col = S_ExportFlg		'수출여부 
		If frm1.vspdData.Text = "1" Then
			ggoSpread.SSSetProtected	S_RetItemFlg, lngRow  '반품 
		Else
			ggoSpread.Spreadunlock	S_RetItemFlg, lngRow  '반품 
		End If


		frm1.vspdData.Col = S_RelDnFlg		'출하여부 
		If frm1.vspdData.Text = "0" Then
			ggoSpread.SSSetProtected	S_AutoDnFlg, lngRow  '자동출하여부 
		End If


		frm1.vspdData.Col = S_RelDnFlg
		If frm1.vspdData.Text = "1" Then
			Call OnChgRelBillFlg("DnCheck",lngRow) '매출타입Required
		else
			Call OnChgRelBillFlg("DnUnCheck",lngRow) '매출타입프로텍트 
		End If


		frm1.vspdData.Col = S_RelBillFlg
		If frm1.vspdData.Text = "1" Then
			Call OnChgRelBillFlg("BillCheck",lngRow) '출하타입Required
		else
			Call OnChgRelBillFlg("BillUnCheck",lngRow) '출하타입프로텍트 
		End If

		ggoSpread.SSSetProtected	S_PostDNTypeNm  ,lngRow    '출하타입명 
		ggoSpread.SSSetProtected	S_PostBillTypeNm, lngRow  '매출형태명 



	Next
	frm1.vspdData.ReDraw = True


End Function


'=================================================================================================================
Sub Form_Load()

	Call InitVariables
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)  '⊙: Format Contents  Field
    Call ggoOper.LockField(Document, "N")


	Call InitSpreadSheet
	Call SetDefaultVal
    Call SetToolbar("1110110100101111")

End Sub
'=================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If lsQueryMode = False Then Exit Sub
	If lsFncCopyFlag = True Then Exit Sub
    ggoSpread.Source = frm1.vspdData

	lgBlnFlgChgValue = True

	IF Row = 0 Then Exit Sub

	With frm1.vspdData

		.ReDraw = False

		SELECT CASE COL

		CASE S_PostDNTypePopup
			.Col = Col - 1
		    .Row = Row
		    Call OpenMinorCd(.Text)
			Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")

		CASE S_PostBillTypePopup
			.Col = Col - 1
		    .Row = Row
		    Call OpenTypePopup(.Text)
			Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")

		CASE S_InterComFlg	'멀티컴퍼니거래여부 
			Call SetInterRelGrid(Col, Row)

		CASE S_StoFlg
			Call SetStoRelGrid(Col, Row)
		CASE S_ExportFlg
			Call SetExportRelGrid(Col, Row)

		CASE S_CiFlg
			Call SetCCRelGrid(Col, Row)

		CASE S_RelDnFlg
			.Col = S_RelDnFlg
			.Row = Row
			If .Text = "" Then .Text = "0"

			If .Text = "1" Then
				Call OnChgRelBillFlg("DnCheck",Row)
			Else
				Call OnChgRelBillFlg("DnUnCheck",Row)
			End If

			Call SetDnRelGrid(Col, Row)

		CASE S_RelBillFlg
			.Col = S_RelBillFlg
			.Row = Row
			If .Text = "" Then .Text = "0"

			If .Text = "1" Then
				Call OnChgRelBillFlg("BillCheck",Row)
			Else
				Call OnChgRelBillFlg("BillUnCheck",Row)
			End If

		END SELECT

		.ReDraw = True


	End With
End Sub


'=================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Call SetPopupMenuItemInf("1111111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub
	End If

    frm1.vspdData.Row = Row


End Sub

'=================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'=================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)

    If Row <= 0 Then
		Exit Sub
    End If

    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End   ) --------------------------------------------------------------

End Sub

'=================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'=================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'=================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub

'=================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If CheckRunningBizProcess = True Then Exit Sub
    	If lgPageNo <> "" Then
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If
End Sub

'=================================================================================================================
Sub rdoUsageFlgAll_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgAll.value
End Sub

Sub rdoUsageFlgYes_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgYes.value
End Sub

Sub rdoUsageFlgNo_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgNo.value
End Sub

'=================================================================================================================
Function FncQuery()
    Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False

    ggoSpread.source = frm1.vspdData
    If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables

     Call DbQuery

    If Err.number = 0 Then
       FncQuery = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncNew()
    Dim IntRetCD

	On Error Resume Next
	Err.Clear
    FncNew = False

    ggoSpread.source = frm1.vspdData
    If ggoSpread.SSCheckChange Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "A")

    Call ggoOper.LockField(Document, "N")
    Call InitVariables
	Call SetRadio
    Call SetToolbar("11101101001111")
    Call SetDefaultVal

    If Err.number = 0 Then
       FncNew = True
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================
Function FncDelete()

    Exit Function

	On Error Resume Next
    Err.Clear

    FncDelete = False

    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If

    If DbDelete = False Then
       Exit Function
    End If

    Call ggoOper.ClearField(Document, "A")
    If Err.number = 0 Then
       FncDelete = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncSave()
    Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False

	ggoSpread.Source = frm1.vspdData


    ggoSpread.source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

    If ggoSpread.SSDefaultCheck = False Then
       Exit Function
    End If

    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then
       FncSave = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncCancel()
	Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    FncCancel = False                                                             '☜: Processing is NG
    ggoSpread.Source = Frm1.vspdData
	frm1.vspdData.ReDraw = False
    ggoSpread.EditUndo
	frm1.vspdData.ReDraw = True

    If Err.number = 0 Then
       FncCancel = True
    End If

    Set gActiveElement = document.ActiveElement
End Function

'=================================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim imRow
    Dim lngRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If

	With frm1
		lsQueryMode = True
		.vspdData.ReDraw = False
		.vspdData.focus

		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow

		For lngRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1

			.vspdData.Row = lngRow

			.vspdData.Col = S_DlvyLt
			.vspdData.Text = "0"

			.vspdData.Col = S_SoMgmtFlg
			.vspdData.Text = "1"

			.vspdData.Col = S_CreditChkFlg
			.vspdData.Text = "1"

			.vspdData.Col = S_UsageFlg
			.vspdData.Text = "1"

		Next

		SetInsertSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
		lgBlnFlgChgValue = True
   End With

    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncDeleteRow()

    Dim lDelRows
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if

    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With
    lgBlnFlgChgValue = True
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncDeleteRow = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()

    If Err.number = 0 Then
       FncPrint = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncPrev()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncPrev = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncNext()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncNext = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncExcel()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then
       FncExcel = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncFind()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_SINGLEMULTI, False)

    If Err.number = 0 Then
       FncFind = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function FncCopy()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	lsFncCopyFlag = True
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False

    ggoSpread.CopyRow

	Call SetCopyRow()

	frm1.vspdData.Focus

    If Err.number = 0 Then
       FncCopy = True
    End If

    Set gActiveElement = document.ActiveElement
End Function


'=================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub

'=================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	Call ResetRelGrid()
End Sub

'=================================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG

    ggoSpread.source = frm1.vspdData
    If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If Err.number = 0 Then
       FncExit = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=================================================================================================================
Function DbDelete()
    On Error Resume Next
End Function

'=================================================================================================================
Function DbDeleteOk()
    On Error Resume Next
End Function

'=================================================================================================================
Function DbQuery()

    Err.Clear

	lsQueryMode = False

    DbQuery = False


	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    Dim strVal

    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtSOtype=" & FilterVar(Trim(frm1.txtHSOtype.value), " ", "SNM")
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtSOType=" & FilterVar(Trim(frm1.txtSOType.value), " ", "SNM")
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

	End Function

'=================================================================================================================
Function DbQueryOk()

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE
	lsQueryMode = True
    Dim LngRow

    ggoSpread.Source = frm1.vspdData

    Call SetToolbar("1110111100111111")
	Call SetDefaultVal()

End Function

'=================================================================================================================
Function DbSave()

    Err.Clear

    Dim lRow
    Dim lGrpCnt
	Dim strVal, strDel

    DbSave = False


	If   LayerShowHide(1) = False Then
	     Exit Function
	End If



	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = 0
				Select Case .vspdData.Text
			        Case ggoSpread.InsertFlag							'☜: 신규 
						strVal = strVal & "C" & Parent.gColSep	& lRow & Parent.gColSep'☜: C=Create

						.vspdData.Col = S_SOType
						strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

						.vspdData.Col = S_SOTypeNm
						strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

				        .vspdData.Col = S_ExportFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_RetItemFlg
				        if Trim(.vspdData.Text) = "1" then
						strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_CIFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_RelDNFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_AutoDNFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_RelBillFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_SpStkFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_PostDNType
				        strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

						.vspdData.Col = S_PostBillType
				        strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

				        .vspdData.Col = S_DlvyLt
				        strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

				        .vspdData.Col = S_SoMgmtFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				         .vspdData.Col = S_CreditChkFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				         .vspdData.Col = S_DepositChkFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_InterComFlg	'멀티컴퍼니거래여부 
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_StoFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if


				        .vspdData.Col = S_UsageFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

					'Extend 필드를 사용하기 위해 변경 ... 시작 ... 2006.07.20 ...
					strVal = strVal & "0" & Parent.gColSep		'Ext1_Qty
					strVal = strVal & "0" & Parent.gColSep		'Ext2_Qty
					strVal = strVal & "0" & Parent.gColSep		'Ext1_Amt
					strVal = strVal & "0" & Parent.gColSep		'Ext2_Amt
					strVal = strVal & "" & Parent.gColSep		'Ext1_Cd
					strVal = strVal & "" & Parent.gRowSep		'Ext2_Cd
					'Extend 필드를 사용하기 위해 변경 ... 끝  ... 2006.07.20 ...


				        lGrpCnt = lGrpCnt + 1


					Case ggoSpread.UpdateFlag							'☜: 수정 
						strVal = strVal & "U" & Parent.gColSep	& lRow & Parent.gColSep'☜: U=Update

						.vspdData.Col = S_SOType
						strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

						.vspdData.Col = S_SOTypeNm
						strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

				        .vspdData.Col = S_ExportFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_RetItemFlg
				        if Trim(.vspdData.Text) = "1" then
						strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_CIFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_RelDNFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_AutoDNFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_RelBillFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_SpStkFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

						.vspdData.Col = S_PostDNType
				        strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

						.vspdData.Col = S_PostBillType
				        strVal = strVal & FilterVar(Trim(.vspdData.Text), " ", "SNM") & Parent.gColSep

				        .vspdData.Col = S_DlvyLt
				        strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep


				        .vspdData.Col = S_SoMgmtFlg
				        if Trim(.vspdData.Text) = "1" then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        else		            
				        strVal = strVal & "N" & Parent.gColSep
				        end if
				        
				        .vspdData.Col = S_CreditChkFlg
				        if Trim(.vspdData.Text) = "1" then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        else		            
				        strVal = strVal & "N" & Parent.gColSep
				        end if
				        
				        .vspdData.Col = S_DepositChkFlg
				        if Trim(.vspdData.Text) = "1" then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        else		            
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_InterComFlg	'멀티컴퍼니거래여부 
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

				        .vspdData.Col = S_StoFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if


				        .vspdData.Col = S_UsageFlg
				        if Trim(.vspdData.Text) = "1" then
				        strVal = strVal & "Y" & Parent.gColSep
				        else
				        strVal = strVal & "N" & Parent.gColSep
				        end if

					'Extend 필드를 사용하기 위해 변경 ... 시작 ... 2006.07.20 ...
					strVal = strVal & "0" & Parent.gColSep		'Ext1_Qty
					strVal = strVal & "0" & Parent.gColSep		'Ext2_Qty
					strVal = strVal & "0" & Parent.gColSep		'Ext1_Amt
					strVal = strVal & "0" & Parent.gColSep		'Ext2_Amt
					strVal = strVal & "" & Parent.gColSep		'Ext1_Cd
					strVal = strVal & "" & Parent.gRowSep		'Ext2_Cd
					'Extend 필드를 사용하기 위해 변경 ... 끝  ... 2006.07.20 ...

				        lGrpCnt = lGrpCnt + 1

					Case ggoSpread.DeleteFlag							'☜: 삭제 
						strDel = strDel & "D" & Parent.gColSep	& lRow & Parent.gColSep'☜: D=Delete
						.vspdData.Col = S_SOType
						strDel = strDel & FilterVar(Trim(.vspdData.Text), " " , "SNM") & Parent.gRowSep

						lGrpCnt = lGrpCnt + 1
				End Select


		Next

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 

	End With

    DbSave = True

End Function

'=================================================================================================================
Function DbSaveOk()

	Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables

    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주형태</font></td>
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
									<TD CLASS=TD5 NOWRAP>수주형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoType" ALT="수주형태" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="15XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSOType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCondtionPopup">&nbsp;<INPUT NAME="txtSOTypeNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>사용여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgAll" value=" " tag = "11XXX" checked>
											<label for="rdoUsageFlgAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgYes" value="Y" tag = "11XXX">
											<label for="rdoUsageFlgYes">사용</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoUsageFlg" id="rdoUsageFlgNo" value="N" tag = "11XXX">
											<label for="rdoUsageFlgNo">미사용</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="21" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSOType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
