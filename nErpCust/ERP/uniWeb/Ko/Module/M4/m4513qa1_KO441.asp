<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M4513QA1
'*  4. Program Name         : 수입진행현황 조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Eun Hee
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit		

<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Const BIZ_PGM_ID = "M4513QB1_KO441.asp"												'☆: 비지니스 로직 ASP명 

Dim lgIsOpenPop

Dim C_PoNo
Dim C_PayMeth
Dim C_PoCur
Dim C_PoDt
Dim C_BpCd
Dim C_LcNo		
Dim C_LcNo_pop
Dim C_LcDocNo	
Dim C_AmendSeq	
Dim C_OpenDt
Dim C_LcType
Dim C_BlNo
Dim C_BlNo_pop
Dim C_BlDocNo
Dim C_LoadingDt
Dim C_BlIssueDt
Dim C_Setlmnt
Dim C_DischgeDt
Dim C_CcNo
Dim C_CcNo_pop
Dim C_IDNo
Dim C_IDDt
Dim C_RcptNo
Dim C_RcptNo_pop
Dim C_MvmtDt

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("M", -1, EndDate, parent.gDateFormat)

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE   
    lgBlnFlgChgValue = False    
    lgPageNo = 0
    frm1.vspdData.MaxRows = 0
End Sub
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtPoFrDt.Text	= StartDate 
	frm1.txtPoToDt.Text	= EndDate 
	frm1.txtPurGrpCd.focus
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If
	Set gActiveElement = document.activeElement
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'========================================================================================================
Sub InitSpreadPosVariables()
	C_PoNo		= 1
	C_PayMeth	= 2
	C_PoCur		= 3
	C_PoDt		= 4
	C_BpCd		= 5
	C_LcNo		= 6
	C_LcNo_pop	= 7
	C_LcDocNo	= 8
	C_AmendSeq	= 9
	C_OpenDt	= 10
	C_LcType	= 11
	C_BlNo		= 12
	C_BlNo_pop	= 13
	C_BlDocNo	= 14
	C_LoadingDt	= 15
	C_BlIssueDt	= 16
	C_Setlmnt	= 17
	C_DischgeDt	= 18
	C_CcNo		= 19
	C_CcNo_pop	= 20
	C_IDNo		= 21
	C_IDDt		= 22
	C_RcptNo	= 23
	C_RcptNo_pop= 24
	C_MvmtDt	= 25
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20030626",, parent.gAllowDragDropSpread
       .ReDraw = false
	
		.MaxCols = C_MvmtDt+1					
		.MaxRows = 0

		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_PoNo,			"Offer번호"	, 18
		ggoSpread.SSSetEdit 	C_PayMeth,		"결제방법"	, 10
		ggoSpread.SSSetEdit 	C_PoCur,		"화폐"	, 5
		ggoSpread.SSSetDate 	C_PoDt,			"발주일"	, 10 ,2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_BpCd,			"공급처"	, 10
		ggoSpread.SSSetEdit 	C_LcNo,			"LC관리번호", 18
		ggoSpread.SSSetButton 	C_LcNo_pop
		ggoSpread.SSSetEdit 	C_LcDocNo,		"LC번호"	, 20
		ggoSpread.SSSetEdit 	C_AmendSeq,		"AMEND차수"	, 10
		ggoSpread.SSSetDate 	C_OpenDt,		"LC개설일"	, 10,2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_LcType,		"L/C유형"	, 10
		ggoSpread.SSSetEdit 	C_BlNo,			"B/L관리번호",18
		ggoSpread.SSSetButton 	C_BlNo_pop
		ggoSpread.SSSetEdit 	C_BlDocNo,		"B/L번호"	, 20
		ggoSpread.SSSetDate 	C_LoadingDt,	"선적일"	, 10,2, Parent.gDateFormat
		ggoSpread.SSSetDate 	C_BlIssueDt,	"B/L접수일"	, 10,2, Parent.gDateFormat
		ggoSpread.SSSetDate 	C_Setlmnt,		"지불예정일", 10,2, Parent.gDateFormat
		ggoSpread.SSSetDate 	C_DischgeDt,	"도착일"	, 10,2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_CcNo,			"통관관리번호", 18
		ggoSpread.SSSetButton 	C_CcNo_pop
		ggoSpread.SSSetEdit 	C_IDNo,			"신고번호", 20
		ggoSpread.SSSetDate 	C_IDDt,			"신고일",	10,2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_RcptNo,		"입고번호", 18
		ggoSpread.SSSetButton 	C_RcptNo_pop
		ggoSpread.SSSetDate 	C_MvmtDt,		"입고일",	10,2, Parent.gDateFormat
		
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = true
		Call SetSpreadLock 
	End With
End Sub
'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
	
	C_PoNo		= iCurColumnPos(1)
	C_PayMeth	= iCurColumnPos(2)
	C_PoCur		= iCurColumnPos(3)
	C_PoDt		= iCurColumnPos(4)
	C_BpCd		= iCurColumnPos(5)
	C_LcNo		= iCurColumnPos(6)
	C_LcNo_pop	= iCurColumnPos(7)
	C_LcDocNo	= iCurColumnPos(8)
	C_AmendSeq	= iCurColumnPos(9)
	C_OpenDt	= iCurColumnPos(10)
	C_LcType	= iCurColumnPos(11)
	C_BlNo		= iCurColumnPos(12)
	C_BlNo_pop	= iCurColumnPos(13)
	C_BlDocNo	= iCurColumnPos(14)
	C_LoadingDt	= iCurColumnPos(15)
	C_BlIssueDt	= iCurColumnPos(16)
	C_Setlmnt	= iCurColumnPos(17)
	C_DischgeDt	= iCurColumnPos(18)
	C_CcNo		= iCurColumnPos(19)
	C_CcNo_pop	= iCurColumnPos(20)
	C_IDNo		= iCurColumnPos(21)
	C_IDDt		= iCurColumnPos(22)
	C_RcptNo	= iCurColumnPos(23)
	C_RcptNo_pop= iCurColumnPos(24)
	C_MvmtDt	= iCurColumnPos(25)       

End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SpreadUnlock C_LcNo_pop,	-1,	C_LcNo_pop,-1
	ggoSpread.SpreadUnlock C_BlNo_pop,	-1,	C_BlNo_pop,-1
	ggoSpread.SpreadUnlock C_CcNo_pop,	-1,	C_CcNo_pop,-1
	ggoSpread.SpreadUnlock C_RcptNo_pop,-1,	C_RcptNo_pop,-1
End Sub
'========================================================================================================
'------------------------------------------  OpenPurGrpCd()  -------------------------------------------------
Function OpenPurGrpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	If frm1.txtPurGrpCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
	End If	

End Function 
'------------------------------------------  OpenSppl()  -------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "공급처"					
    arrHeader(1) = "공급처명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function
'========================================================================================================
'------------------------------------------  OpenPoNo()  -------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(1)
	
	
			
	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lgIsOpenPop = True
	
	arrParam(0) = ""
	arrParam(1) = ""
		
	iCalledAspName = AskPRAspName("M3111PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA2", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCDtl()  +++++++++++++++++++++++++++++++++++++++
Function OpenLCDtl()
 	Dim arrRet
 	Dim arrParam(5)
 	Dim iCalledAspName
 	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
		
 	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"X","X"))
 	'arrParam(1) = Trim(frm1.txtBpNm.Value)
 	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_LcDocNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(3) = Trim(GetSpreadText(frm1.vspdData,C_AmendSeq,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(4) = Trim(GetSpreadText(frm1.vspdData,C_LcNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(5) = Trim(GetSpreadText(frm1.vspdData,C_PoCur,frm1.vspdData.ActiveRow,"X","X"))
	
	If arrParam(4) <> "" Then
 		iCalledAspName = AskPRAspName("M3212PA1")

 		If Trim(iCalledAspName) = "" Then
 			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212PA1", "X")
 			lgIsOpenPop = False
 			Exit Function
 		End If
			
 		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	lgIsOpenPop = False
	
 End Function
 '+++++++++++++++++++++++++++++++++++++++++++++++  OpenBLDtl()  +++++++++++++++++++++++++++++++++++++++
Function OpenBLDtl()
 	Dim arrRet
 	Dim arrParam(4)
 	Dim iCalledAspName
 	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
		
 	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"X","X"))
 	'arrParam(1) = Trim(frm1.txtBpNm.Value)
 	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_BlDocNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(3) = Trim(GetSpreadText(frm1.vspdData,C_BlNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(4) = Trim(GetSpreadText(frm1.vspdData,C_PoCur,frm1.vspdData.ActiveRow,"X","X"))
	
	If arrParam(3) <> "" Then
 		iCalledAspName = AskPRAspName("M5212PA1")

 		If Trim(iCalledAspName) = "" Then
 			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5212PA1", "X")
 			lgIsOpenPop = False
 			Exit Function
 		End If
			
 		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	lgIsOpenPop = False
	
 End Function
  '+++++++++++++++++++++++++++++++++++++++++++++++  OpenCCDtl()  +++++++++++++++++++++++++++++++++++++++
Function OpenCCDtl()
 	Dim arrRet
 	Dim arrParam(3)
 	Dim iCalledAspName
 	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
		
 	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(1) = Trim(GetSpreadText(frm1.vspdData,C_IDNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_CcNo,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(3) = Trim(GetSpreadText(frm1.vspdData,C_PoCur,frm1.vspdData.ActiveRow,"X","X"))
	
	If arrParam(2) <> "" Then
 		iCalledAspName = AskPRAspName("M4212PA1")

 		If Trim(iCalledAspName) = "" Then
 			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212PA1", "X")
 			lgIsOpenPop = False
 			Exit Function
 		End If
			
 		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	lgIsOpenPop = False
	
 End Function
  '+++++++++++++++++++++++++++++++++++++++++++++++  OpenRcptDtl()  +++++++++++++++++++++++++++++++++++++++
Function OpenRcptDtl()
 	Dim arrRet
 	Dim arrParam(1)
 	Dim iCalledAspName
 	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
		
 	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"X","X"))
 	arrParam(1) = Trim(GetSpreadText(frm1.vspdData,C_RcptNo,frm1.vspdData.ActiveRow,"X","X"))
	
	If arrParam(1) <> "" Then
 		iCalledAspName = AskPRAspName("M4112PA1")

 		If Trim(iCalledAspName) = "" Then
 			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4112PA1", "X")
 			lgIsOpenPop = False
 			Exit Function
 		End If
			
 		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 				"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	lgIsOpenPop = False
	
 End Function
'========================================================================================================
Sub Form_Load()
        
    Call LoadInfTB19029   
    Call ggoOper.LockField(Document, "N")  
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                                                    
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables
    Call SetToolbar("1100000000001111")
End Sub

'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoFrDt.Focus
	End If
End Sub
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoToDt.Focus
	End If
End Sub
'========================================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	 gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101111111")

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
    
	frm1.vspdData.Row = Row
End Sub
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '☜: 재쿼리 체크 
		
		if Trim(lgPageNo) = "" then exit sub
		If lgPageNo > 0   Then            '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음						
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
'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    Dim sPayType
	Dim sCurCurrency
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		.Row = Row

		If Row > 0 then

			if  Col = C_LcNo_pop Then
				Call OpenLCDtl()
			elseIf Col = C_BlNo_pop Then
			    Call OpenBLDtl()
			elseIf Col = C_CcNo_pop Then
				Call OpenCCDtl()
			elseIf  Col = C_RcptNo_pop Then
				Call OpenRcptDtl()
			end if

		End If
    
    End With
End Sub
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    On Error Resume Next
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    With frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			Exit Function
		End if   
	End with
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData			
    Call InitVariables
    														
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
'==============================  FncSave()  ================================================
Function FncSave()     
End Function
'==============================  FncPrint()  ================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'==============================  FncExcel()  ================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'==============================  FncFind()  ================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)   
    Set gActiveElement = document.activeElement                 
End Function
'==============================  FncExit()  ================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet() 
    Call ggoSpread.ReOrderingSpreadData()
        
End Sub
'========================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim strVal
	
	DbQuery = False
    
    If LayerShowHide(1) = False then
       Exit Function 
    End if
    
    Err.Clear                                                           

	With frm1
    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPoFrDt=" & Trim(.hdnPoFrDt.value)
		    strVal = strVal & "&txtPoToDt=" & Trim(.hdnPoToDt.value)
		    strVal = strVal & "&txtPoNo=" & Trim(.hdnPoNo.Value)
		    strVal = strVal & "&txtPurGrpCd=" & Trim(.hdnPurGrpCd.Value)
		    strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.Value)
		Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtPoFrDt=" & Trim(.txtPoFrDt.text)
		    strVal = strVal & "&txtPoToDt=" & Trim(.txtPoToDt.text)
		    strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.Value)
		    strVal = strVal & "&txtPurGrpCd=" & Trim(.txtPurGrpCd.Value)
		    strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.Value)
		End If 
		
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	   
		strVal = strVal & "&lgPageNo="		 & lgPageNo 

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  


		.hdnMaxRow.value = .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    Call SetToolbar("1100000000011111")	
    DbQuery = True
End Function
'========================================================================================================
Function DbQueryOk()													
	Dim iRow
	
    lgBlnFlgChgValue = False
    lgIntFlgMode = Parent.OPMD_UMODE

	Call ggoOper.LockField(Document, "Q")	

    If frm1.vspdData.MaxRows > 0 then
		frm1.vspdData.focus	
	Else
		frm1.txtPlantCd.focus
	End if
	
End Function
'========================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입진행현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11xXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>	
									
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4513qa1_fpDateTime2_txtPoFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4513qa1_fpDateTime2_txtPoToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>
							         <TD CLASS="TD5" NOWRAP>Offer번호</TD>
									 <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=29 MAXLENGTH=18 ALT="Offer번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>		
									
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m4513qa1_A_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="hdnPoFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMaxRow" tag="24">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
