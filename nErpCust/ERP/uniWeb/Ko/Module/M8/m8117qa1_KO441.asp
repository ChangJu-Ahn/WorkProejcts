<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M8117QA1
'*  4. Program Name         : 매입가계정잔액현황 조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/18
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Jin Ha
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
Const BIZ_PGM_ID = "M8117QB1_KO441.asp"												'☆: 비지니스 로직 ASP명 

Dim lgIsOpenPop

Dim C_MvmtRcptNo 		
Dim C_MvmtDt		
Dim C_MvmtType
Dim C_MvmtTypeNm
Dim C_MvmtAmt		
Dim C_IvLocAmt		
Dim C_IvTmpAcctBal
Dim C_BizArea
Dim C_BizAreaNm
Dim C_SpplCd
Dim C_SpplNm
Dim C_PoNo
Dim C_PoSeqNo
Dim C_GrNo
Dim C_GRSeqNo

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
	
	frm1.txtMvmtFrDt.text = StartDate
    frm1.txtMvmtToDt.text = EndDate 
    
    frm1.txtMvmtFrDt.focus 
	Set gActiveElement = document.activeElement
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "*", "NOCOOKIE", "QA") %>
End Sub
'========================================================================================================
Sub InitSpreadPosVariables()
	C_MvmtRcptNo 		= 1
	C_MvmtDt			= 2
	C_MvmtType 			= 3
	C_MvmtTypeNm		= 4
	C_MvmtAmt			= 5
	C_IvLocAmt			= 6
	C_IvTmpAcctBal		= 7
	C_BizArea			= 8
	C_BizAreaNm			= 9
	C_SpplCd			= 10
	C_SpplNm			= 11
	C_PoNo				= 12
	C_PoSeqNo			= 13
	C_GrNo				= 14
	C_GRSeqNo			= 15
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20030626",, parent.gAllowDragDropSpread
       .ReDraw = false
	
		.MaxCols = C_GRSeqNo + 1					
		.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_MvmtRcptNo,	"입고번호", 15
		ggoSpread.SSSetDate 	C_MvmtDt,		"입고일",	10,2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_MvmtType,		"입고형태", 10
		ggoSpread.SSSetEdit 	C_MvmtTypeNm,	"입고형태명", 15
		ggoSpread.SSSetFloat 	C_MvmtAmt,		"입고금액(GR)", 15,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 	C_IvLocAmt,		"매입금액(IR)", 15,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat 	C_IvTmpAcctBal,	"매입가계정잔액", 15,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit 	C_BizArea,		"사업장",	10
		ggoSpread.SSSetEdit 	C_BizAreaNm,	"사업장명",	15
		ggoSpread.SSSetEdit 	C_SpplCd,		"공급처",	10
		ggoSpread.SSSetEdit 	C_SpplNm,		"공급처명",	15
		ggoSpread.SSSetEdit 	C_PoNo,			"발주번호", 15
		ggoSpread.SSSetFloat 	C_PoSeqNo,		"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 	C_GrNo,			"재고처리번호", 15
		ggoSpread.SSSetFloat 	C_GRSeqNo,		"재고처리순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		
		'Call ggoSpread.SSSetColHidden(C_PoSeqNo,C_PoSeqNo,	True)
		'Call ggoSpread.SSSetColHidden( C_GRSeqNo,	C_GRSeqNo,	True)
		Call ggoSpread.SSSetColHidden( C_PoNo,	C_PoNo,	True)
		Call ggoSpread.SSSetColHidden( C_PoSeqNo,	C_PoSeqNo,	True)
		Call ggoSpread.SSSetColHidden( C_GRSeqNo,	C_GRSeqNo,	True)
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
            
	C_MvmtRcptNo 		= iCurColumnPos(1)
	C_MvmtDt			= iCurColumnPos(2)
	C_MvmtType 			= iCurColumnPos(3)
	C_MvmtTypeNm		= iCurColumnPos(4)
	C_MvmtAmt			= iCurColumnPos(5)
	C_IvLocAmt			= iCurColumnPos(6)
	C_IvTmpAcctBal		= iCurColumnPos(7)
	C_BizArea			= iCurColumnPos(8)
	C_BizAreaNm			= iCurColumnPos(9)
	C_SpplCd			= iCurColumnPos(10)
	C_SpplNm			= iCurColumnPos(11)
	C_PoNo				= iCurColumnPos(12)
	C_PoSeqNo			= iCurColumnPos(13)
	C_GrNo				= iCurColumnPos(14)
	C_GRSeqNo			= iCurColumnPos(15)
End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    frm1.vspdData.ReDraw = False
	
	frm1.vspdData.ReDraw = True
End Sub
'========================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "사업장"					
	arrParam(1) = "B_BIZ_AREA"					
	arrParam(2) = Trim(frm1.txtBizArea.Value)	
	arrParam(5) = "사업장"					

    arrField(0) = "BIZ_AREA_CD"					
    arrField(1) = "BIZ_AREA_NM"					
    
    
    arrHeader(0) = "사업장"					
    arrHeader(1) = "사업장명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea.focus
		Exit Function
	Else
		frm1.txtBizArea.Value	= arrRet(0)
		frm1.txtBizAreaNm.value = arrRet(1)
		frm1.txtBizArea.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'========================================================================================================
Function OpenMvmtType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtMvmtType.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True
	 
	arrParam(0) = "입출고형태" 
	arrParam(1) = "M_MVMT_TYPE"    
	arrParam(2) = UCase(Trim(frm1.txtMvmtType.Value))
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "입출고형태" 
	 
	arrField(0) = "io_type_cd" 
	arrField(1) = "io_type_Nm" 
	    
	arrHeader(0) = "입출고형태" 
	arrHeader(1) = "입출고형태명" 
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	lgIsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtMvmtType.Value = arrRet(0)
		frm1.txtMvmtTypeNm.Value = arrRet(1)
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
	End If 
 
End Function
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	ggoOper.FormatFieldByObjectOfCur frm1.txtTotBalanceAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	
End Sub
'========================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
    End Select
         
End Sub
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
'========================================================================================================
Function FindErrRow(iRow)
	
	frm1.vspdData.focus
	ggoSpread.Source = frm1.vspdData
		
	frm1.vspdData.Row = iRow
	frm1.vspdData.Col = 1
	frm1.vspdData.Action = 0    
End Function

'========================================================================================================
Sub txtMvmtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtMvmtFrDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtMvmtFrDt.focus
	End if
End Sub
'========================================================================================================
Sub txtMvmtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtMvmtToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtMvmtToDt.focus
	End if
End Sub
'========================================================================================================
Sub txtMvmtFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub
'========================================================================================================
Sub txtMvmtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub
'========================================================================================================
Sub txtIVFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIVFrDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtIVFrDt.focus
	End if
End Sub
'========================================================================================================
Sub txtIVToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIVToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtIVToDt.focus
	End if
End Sub
'========================================================================================================
Sub txtIVFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub
'========================================================================================================
Sub txtIVToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    On Error Resume Next
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    With frm1
	    if (UniConvDateToYYYYMMDD(.txtMvmtFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtMvmtToDt.text,Parent.gDateFormat,"")) OR Trim(.txtMvmtFrDt.text) = "" OR Trim(.txtMvmtToDt.text) = "" then	
			Call DisplayMsgBox("17a003", "X","입고일", "X")	
			Exit Function
		End if   
		
		if (UniConvDateToYYYYMMDD(.txtIvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtIvToDt.text,Parent.gDateFormat,"")) And Trim(.txtIvFrDt.text) <> "" And Trim(.txtIvToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","매입등록일", "X")	
			Exit Function
		End if   
		
	End with
	
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.activeElement
End Function
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
	Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData    
    
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")
    FncNew = True                                           
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement	
End Function
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement	
End Function
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)  
    Set gActiveElement = document.activeElement	                  
End Function
'========================================================================================
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
    
    Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData, -1, -1 ,parent.gCurrency		,C_MvmtAmt	,"A" ,"Q","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData, -1, -1 ,parent.gCurrency		,C_IvLocAmt	,"A" ,"Q","X","X")
    Call ReFormatSpreadCellByCellByCurrency2(frm1.vspdData, -1, -1 ,parent.gCurrency		,C_IvTmpAcctBal	,"A" ,"Q","X","X")
    
End Sub
'========================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim strVal
	
	DbQuery = False
    
    If LayerShowHide(1) = False then
       Exit Function 
    End if
    
    Err.Clear                                                           

	With frm1
    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&txtMvmtFrDt=" & .hdnMvmtFrDt.value
		    strVal = strVal & "&txtMvmtToDt=" & .hdnMvmtToDt.value
		    strVal = strVal & "&txtIvFrDt=" & .hdnIvFrDt.value
		    strVal = strVal & "&txtIvToDt=" & .hdnIvToDt.value
		    strVal = strVal & "&txtMvmtType=" & .hdnMvmtType.Value
		    strVal = strVal & "&rdoAppflg=" & .hdnrdoflg.value
		Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&txtMvmtFrDt=" & .txtMvmtFrDt.Text
		    strVal = strVal & "&txtMvmtToDt=" & .txtMvmtToDt.Text
		    strVal = strVal & "&txtIvFrDt=" & .txtIvFrDt.Text
		    strVal = strVal & "&txtIvToDt=" & .txtIvToDt.Text
		    strVal = strVal & "&txtMvmtType=" & .txtMvmtType.Value
		    
		    If .rdoAppflg(0).checked = True Then
			strVal = strVal & "&rdoAppflg=" & "B"	'차이분 
			Else
			strVal = strVal & "&rdoAppflg=" & "A"
			End if
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
		'테스트용 
		'dim i, a, b
		'a = 0
		'b = 0
		'for i = 1 to frm1.vspdData.MaxRows
		'	a = a + UNICDbl(GetSpreadText(frm1.vspdData,C_MvmtAmt,i,"X","X"))
		'	b = b + UNICDbl(GetSpreadText(frm1.vspdData,C_IvLocAmt,i,"X","X"))
		'Next
		'msgbox "입고금액합 ===>  " & a & "  :    " & "매입금액합 ===>  " & b
	Else
		Call SetFocusToDocument("M")	
		frm1.txtMvmtFrDt.focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입가계정잔액현황</font></td>
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
									<TD CLASS="TD5" NOWRAP>입고일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/m8117qa1_fpDateTime1_txtMvmtFrDt.js'></script>~&nbsp
										<script language =javascript src='./js/m8117qa1_fpDateTime1_txtMvmtToDt.js'></script>
									</TD>
									</TD>
									<TD CLASS="TD5" NOWRAP>매입등록일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/m8117qa1_fpDateTime2_txtIvFrDt.js'></script> ~&nbsp
										<script language =javascript src='./js/m8117qa1_fpDateTime2_txtIvToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>입출고형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMvmtType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT Alt="입출고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>조회구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=radio Class="Radio" ALT="조회구분" NAME="rdoAppflg" id = "rdoAppflg1" Value="B" checked tag="11"><label for="rdoAppflg1">&nbsp;차이분&nbsp;</label>
										<INPUT TYPE=radio Class="Radio" ALT="조회구분" NAME="rdoAppflg" id = "rdoAppflg2" Value="A" tag="11"><label for="rdoAppflg2">&nbsp;전체&nbsp;</label>
									</TD>			
									
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
								<TD CLASS=TD5 NOWRAP>입고금액합계</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m8117qa1_fpDoubleSingle1_txtTotMvmtAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>매입금액합계</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m8117qa1_fpDoubleSingle2_txtTotIvAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>매입가계정잔액합계</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m8117qa1_fpDoubleSingle3_txtTotBalanceAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m8117qa1_vaSpread1_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="hdnMvmtFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnrdoflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMaxRow" tag="24">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
