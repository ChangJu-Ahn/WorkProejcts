<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : u2123ra1.asp																			*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2004/08/11																*
'*  8. Modified date(Last)  : 2004/08/11																*
'*  9. Modifier (First)     : Park, BumSoo																*
'* 10. Modifier (Last)      : Park, BumSoo																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

'================================================================================================================================
Const BIZ_PGM_ID 		= "u2123rb1.asp"

'================================================================================================================================

Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_PlantCd
Dim C_PlantNm
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_OrderQty
Dim C_DvryDt
Dim	C_RcptQty
Dim	C_UnRcptQty
Dim	C_FirmDvryQty
Dim C_RemainQty
Dim C_DvryPlanDt
Dim C_DvryQty
Dim C_SLCD
Dim C_SLPOP
Dim C_SLNM
Dim C_SPLIT_SEQ_NO

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'================================================================================================================================

'================================================================================================================================
Const C_MaxKey          = 28                                           '☆: key count of SpreadSheet
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate
'================================================================================================================================    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("d", -7, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
	
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	
	Dim iCodeArr
		
	Err.Clear
	
	With frm1
		.txtFrPoDt.text = StartDate
		.txtToPoDt.text = EndDate
		.txtBpCd.value = arrParam(0)
		Call CommonQueryRs(" BP_NM", " B_BIZ_PARTNER", " BP_CD = '" & FilterVar(Trim(.txtBpCd.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		.txtBpNm.value = Replace(lgF0, Chr(11),"")
		.hdnSupplierCd.value 	= arrParam(0)
		.hdnGroupCd.value 		= arrParam(2)
		.hdnGroupNm.value 		= arrParam(3)
		.hdnRefType.value 		= arrParam(8)
		.hdnRcptType.value 		= arrParam(9)
		
		.txtFrPoDt.Text			= arrParam(10)
		.txtToPoDt.Text			= arrParam(10)
		.HDNPlantCd.value		= PopupParent.gPlant
		'.HDNPlantNm.value		= PopupParent.gPlantNm
	End With
	
	Call CommonQueryRs(" RCPT_FLG", " M_MVMT_TYPE", " IO_TYPE_CD = '" & FilterVar(Trim(frm1.hdnRcptType.value),"","SNM") & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    IF Len(lgF0) Then
		iCodeArr = Split(lgF0, Chr(11))
		    
		If Err.number <> 0 Then
			MsgBox Err.description,vbInformation,PopupParent.gLogoName 
			Err.Clear 
			Exit Sub
		End If
		frm1.hdnRcptFlg.value 	= iCodeArr(0)
	End if	
	
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050420", ,PopupParent.gAllowDragDropSpread
		frm1.vspdData.OperationMode = 5
		
		.ReDraw = false
					
		.MaxCols = C_SPLIT_SEQ_NO + 1    
		.MaxRows = 0    
			
		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit		C_OrderNo,		"수주번호", 15
		ggoSpread.SSSetEdit		C_OrderSeq,		"행번"    ,  7
		ggoSpread.SSSetEdit		C_ItemCode,		"품목"    , 10,,,18,2
		ggoSpread.SSSetEdit		C_ItemName,		"품목명"  , 18
		ggoSpread.SSSetEdit		C_Spec,			"규격"    , 15
		ggoSpread.SSSetEdit		C_OrderUnit,	"단위"    ,  7,,,3,2
		ggoSpread.SSSetDate 	C_DvryPlanDt,	"납품예정일자",12, 2, Popupparent.gDateFormat
		ggoSpread.SSSetFloat	C_DvryQty,		"납품예정수량",12,Popupparent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,Popupparent.gComNum1000,Popupparent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit		C_PlantCd,		"납품처"  , 10
		ggoSpread.SSSetEdit		C_PlantNm,		"납품처명",	12
		ggoSpread.SSSetEdit		C_SPLIT_SEQ_NO,	"분할번호",	 7

			
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
		ggoSpread.SSSetSplit2(2)
						
		Call SetSpreadLock()
						
		.ReDraw = true    
    
	End With
	   
End Sub
'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()
		
		C_OrderNo		= 1
		C_OrderSeq		= 2
		C_ItemCode		= 3
		C_ItemName		= 4
		C_Spec			= 5
		C_OrderUnit		= 6
		C_DvryPlanDt	= 7
		C_DvryQty		= 8
		C_PlantCd		= 9
		C_PlantNm		= 10
		C_SPLIT_SEQ_NO	= 11
		
End Sub
'================================================================================================================================
Sub GetSpreadColumnPos()
      
    Dim iCurColumnPos
    
 	ggoSpread.Source = frm1.vspdData
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_OrderNo		= iCurColumnPos( 1)
			C_OrderSeq		= iCurColumnPos(2)
			C_ItemCode		= iCurColumnPos(3)
			C_ItemName		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_OrderUnit		= iCurColumnPos(6)
			C_DvryPlanDt	= iCurColumnPos(7)
			C_DvryQty		= iCurColumnPos(8)
			C_PlantCd		= iCurColumnPos(9)
			C_PlantNm		= iCurColumnPos(10)
			C_SPLIT_SEQ_NO	= iCurColumnPos(11)

End Sub    
'================================================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols - 2)

			For intRowCnt = 1 To frm1.vspdData.MaxRows

				frm1.vspdData.Row = intRowCnt

				If frm1.vspdData.SelModeSelected Then
								
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
						frm1.vspdData.Col = intColCnt+1 ' GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next
					intInsRow = intInsRow + 1
				End IF								
			Next
			
		End if			
		Self.Returnvalue = arrReturn
		Self.Close()
End Function	
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtPoNo.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
	
	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M3111PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If

End Function
'================================================================================================================================
Function OpenSlcd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtslCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "납품창고"	
	arrParam(1) = "(SELECT DISTINCT B.D_BP_CD,C.SL_NM          "
	arrParam(1) = arrParam(1) & " FROM M_PUR_ORD_HDR 	    A, "
	arrParam(1) = arrParam(1) & "       m_scm_firm_pur_rcpt B, "
	arrParam(1) = arrParam(1) & "       b_storage_location  C  "
	arrParam(1) = arrParam(1) & " WHERE A.PO_NO = B.PO_NO      " 
	arrParam(1) = arrParam(1) & "   AND B.D_BP_CD = C.SL_CD    "		
	arrParam(1) = arrParam(1) & "   AND A.BP_CD = " & FilterVar(Trim(frm1.txtBpCd.value),"''","S") & ") A "			
	
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = Trim(frm1.txtSLNM.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "납품창고"			
	
    arrField(0) = "D_BP_CD"	
    arrField(1) = "SL_NM"
    
    arrHeader(0) = "납품창고"		
    arrHeader(1) = "납품창고명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtslCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtslCd.Value= arrRet(0)		
		frm1.txtslNm.Value= arrRet(1)	
		frm1.txtslCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 


'===============================================
Function OpenItemcd()
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

If gblnWinEvent = True Or UCase(frm1.txtItemCd.className) = Ucase(PopupParent.UCN_PROTECTED) Then Exit Function

gblnWinEvent = True

arrParam(0) = "품목" 
arrParam(1) = "(SELECT DISTINCT D.ITEM_CD,C.ITEM_NM          "
arrParam(1) = arrParam(1) & " FROM M_PUR_ORD_HDR     A, "
arrParam(1) = arrParam(1) & "       m_scm_firm_pur_rcpt B, "
arrParam(1) = arrParam(1) & "       b_item  C , "
arrParam(1) = arrParam(1) & "       M_PUR_ORD_DTL     D "
arrParam(1) = arrParam(1) & " WHERE A.PO_NO = B.PO_NO      " 
arrParam(1) = arrParam(1) & "   AND A.PO_NO = D.PO_NO      " 
arrParam(1) = arrParam(1) & "   AND D.ITEM_CD = C.ITEM_CD    " 
arrParam(1) = arrParam(1) & "   AND A.BP_CD = " & FilterVar(Trim(frm1.txtBpCd.value),"''","S") & ") A " 

arrParam(2) = Trim(frm1.txtSLCd.Value)
arrParam(3) = Trim(frm1.txtSLNM.Value) 

arrParam(4) = "" 
arrParam(5) = "품목" 

    arrField(0) = "ITEM_CD" 
    arrField(1) = "ITEM_NM"
    
    arrHeader(0) = "품목코드" 
    arrHeader(1) = "품목명" 
    
arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

gblnWinEvent = False

If arrRet(0) = "" Then
frm1.txtItemCd.focus 
Set gActiveElement = document.activeElement
Exit Function
Else
frm1.txtItemCd.Value= arrRet(0) 
frm1.txtItemNm.Value= arrRet(1) 
frm1.txtItemCd.focus 
Set gActiveElement = document.activeElement
End If 

End Function
'===============================================



'================================================================================================================================
Function OpenSortPopup()

	
	On Error Resume Next
	
End Function
'================================================================================================================================
Function OpentxtDlvyNo()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발행번호"	
	arrParam(1) = "M_SCM_DLVY_PUR_RCPT A ,B_BIZ_PARTNER B"
	arrParam(2) = Trim(frm1.txtDlvyNo.value)
	arrParam(3) = ""
	arrParam(4) = "A.BP_CD = B.BP_CD"			
	arrParam(5) = "발행번호"			
	
    arrField(0) = "ED15" & PopupParent.gColSep & "Dlvy_No"	
    arrField(1) = "ED06" & PopupParent.gColSep & "A.BP_CD"
    arrField(2) = "ED20" & PopupParent.gColSep & "B.BP_NM"	
    
    arrHeader(0) = "발행번호"		
    arrHeader(1) = "공급처"
    arrHeader(2) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtDlvyNo.value = arrRet(2)
		frm1.txtDlvyNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	
'================================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	                                           
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field 
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
		
	Call InitSpreadSheet()
		
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,PopupParent.gDateFormat,"")) And trim(.txtFrPoDt.text) <> "" And trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			.txtToPoDt.Focus()
			Exit Function
		End if   
	End with
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
        
	Call InitVariables												
	
	If CheckRunningBizProcess = True Then Exit Function
    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function
'================================================================================================================================
Function DbQuery()
	
	Dim strVal
	
	Err.Clear															'☜: Protect system from crashing

	DbQuery = False														'⊙: Processing is NG

    If LayerShowHide(1) = False Then Exit Function
    
    Call MakeKeyStream()
    
	strVal = BIZ_PGM_ID & "?txtMode="	& PopupParent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgPageNo

	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

	DbQuery = True														'⊙: Processing is NG
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPoNo.focus
	End If

End Function
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()

	With frm1
		
		lgKeyStream =               UCase(Trim(.txtBpCd.value))   & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtfrpodt.text))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txttopodt.text))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtSlCd.value))   & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtITEMCd.value)) & PopupParent.gColSep
		
		If .rdoAppflg1.checked = True Then
			lgKeyStream = lgKeyStream & "" & PopupParent.gColSep
		ElseIf .rdoAppflg2.checked = True Then
			lgKeyStream = lgKeyStream & "Y" & PopupParent.gColSep
		ElseIf .rdoAppflg3.checked = True Then
			lgKeyStream = lgKeyStream & "N" & PopupParent.gColSep
		End If
		 
			
		
	End With
			 
End Sub    

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>업체</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="업체" tag="14NXXU">
											   <INPUT TYPE=TEXT ALT="업체명" ID="txtBpNm" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>납품예정일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="납품예정일" NAME="txtFrPoDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="납품예정일" NAME="txtToPoDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>납품창고</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=10 MAXLENGTH=7 ALT="납품창고" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSlcd()">
											   <INPUT TYPE=TEXT AlT="납품창고명" ID="txtSlNm" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 ALT="품목" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemcd()">
											   <INPUT TYPE=TEXT AlT="품목명" ID="txtItemNm" tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>구분</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="구분" NAME="rdoAppflg" id = "rdoAppflg1" Value="A" tag="11"><label for="rdoAppflg1">&nbsp;전체&nbsp;</label>
											   <INPUT TYPE=radio Class="Radio" ALT="구분" NAME="rdoAppflg" id = "rdoAppflg2" Value="N" checked tag="11"><label for="rdoAppflg2">&nbsp;정상&nbsp;</label>
											   <INPUT TYPE=radio Class="Radio" ALT="구분" NAME="rdoAppflg" id = "rdoAppflg3" Value="Y" tag="11"><label for="rdoAppflg3">&nbsp;반품&nbsp;</label>
						</TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% ID=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>



<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>