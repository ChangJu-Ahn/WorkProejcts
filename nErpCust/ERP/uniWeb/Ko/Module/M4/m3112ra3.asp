<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : m3112ra3																	*
'*  4. Program Name         : 발주내역참조																*
'*  5. Program Desc         : Local L/C 내역등록을 위한 발주내역참조 (Business Logic Asp)					*
'*  6. Comproxy List        : M31128ListPoDtlForLcSvr													*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/07/11																*
'*  9. Modifier (First)     : 																			*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"										*
'*                            this mark(☆) Means that "must change"										*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>발주내역참조</TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
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

Const BIZ_PGM_ID 		= "m3112rb3.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 14                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgSelectList                   
Dim lgSelectListDT                 

Dim lgSortFieldNm                  
Dim lgSortFieldCD                  

Dim lgPopUpR                       

Dim lgKeyPos                       
Dim lgKeyPosVal                    
Dim lgCookValue 

Dim IsOpenPop  
Dim gblnWinEvent

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
 lgStrPrevKey     = ""
 lgPageNo         = ""
 lgBlnFlgChgValue = False
 lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
 lgSortKey        = 1
			
 gblnWinEvent = False
 ReDim arrReturn(0,0)
 Self.Returnvalue = arrReturn
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
 Sub SetDefaultVal()
 	Dim arrParam
		
 	arrParam = arrParent(1)
		
 	With frm1
 		.txtPurGrp.Value 		= arrParam(0)		<%'구매그룹 %>    
 		.txtPurGrpNm.value 		= arrParam(1)       <%'구매그룹명 %>  
 		.txtBeneficiary.value	= arrParam(2)       <%'수혜자 %>      
 		.txtBeneficiaryNm.value = arrParam(3)       <%'수혜자명 %>    
 		.txtCurrency.value 		= arrParam(4)       <%'화폐 %>        
 		'.txtPayTerms.value 		= arrParam(5)       <%'결제방법 %>    
 		'.txtPayTermsNm.value 	= arrParam(6)       <%'결제방법명 %>  
 		.txtPONo.value 			= arrParam(7)       <%'발주번호 %>    
 		.vspdData.OperationMode = 5
 	End With

 End Sub

'=========================================  LoadInfTB19029()  ============================================
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

 	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
 	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'=========================================  InitSpreadSheet()  ============================================
 Sub InitSpreadSheet()
 	Call SetZAdoSpreadSheet("m3112ra3","S","A","V20030711",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
 									C_MaxKey, "X","X")
 	Call SetSpreadLock 	    
 End Sub
'=========================================  SetSpreadLock()  ============================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================  2.3.1 OkClick()  ==========================================
 Function OKClick()
	
 	Dim intColCnt, intRowCnt, intInsRow

 	If frm1.vspdData.SelModeSelCount > 0 Then 

 		intInsRow = 0

 		Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-1)

 		For intRowCnt = 1 To frm1.vspdData.MaxRows

 			frm1.vspdData.Row = intRowCnt

 			If frm1.vspdData.SelModeSelected Then
 				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
 					'frm1.vspdData.Col = intColCnt + 1
 					'arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
 					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
 					arrReturn(intInsRow, intColCnt) = Trim(frm1.vspdData.Text)
 				Next

 				intInsRow = intInsRow + 1

 			End IF
 		Next
 	End if			
		
 	Self.Returnvalue = arrReturn
 	Self.Close()
 End Function	

'=========================================  2.3.2 CancelClick()  ========================================
 Function CancelClick()
 	Redim arrReturn(1,1)
 	arrReturn(0,0) = ""
 	Self.Returnvalue = arrReturn
 	Self.Close()
 End Function
'++++++++++++++++++++++++++++++++++++++++++++  OpenItem()  ++++++++++++++++++++++++++++++++++++++++++++++
 Function OpenItem()
 	Dim arrRet
 	Dim arrParam(5), arrField(1)
 	Dim iCalledAspName
		
 	If gblnWinEvent = True Then Exit Function
	
 	gblnWinEvent = True
	
 	arrParam(0) = Trim(frm1.txtItem.value)		' Item Code
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "30"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명						
	
 	iCalledAspName = AskPRAspName("B1B01PA2")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "B1B01PA2", "X")
 		IsOpenPop = False
 		Exit Function
 	End If
			
 	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrparam,arrField), _
 		"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")


 	gblnWinEvent = False
	
 	If arrRet(0) = "" Then
 		frm1.txtItem.focus
 		Exit Function
 	Else
 		frm1.txtItem.Value = arrRet(0)
 		frm1.txtItemNm.Value = arrRet(1)
 		frm1.txtItem.focus
 		Set gActiveElement = document.activeElement
 	End If
 End Function
'+++++++++++++++++++++++++++++++++++++  OpenTrackingNo()  ++++++++++++++++++++++++++++++++++++++++++++++ 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		frm1.txtTrackingNo.focus
		lgBlnFlgChgValue = True
	End If	

End Function
'=========================================  3.1.1 Form_Load()  ==========================================
 Sub Form_Load()
 	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

 	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>

 	Call InitVariables														    '⊙: Initializes local global variables
 	Call SetDefaultVal	
 	Call InitSpreadSheet()
 	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
 	Call FncQuery()
 End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub
'=========================================  3.3.1 vspdData_DblClick()  ==================================
 Function vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
       Exit Function
     End If

 	If frm1.vspdData.MaxRows > 0 Then
 		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
 			Call OKClick
 		End If
 	End If
 End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
 Function vspdData_KeyPress(KeyAscii)
      On Error Resume Next
      If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
         Call OKClick()
      ElseIf KeyAscii = 27 Then
         Call CancelClick()
      End If
 End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
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
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    
	FncQuery = False                                                 
	   
	Err.Clear                                                        

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	   							
	Call InitVariables												

	If DbQuery = False Then Exit Function							

	FncQuery = True		
	Set gActiveElement = document.activeElement							
        
End Function

'********************************************  5.1 DbQuery()  *******************************************
 Function DbQuery()
 	Err.Clear															<%'☜: Protect system from crashing%>

 	DbQuery = False														<%'⊙: Processing is NG%>

 	If LayerShowHide(1) = False Then
 		Exit Function
 	End If

 	Dim strVal
		
 	With frm1
		
 	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
 		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
 																				<%'☆: 조회 조건 데이타 %>
 		strVal = strVal & "&txtItem=" 			& Trim(.txtHItem.value)			<%'품목 %>
 		strVal = strVal & "&txtPurgrp=" 		& Trim(.txtHPurgrp.value)		<%'구매그룹 %>
 		strVal = strVal & "&txtBeneficiary=" 	& Trim(.txtHBeneficiary.value)	<%'수혜자 %>
 		strVal = strVal & "&txtCurrency=" 		& Trim(.txtHCurrency.value)		<%'화폐 %>
 		'strVal = strVal & "&txtPayTerms=" 		& Trim(.txtHPayTerms.value)		<%'결제방법 %>
 		strVal = strVal & "&txtPONo=" 			& Trim(.txtHPONo.value)			<%'발주번호 %>
 		strVal = strVal & "&lgStrPrevKey=" 		& lgStrPrevKey						
 	Else
 		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
 																				<%'☆: 조회 조건 데이타 %>
 		strVal = strVal & "&txtItem=" 			& Trim(.txtItem.value)			<%'품목 %>
 		strVal = strVal & "&txtPurgrp=" 		& Trim(.txtPurgrp.value)		<%'구매그룹 %>
 		strVal = strVal & "&txtBeneficiary=" 	& Trim(.txtBeneficiary.value)	<%'수혜자 %>
 		strVal = strVal & "&txtCurrency=" 		& Trim(.txtCurrency.value)		<%'화폐 %>
 		'strVal = strVal & "&txtPayTerms=" 		& Trim(.txtPayTerms.value)		<%'결제방법 %>
 		strVal = strVal & "&txtPONo=" 			& Trim(.txtPONo.value)			<%'발주번호 %>
 		strVal = strVal & "&lgStrPrevKey=" 		& lgStrPrevKey
 	End If
		strVal = strVal & "&txtTrackingNo="		& Trim(.txtTrackingNo.value)
 	End With
    strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
 	strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")	
         
    strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
 	strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

 	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

 	DbQuery = True														<%'⊙: Processing is NG%>
 End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

 lgIntFlgMode = PopupParent.OPMD_UMODE
	
 If frm1.vspdData.MaxRows > 0 Then
 	frm1.vspdData.Focus
 	frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
 Else
 	frm1.txtItem.focus
 	Set gActiveElement = document.activeElement
 End If

End Function


'========================================================================================================
' Function Name : OpenOrderBy
'========================================================================================================
Function OpenOrderByPopup()
 Dim arrRet
	
 On Error Resume Next
	
 'If lgIsOpenPop = True Then Exit Function
 lgIsOpenPop = True

 arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

 lgIsOpenPop = False
	
 If arrRet(0) = "X" Then
    Exit Function
 Else
    Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
    Call InitVariables
    Call InitSpreadSheet()       
End If
End Function

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
 					<TD CLASS=TD5 NOWRAP>품목</TD>
 					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SIZE=10  TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 TAG="14"></TD>
 					<TD CLASS=TD5 NOWRAP>구매그룹</TD>
 					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="14XXXU" ALT="구매그룹">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="14"></TD>
 				</TR>		
 				<TR>
 					<TD CLASS=TD5 NOWRAP>수혜자</TD>
 					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="14XXXU" ALT="수혜자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
 					<TD CLASS=TD5 NOWRAP>화폐</TD>
 					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="14XXXU" ALT="화폐"></TD>
 				</TR>	
 				<TR>
					<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
					<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
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
 				<TD HEIGHT="100%">
 					<script language =javascript src='./js/m3112ra3_vaSpread1_vspdData.js'></script>
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
 				<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
 									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>	
 				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
 						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
 				<TD WIDTH=10>&nbsp;</TD>
 			</TR>
 		</TABLE>
 	</TD>
 </TR>
 <TR>
 	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
 	</TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurgrp" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHBeneficiary" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPONo" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
