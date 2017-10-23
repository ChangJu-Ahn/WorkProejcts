<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4112ra4.asp																*
'*  4. Program Name         : Local 출하내역참조(Local L/C내역등록에서)									*
'*  5. Program Desc         : Local 출하내역참조(Local L/C내역등록에서)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2002/04/24																*
'*  9. Modifier (First)     : Hyungsuk Kim																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'*                            2. 2002.04/24 : Ado 변환 													*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>출하내역참조</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "s4112rb4.asp"	
Const C_MaxKey   = 16                                           

Dim gblnWinEvent
Dim arrReturn										
Dim arrParam
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
 Function InitVariables()
 	lgStrPrevKey = ""										
 	lgPageNo         = ""
 	lgBlnFlgChgValue = False		
 	lgIntFlgMode = PopupParent.OPMD_CMODE								
 	lgSortKey        = 1			
		
 	gblnWinEvent = False
 	Redim arrReturn(0, 0)
 	Self.Returnvalue = arrReturn
 End Function

'========================================================================================================
 Sub SetDefaultVal()		
 	ArrParam = ArrParent(1)		
 	frm1.txtSONo.Value = arrParam(0)
 	frm1.txtApplicant.value = arrParam(1)
 	frm1.txtApplicantNm.value = arrParam(2)
 	frm1.txtSalesGroup.value = arrParam(3)
 	frm1.txtSalesGroupNm.value = arrParam(4)
 	frm1.txtCurrency.value = arrParam(7)				
 End Sub

'========================================================================================================
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
 	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %> 
 End Sub

'========================================================================================================
 Sub InitSpreadSheet()
	    
     Call SetZAdoSpreadSheet("S4112RA4","S","A","V20030318", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
 	With frm1.vspdData
 		ggoSpread.Source = frm1.vspdData
 		.OperationMode = 5
 		Call SetSpreadLock 
 	End With
	      
 End Sub	

'========================================================================================================
 Sub SetSpreadLock()
     With frm1
     .vspdData.ReDraw = False
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	ggoSpread.SpreadLockWithOddEvenRowColor()
 	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
     .vspdData.ReDraw = True
     End With
	    
 End Sub

'====================================================================================================
Sub CurFormatNumSprSheet()
 	ggoSpread.Source = frm1.vspdData
 	'단가 
 	ggoSpread.SSSetFloatByCellOfCur C_Price,-1, txtCurrency.value, PopupParent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
 	'금액 
 	ggoSpread.SSSetFloatByCellOfCur C_Amount,-1, txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
End Sub


'========================================================================================================
 Function OKClick()
	
 	Dim intColCnt, intRowCnt, intInsRow
		
 	If frm1.vspdData.SelModeSelCount > 0 Then 
 		intInsRow = 0
 		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1 , frm1.vspdData.MaxCols - 1  )
		
			
 		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1 
 			frm1.vspdData.Row = intRowCnt + 1
				
 			If frm1.vspdData.SelModeSelected Then
 				For intColCnt = 0 To 15
 					frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
 					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text						
 				Next
 				intInsRow = intInsRow + 1
					
 			End IF
 		Next
 	End if					
		
 	Self.Returnvalue = arrReturn
 	Self.Close()
	
 End Function	

'========================================================================================================
 Function CancelClick()
 	Redim arrReturn(1,1)
 	arrReturn(0,0) = ""
 	Self.Returnvalue = arrReturn
 	Self.Close()
 End Function
'========================================================================================================
 Function OpenItem()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "품목"							
 	arrParam(1) = "B_ITEM"								
 	arrParam(2) = Trim(frm1.txtItem.value)					
 	arrParam(3) = ""									
 	arrParam(4) = ""									
 	arrParam(5) = "품목"							

 	arrField(0) = "ITEM_CD"								
 	arrField(1) = "ITEM_NM"								
        arrField(2) = "SPEC"

 	arrHeader(0) = "품목"							
 	arrHeader(1) = "품목명"						
        arrHeader(2) = "규격"

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 	"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
		
 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetItem(arrRet)
 	End If
 End Function

'========================================================================================================
Function OpenTrackingNo()
	Dim iCalledAspName
	Dim strRet
	'Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	'2002-10-07 s3135pa1.asp 추가 
	Dim arrTNParam(5), i

	If Len(frm1.txtApplicant.value) Then
		arrTNParam(0) = frm1.txtApplicant.value
	End If
	
	If Len(frm1.txtSalesGroup.value) Then
		arrTNParam(1) = frm1.txtSalesGroup.value
	End If

	If Len(frm1.txtItem.value) Then
		arrTNParam(3) = frm1.txtItem.value
	End If
	
	If Len(frm1.txtItem.value) Then
		arrTNParam(3) = frm1.txtItem.value
	End If
		
	If Len(frm1.txtSONo.value) Then
		arrTNParam(4) = frm1.txtSONo.value
	End If

	arrTNParam(5) = "LD"

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3135pa3")	
	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa3", "x")
		lgIsOpenPop = False
		Exit Function
	End if
	lgIsOpenPop = True

	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrTNParam), _
		"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.value = strRet 
	End If		
		
	frm1.txtTrackingNo.focus
End Function	

'========================================================================================================
 Function SetItem(arrRet)
 	frm1.txtItem.Value = arrRet(0)
 	frm1.txtItemNm.Value = arrRet(1)
 	frm1.txtItem.focus
 End Function		

'========================================================================================================
Function OpenSortPopup()
	
 Dim arrRet
	
 On Error Resume Next
	
 If lgIsOpenPop = True Then Exit Function
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

'========================================================================================================
 Sub Form_Load()
		
 	Call LoadInfTB19029				                                           
 	Call InitSpreadSheet()
 	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.LockField(Document, "N")						
 	Call InitVariables													    		
 	Call SetDefaultVal	
 	Call FncQuery()
		
 End Sub

'========================================================================================================
 Function vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
 		Exit Function
 	End If				
 	If frm1.vspdData.MaxRows > 0 Then
 		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
 			Call OKClick
 		End If
 	End If
 End Function	
'========================================================================================================
 Function vspdData_KeyPress(KeyAscii)
      On Error Resume Next
      If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
         Call OKClick()
      ElseIf KeyAscii = 27 Then
         Call CancelClick()
      End If
 End Function
 
 
'========================================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
 	If OldLeft <> NewLeft Then
 		Exit Sub
 	End If
 	If CheckRunningBizProcess = True Then
 	   Exit Sub
 	End If
 	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
 		If lgPageNo <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbQuery = False Then
 				Exit Sub
 			End if
 		End If
 	End If		 
 End Sub

'========================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                     
    Err.Clear                                                        

	Call ggoOper.ClearField(Document, "2")							
	Call InitVariables												
	
	frm1.txtHLCNo.value = arrParam(8)
    If DbQuery = False Then Exit Function							

    FncQuery = True									
        
End Function

'========================================================================================================
Function DbQuery()
	Dim strVal
		
	Err.Clear															
	DbQuery = False														
		
	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
		
	If frm1.rdoLocalLCFlg1.checked Then
		frm1.txtRadio.value = "L"
	ElseIf frm1.rdoLocalLCFlg2.checked Then
		frm1.txtRadio.value = "N"
	End If	
		
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtItem=" & Trim(frm1.txtHItem.value)			
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtHSONo.value)		
		strVal = strVal & "&txtApplicant=" & Trim(frm1.txtHApplicant.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtHSalesGroup.value)
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)
                strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtHTrackingNo.value)

		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)			
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSONo.value)		
		strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtSalesGroup.value)
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&txtLCNo=" & Trim(frm1.txtHLCNo.value)
                strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)

		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If		
'--------- Developer Coding Part (End) ------------------------------------------------------------
    strVal =     strVal & "&lgPageNo="       & lgPageNo                                                          
	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function

'========================================================================================================
Function DbQueryOk()														

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtItem.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>품목</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SIZE=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="14XXXU" ALT="영업그룹">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
						</TR>
							<TD CLASS=TD5 NOWRAP>주문처</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="주문처">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>수주번호</TD>
							<TD CLASS=TD6><INPUT NAME="txtSONo" ALT="수주번호" TYPE=TEXT MAXLENGTH=18 SIZE=34 TAG="14XXXU" TABINDEX=-1></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>화폐</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="14XXXU" ALT="화폐"></TD>
                                                        <TD CLASS=TD5 NOWRAP>Tracking No.</TD>
							<TD CLASS=TD6><INPUT NAME="txtTrackingNo" ALT="Tracking No." TYPE=TEXT MAXLENGTH=25 SIZE=30 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
                                                </TR>
                                                <TR>  
							<TD CLASS=TD5 NOWRAP>LOCAL L/C 여부</TD> 
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLocalLCFlg" TAG="11X" VALUE="L" CHECKED ID="rdoLocalLCFlg1"><LABEL FOR="rdoLocalLCFlg1">예</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLocalLCFlg" TAG="11X" VALUE="N" ID="rdoLocalLCFlg2"><LABEL FOR="rdoLocalLCFlg2">아니오</LABEL>
							</TD>
                                                        <TD CLASS=TD5 NOWRAP></TD> 
							<TD CLASS=TD6 NOWRAP></TD> 
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
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s4112ra4_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
						<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCOpenDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHTrackingNo" TAG="24">

<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
