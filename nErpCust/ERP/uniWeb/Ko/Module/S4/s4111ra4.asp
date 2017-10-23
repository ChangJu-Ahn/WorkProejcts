<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4111ra4.asp																*
'*  4. Program Name         : Local L/C 출하참조(Local L/C등록에서)										*
'*  5. Program Desc         : Local L/C 출하참조(Local L/C등록에서)										*
'*  6. Comproxy List        : S41118ListDnHdrForLcSvr													*
'*  7. Modified date(First) : 2000/10/11																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>출하참조</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_QRY_ID	= "s4111rb4.asp"			
Const C_MaxKey          = 12                                           

Dim gblnWinEvent											 
Dim arrReturn												
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim prDBSYSDate

Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = PopupParent.UNIConvDateAToB(prDBSYSDate ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)               'Convert DB date type to Company
StartDate = PopupParent.UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

Const gstrPayTermsMajor = "B9004"

'========================================================================================================
 Function InitVariables()
 	lgStrPrevKey		= ""										
 	lgPageNo			= ""
 	lgBlnFlgChgValue	= False									
 	lgIntFlgMode		= PopupParent.OPMD_CMODE								
 	lgSortKey			= 1   
		
 	gblnWinEvent		= False
     Redim arrReturn(0)        
     Self.Returnvalue	= arrReturn(0)     
 End Function
	
'========================================================================================================
 Sub SetDefaultVal()
 	txtFromDt.text = StartDate
 	txtToDt.text = EndDate
 	txtApplicant.focus	  
 End Sub

'========================================================================================================
 Sub LoadInfTB19029()
 <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 <% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
 <% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
 End Sub


'========================================================================================================
 Sub InitSpreadSheet()
		
 	Call SetZAdoSpreadSheet("S4111RA4","S","A","V20030318", PopupParent.C_SORT_DBAGENT, vspdData, C_MaxKey, "X", "X" )
 	With vspdData
 		ggoSpread.Source = vspdData
 		.OperationMode = 3
 		Call SetSpreadLock 
 	End With

 End Sub

'========================================================================================================
 Sub SetSpreadLock()
 	vspdData.ReDraw = False
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	ggoSpread.SpreadLockWithOddEvenRowColor()
 	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
     vspdData.ReDraw = True
  
 End Sub	

'========================================================================================================
 Function OKClick()
 	If vspdData.ActiveRow > 0 Then	
 		vspdData.Row = vspdData.ActiveRow
 		vspdData.Col = GetKeyPos("A",2)
 		arrReturn = Trim(vspdData.Text)

 		Self.Returnvalue = arrReturn
 	End If

 	Self.Close()
 End Function
	

'========================================================================================================
 Function CancelClick()
 	Redim arrReturn(0)
 	arrReturn(0) = ""
 	Self.Returnvalue = arrReturn(0)
 	Self.Close()
 End Function

'========================================================================================================
 Function OpenBizPartner()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "개설신청인"						
 	arrParam(1) = "B_BIZ_PARTNER"						
 	arrParam(2) = Trim(txtApplicant.value)				
 	arrParam(3) = ""									
 	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
 	arrParam(5) = "개설신청인"						

 	arrField(0) = "BP_CD"								
 	arrField(1) = "BP_NM"								

 	arrHeader(0) = "개설신청인"						
 	arrHeader(1) = "개설신청인명"					

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetBizPartner(arrRet)
 	End If
 End Function

'========================================================================================================
 Function OpenMinorCd()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function
 	gblnWinEvent = True
		
 	Dim LCKind
		
 	If rdoLocalLCFlg1.checked Then
 		LCKind = "L"
 	ElseIf rdoLocalLCFlg2.checked Then
 		LCKind = "N"
 	End If	
		
 	arrParam(0) = "결제방법"													
 	arrParam(1) = "b_minor,b_configuration"													
 	arrParam(2) = Trim(txtPayTerms.Value)											
 	arrParam(3) = ""											
 	arrParam(4) = "b_minor.MINOR_CD = b_configuration.MINOR_CD AND b_minor.MAJOR_CD = " & FilterVar(gstrPayTermsMajor, "''", "S") & " AND b_configuration.REFERENCE =  " & FilterVar(LCKind, "''", "S") & ""
 	arrParam(5) = "결제방법"													

 	arrField(0) = "b_minor.Minor_CD"														
 	arrField(1) = "b_minor.Minor_NM"														

 	arrHeader(0) = "결제방법"													
 	arrHeader(1) = "결제방법명"													
		
 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetMinorCd(arrRet)
 	End If
 End Function

'========================================================================================================
 Function OpenSalesGroup()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "영업그룹"								
 	arrParam(1) = "B_SALES_GRP"									
 	arrParam(2) = Trim(txtSalesGroup.value)						
 	arrParam(3) = ""											
 	arrParam(4) = ""											
 	arrParam(5) = "영업그룹"								

 	arrField(0) = "SALES_GRP"									
 	arrField(1) = "SALES_GRP_NM"								

 	arrHeader(0) = "영업그룹"								
 	arrHeader(1) = "영업그룹명"								

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetSalesGroup(arrRet)
 	End If
 End Function

'========================================================================================================
 Function SetBizPartner(arrRet)
 	txtApplicant.Value = arrRet(0)
 	txtApplicantNm.Value = arrRet(1)
 	txtApplicant.focus
 End Function

'========================================================================================================
 Function SetMinorCd(arrRet)
 	txtPayTerms.Value = arrRet(0)
 	txtPayTermsNm.Value = arrRet(1)
 	txtPayTerms.focus
 End Function

'========================================================================================================
 Function SetSalesGroup(arrRet)
 	txtSalesGroup.Value = arrRet(0)
 	txtSalesGroupNm.Value = arrRet(1)
 	txtSalesGroup.focus
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
 	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)		
 	Call ggoOper.LockField(Document, "N")                         
 	Call InitVariables											  
 	Call SetDefaultVal	
 	Call InitSpreadSheet()
 	call FncQuery
	
 End Sub

'========================================================================================================
 Sub btnApplicantOnClick()
 	Call OpenBizPartner()
 End Sub

'========================================================================================================
 Sub btnPayTermsOnClick()
 	Call OpenMinorCd()
 End Sub

'========================================================================================================
 Sub btnSalesGroupOnClick()
 	Call OpenSalesGroup()
 End Sub

'========================================================================================================
 Sub txtFromDt_DblClick(Button)
     If Button = 1 Then
         txtFromDt.Action = 7 
         Call SetFocusToDocument("P")
 		 txtFromDt.Focus
     End If
 End Sub

'========================================================================================================
 Sub txtToDt_DblClick(Button)
     If Button = 1 Then
         txtToDt.Action = 7
         Call SetFocusToDocument("P")
 		 txtToDt.Focus
     End If
 End Sub

'========================================================================================================
 Sub txtFromDt_Keypress(KeyAscii)
 	On Error Resume Next
 	If KeyAscii = 27 Then
 		Call CancelClick()
 	Elseif KeyAscii = 13 Then
 		Call FncQuery()
 	End if
 End Sub

 Sub txtToDt_Keypress(KeyAscii)
 	On Error Resume Next
 	If KeyAscii = 27 Then
 		Call CancelClick()
 	Elseif KeyAscii = 13 Then
 		Call FncQuery()
 	End if
 End Sub

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
 On Error Resume Next
 If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
 	Call OKClick()
 ElseIf KeyAscii = 27 Then
 	Call CancelClick()
 End If
End Function
	
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
 If Row = 0 Or vspdData.MaxRows = 0 Then 
 	Exit Function
 End If
				
 If vspdData.MaxRows > 0 Then
 	If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
 		Call OKClick
 	End If
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
 If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    '☜: 재쿼리 체크	
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

		
 	If Not chkField(Document, "1") Then							
 		Exit Function
 	End If

		
 	If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function

		
 	Call DbQuery()														

 	FncQuery = True														
 End Function

'========================================================================================================
 Function DbQuery()

 	Err.Clear															
 	DbQuery = False														
					
 	If   LayerShowHide(1) = False Then
 	         Exit Function 
 	End If

 	If rdoLocalLCFlg1.checked Then
 		txtRadio.value = "L"
 	ElseIf rdoLocalLCFlg2.checked Then
 		txtRadio.value = "N"
 	End If	

 	Dim strVal

 	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001				
 		strVal = strVal & "&txtApplicant=" & Trim(txtHApplicant.value)	
 		strVal = strVal & "&txtSalesGroup=" & Trim(txtHSalesGroup.value)
 		strVal = strVal & "&txtPayTerms=" & Trim(txtHPayTerms.value)
 		strVal = strVal & "&txtFromDt=" & Trim(txtHFromDt.value)
 		strVal = strVal & "&txtToDt=" & Trim(txtHToDt.value)
 		strVal = strVal & "&txtRadio=" & Trim(txtRadio.value)
 		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	Else
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001				
 		strVal = strVal & "&txtApplicant=" & Trim(txtApplicant.value)	
 		strVal = strVal & "&txtSalesGroup=" & Trim(txtSalesGroup.value)
 		strVal = strVal & "&txtPayTerms=" & Trim(txtPayTerms.value)
 		strVal = strVal & "&txtFromDt=" & Trim(txtFromDt.text)
 		strVal = strVal & "&txtToDt=" & Trim(txtToDt.text)
 		strVal = strVal & "&txtRadio=" & Trim(txtRadio.value)
 		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	End If
		
 	strVal = strVal & "&lgPageNo="		 & lgPageNo						        
 	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
 	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
 	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
				
 	Call RunMyBizASP(MyBizASP, strVal)									

 	DbQuery = True														
 End Function
	
'========================================================================================================
 Function DbQueryOk()													
 	lgIntFlgMode = PopupParent.OPMD_UMODE											
		
 	If vspdData.MaxRows > 0 Then
 		vspdData.Focus
 		vspdData.Row = 1	
 		vspdData.SelModeSelected = True		
 	Else
 		txtApplicant.focus
 	End If
 End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>개설신청인</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="개설신청인"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
							</TD>
						</TR>	
							<TD CLASS=TD5>결제방법</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayTermsOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5>수주일</TD>
							<TD CLASS=TD6>
								<script language =javascript src='./js/s4111ra4_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s4111ra4_fpDateTime2_txtToDt.js'></script>
							</TD>	
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
							<script language =javascript src='./js/s4111ra4_vaSpread_vspdData.js'></script>
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
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHIncoTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="24">

<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
