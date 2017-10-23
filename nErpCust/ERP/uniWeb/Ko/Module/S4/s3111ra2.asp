<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3111ra2.asp																*
'*  4. Program Name         : 수주참조(Local L/C등록)													*
'*  5. Program Desc         : 수주참조(Local L/C등록)													*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>수주참조</TITLE>

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

Const BIZ_PGM_QRY_ID = "s3111rb2.asp"	
Const C_MaxKey          = 16                                           

Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
Dim arrReturn												'☜: Return Parameter Group
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
Const gstrIncoTermsMajor = "B9006"

'========================================================================================================
 Function InitVariables()
 	lgStrPrevKey	 = ""										
 	lgPageNo         = ""
 	lgBlnFlgChgValue = False									
 	lgIntFlgMode	 = PopupParent.OPMD_CMODE								
 	lgSortKey        = 1   
		
 	gblnWinEvent	 = False
     Redim arrReturn(0)        
     Self.Returnvalue = arrReturn(0)     
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
		
 	Call SetZAdoSpreadSheet("S3111RA1","S","A","V20030318", PopupParent.C_SORT_DBAGENT, vspdData, C_MaxKey, "X", "X" )
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
 	ReDim strReturn(1)
		
 	If vspdData.ActiveRow > 0 Then
 		vspdData.Row = vspdData.ActiveRow
 		vspdData.Col = GetKeyPos("A",1)
 		strReturn(0) = Trim(vspdData.Text)
			
 		Self.Returnvalue = strReturn(0)
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

 	arrParam(0) = "주문처"							
 	arrParam(1) = "B_BIZ_PARTNER"						
 	arrParam(2) = Trim(txtApplicant.value)				
 	arrParam(3) = ""									
 	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
 	arrParam(5) = "주문처"							

 	arrField(0) = "BP_CD"								
 	arrField(1) = "BP_NM"								

 	arrHeader(0) = "주문처"							
 	arrHeader(1) = "주문처명"						

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
 Function OpenMinorCd(ByVal i)
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)
 	Dim LCKind
		
 	If rdoLocalLCFlg1.checked Then
 		LCKind = "L"
 	ElseIf rdoLocalLCFlg2.checked Then
 		LCKind = "N"
 	End If	
		
 	If gblnWinEvent = True Then Exit Function
 	gblnWinEvent = True
		
 	arrParam(1) = "b_minor,b_configuration"													
		
 	arrField(0) = "b_minor.Minor_CD"														
 	arrField(1) = "b_minor.Minor_NM"														
		
 	Select Case i
 		Case 0
 			arrParam(0) = "결제방법"													
 			arrParam(2) = Trim(txtPayTerms.Value)											
 			arrParam(3) = ""																
 			arrParam(4) = "b_minor.MINOR_CD = b_configuration.MINOR_CD AND b_minor.MAJOR_CD = " & FilterVar(gstrPayTermsMajor, "''", "S") & " AND b_configuration.REFERENCE =  " & FilterVar(LCKind, "''", "S") & ""
 			arrParam(5) = "결제방법"													

 			arrHeader(0) = "결제방법"													
 			arrHeader(1) = "결제방법명"													
		
 		Case 1 
 			arrParam(0) = "가격조건"												
 			arrParam(2) = Trim(txtIncoTerms.Value)										
 			arrParam(3) = ""															
 			arrParam(4) = "B_MINOR.MAJOR_CD= " & FilterVar(gstrIncoTermsMajor, "''", "S") & ""				
 			arrParam(5) = "가격조건"												

 			arrHeader(0) = "가격조건"												
 			arrHeader(1) = "가격조건명"												
		
 	End Select
		
 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetMinorCd(arrRet, i)
 	End If
 End Function

'========================================================================================================
 Function OpenSOType()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "수주형태"								
 	arrParam(1) = "S_SO_TYPE_CONFIG"							
 	arrParam(2) = Trim(txtSOType.value)							
 	arrParam(3) = ""											
 	arrParam(4) = ""											
 	arrParam(5) = "수주형태"								

 	arrField(0) = "SO_TYPE"										
 	arrField(1) = "SO_TYPE_NM"										

 	arrHeader(0) = "수주형태"								
 	arrHeader(1) = "수주형태명"								

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetSOType(arrRet)
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
 Function SetMinorCd(arrRet, Byval i)
 	Select Case i
 		Case 0
 			txtPayTerms.Value = arrRet(0)
 			txtPayTermsNm.Value = arrRet(1)
 			txtPayTerms.focus
 		Case 1
 			txtIncoTerms.value = arrRet(0)	
 			txtIncoTermsNm.value = arrRet(1)
 			txtIncoTerms.focus
 	End Select		
 End Function

'========================================================================================================
 Function SetSOType(arrRet)
 	txtSOType.Value = arrRet(0)
 	txtSOTypeNm.Value = arrRet(1)
 	txtSOType.focus
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
		
 	Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
 	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
 	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
 	Call InitVariables											  '⊙: Initializes local global variables
 	Call SetDefaultVal	
 	Call InitSpreadSheet()
 	Call FncQuery
		
 End Sub


'========================================================================================================
 Sub btnApplicantOnClick()
 	Call OpenBizPartner()
 End Sub


'========================================================================================================
 Sub btnPayTermsOnClick()
 	Call OpenMinorCd(0)
 End Sub

'========================================================================================================
 Sub btnSOTypeOnClick() 
 	Call OpenSOType()
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

'==========================================================================================
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
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
 	With vspdData
 		If Row >= NewRow Then
 			Exit Sub
 		End If

 		If NewRow = .MaxRows Then
 			If lgStrPrevKey <> "" Then							
 				DbQuery
 			End If
 		End If
 	End With
 End Sub
	

'========================================================================================================
 Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
 	If OldLeft <> NewLeft Then
 	  	Exit Sub
 	End If
 	If CheckRunningBizProcess = True Then
 	   Exit Sub
 	End If
 	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	  '☜: 재쿼리 체크	
 		If lgPageNo <> "" Then                   '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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
 		strVal = strVal & "&txtSOType=" & Trim(txtHSOType.value)
 		strVal = strVal & "&txtPayTerms=" & Trim(txtHPayTerms.value)
 		strVal = strVal & "&txtFromDt=" & Trim(txtHFromDt.value)
 		strVal = strVal & "&txtToDt=" & Trim(txtHToDt.value)
 		strVal = strVal & "&txtRadio=" & Trim(txtRadio.value)
 		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	Else
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001				
 		strVal = strVal & "&txtApplicant=" & Trim(txtApplicant.value)	
 		strVal = strVal & "&txtSalesGroup=" & Trim(txtSalesGroup.value)
 		strVal = strVal & "&txtSOType=" & Trim(txtSOType.value)
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
							<TD CLASS=TD5>주문처</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
							</TD>
						</TR>	
							<TD CLASS=TD5>수주형태</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtSOType" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSOType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSOTypeOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtSOTypeNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5>수주일</TD>
							<TD CLASS=TD6>
								<script language =javascript src='./js/s3111ra2_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s3111ra2_fpDateTime2_txtToDt.js'></script>
							</TD>	
						</TR>
						<TR>
							<TD CLASS=TD5>결제방법</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayTermsOnClick()">&nbsp;
								<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5 NOWRAP>LOCAL L/C 여부</TD> 
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLocalLCFlg" TAG="11X" VALUE="L" CHECKED ID="rdoLocalLCFlg1"><LABEL FOR="rdoLocalLCFlg1">예</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLocalLCFlg" TAG="11X" VALUE="N" ID="rdoLocalLCFlg2"><LABEL FOR="rdoLocalLCFlg2">아니오</LABEL>
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
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s3111ra2_vaSpread_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtHSOType" TAG="24">
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
