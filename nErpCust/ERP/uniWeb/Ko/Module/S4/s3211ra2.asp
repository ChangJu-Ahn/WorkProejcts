<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3211ra2.asp																*
'*  4. Program Name         : Local L/C참조(Local L/C Amend등록에서)									*
'*  5. Program Desc         : Local L/C참조(Local L/C Amend등록에서)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/04																*
'*  8. Modified date(Last)  : 2002/04/25																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/04 : 화면 design												*
'*                            2. 2002/04/25 : Ado 변환													*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>LOCAL L/C참조</TITLE>

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

Const BIZ_PGM_ID 		= "s3211rb2.asp"   
Const C_MaxKey          = 12                                           
Const gstrPayTermsMajor = "B9004"
	
Dim lgIsOpenPop
Dim gblnWinEvent											 
Dim arrReturn												
Dim strReturn					

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

'========================================================================================================
 Function InitVariables()
 	lgStrPrevKey     = ""								   
 	lgPageNo         = ""
     lgBlnFlgChgValue = False	                           
     lgIntFlgMode     = PopupParent.OPMD_CMODE                          
     lgSortKey        = 1   
        
     gblnWinEvent = False
     Redim arrReturn(0)        
     Self.Returnvalue = arrReturn (0)    
 End Function

'========================================================================================================
 Sub SetDefaultVal()
 	frm1.txtFromDt.text = StartDate
 	frm1.txtToDt.text = EndDate
 	frm1.txtApplicant.focus	  
 End Sub
		
'========================================================================================================
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
 	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %> 
 End Sub

'========================================================================================================
 Sub InitSpreadSheet()
	
 	Call SetZAdoSpreadSheet("S3211RA2","S","A","V20030318", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
 	With frm1.vspdData
 		ggoSpread.Source = frm1.vspdData
 		.OperationMode = 3
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

'========================================================================================================
 Function OKClick()
			
 	If frm1.vspdData.ActiveRow > 0 Then	
 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = GetKeyPos("A",1)
 		strReturn = Trim(frm1.vspdData.Text)

 		Self.Returnvalue = strReturn
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
Function OpenConSItemDC(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If gblnWinEvent = True Then Exit Function
 gblnWinEvent = True
	
 Select Case iWhere
 Case 0
 	arrParam(1) = "B_BIZ_PARTNER"						
 	arrParam(2) = Trim(frm1.txtApplicant.value)				
 	arrParam(3) = ""									
 	arrParam(4) = ""									
 	arrParam(5) = "개설신청인"							

 	arrField(0) = "BP_CD"								
 	arrField(1) = "BP_NM"								

 	arrHeader(0) = "개설신청인"							
 	arrHeader(1) = "개설신청인명"								

 Case 1
 	arrParam(1) = "B_SALES_GRP"									
 	arrParam(2) = Trim(frm1.txtSalesGroup.value)						
 	arrParam(3) = ""											
 	arrParam(4) = ""											
 	arrParam(5) = "영업그룹"								

 	arrField(0) = "SALES_GRP"									
 	arrField(1) = "SALES_GRP_NM"										

 	arrHeader(0) = "영업그룹"								
 	arrHeader(1) = "영업그룹명"										
 Case 2		        	
 	arrParam(1) = "S_LC_HDR"												
 	arrParam(2) = Trim(frm1.txtLCDocNo.value)									
 	arrParam(3) = ""														
 	arrParam(4) = "LC_KIND=" & FilterVar("L", "''", "S") & " "												
 	arrParam(5) = "L/C번호"												

 	arrField(0) = "LC_DOC_NO"							
 	arrField(1) = "CONVERT(char(3),LC_AMEND_SEQ)"		

 	arrHeader(0) = "L/C번호"											
 	arrHeader(1) = "AMEND차수"											

	
 Case 3		
 	arrParam(1) = "B_CURRENCY"														
 	arrParam(2) = Trim(frm1.txtCurrency.Value)											
 	arrParam(3) = ""																
 	arrParam(4) = ""																
 	arrParam(5) = "화폐"													

 	arrField(0) = "Currency"														
 	arrField(1) = "Currency_desc"														

 	arrHeader(0) = "화폐"													
 	arrHeader(1) = "화폐명"													
	
 case 4
			
 	arrParam(1) = "B_Bank"								
 	arrParam(2) = Trim(frm1.txtOpenBank.value)													
 	arrParam(3) = ""																			
 	arrParam(4) = ""									
 	arrParam(5) = "은행"							
	
 	arrField(0) = "Bank_cd"								
 	arrField(1) = "BANK_NM"							
	   
 	arrHeader(0) = "은행"							
 	arrHeader(1) = "은행명"							
 End Select

 arrParam(0) = arrParam(5)											
	
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 gblnWinEvent = False

 If arrRet(0) = "" Then
 	Exit Function
 Else
 	Call SetConSItemDC(arrRet, iWhere)
 End If	
	
End Function

'========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
 With frm1
 	Select Case iWhere
 	Case 0			
 		.txtApplicant.Value = arrRet(0)
 		.txtApplicantNm.Value = arrRet(1)
 		.txtApplicant.focus
 	Case 1
 		.txtSalesGroup.Value = arrRet(0)
 		.txtSalesGroupNm.Value = arrRet(1)
 		.txtSalesGroup.focus
 	Case 2
 		.txtLCDocNo.Value = arrRet(0)
 		.txtLCDocNo.focus
 	Case 3
 		.txtCurrency.Value = arrRet(0)
 		.txtCurrency.focus
 	case 4
 		.txtOpenBank.value = arrRet(0)
 		.txtOpenBankNm.value = arrRet(1)
 		.txtOpenBank.focus
 	End Select
 End With
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
 		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbQuery = False Then
 				Exit Sub
 			End if
 		End If
 	End If		 
 End Sub
	
'========================================================================================================
 Sub txtFromDt_DblClick(Button)
     If Button = 1 Then
         frm1.txtFromDt.Action = 7 
         Call SetFocusToDocument("P")
         frm1.txtFromDt.Focus
     End If
 End Sub

 Sub txtToDt_DblClick(Button)
     If Button = 1 Then
         frm1.txtToDt.Action = 7
         Call SetFocusToDocument("P")
         frm1.txtToDt.Focus
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
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
		
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function	

    Call ggoOper.ClearField(Document, "2")	         						
    Call InitVariables 														
   
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================================================================================
Function DbQuery() 

	Err.Clear														
	DbQuery = False													
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1		
		
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtApplicant=" & Trim(.txtHApplicant.value)	
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtHSalesGroup.value)
			strVal = strVal & "&txtLCDocNo=" & Trim(.txtHLCDocNo.value)
			strVal = strVal & "&txtCurrency=" & Trim(.txtHCurrency.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtOpenBank=" & Trim(.txtHOpenBank.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtApplicant=" & Trim(.txtApplicant.value)	
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtLCDocNo=" & Trim(.txtLCDocNo.value)
			strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
			strVal = strVal & "&txtOpenBank=" & Trim(.txtOpenBank.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If
		
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						              
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Call RunMyBizASP(MyBizASP, strVal)		    						
        
    End With
    
    DbQuery = True    
End Function

'========================================================================================================
Function DbQueryOk()	    												

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtApplicant.focus
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
							<TD CLASS=TD5 NOWRAP>개설신청인</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="개설신청인"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>				
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCDocNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 2"></TD>
							<TD CLASS=TD5 NOWRAP>화폐</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="11XXXU" ALT="화폐"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 3"></TD>
						</TR>	
						<TR>
							<TD CLASS=TD5 NOWRAP>개설은행</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="11XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 4" >&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5>개설일</TD>
							<TD CLASS=TD6>
								<script language =javascript src='./js/s3211ra2_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s3211ra2_fpDateTime2_txtToDt.js'></script>
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
							<script language =javascript src='./js/s3211ra2_vaSpread_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtHLCDocNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHOpenBank" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
