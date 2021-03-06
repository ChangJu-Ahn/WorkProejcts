<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3112ra6.asp																*
'*  4. Program Name         : 수주내역참조(통관내역등록에서)											*
'*  5. Program Desc         : 수주내역참조(통관내역등록에서)											*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2002/05/08																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'*                            2. 2002/05/08 : Ado 변환													*
'********************************************************************************************************

%>
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
<SCRIPT LANGUAGE="VBScript">

Option Explicit		
'========================================================================================================
Const BIZ_PGM_ID 		= "s3112rb6.asp"                              
'========================================================================================================
Const C_MaxKey       = 15                                           

<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

Dim IsOpenPop  
Dim gblnWinEvent

Dim arrReturn										
Dim arrParam	
Dim arrParent
Dim lgIsOpenPop

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'========================================================================================================
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
'========================================================================================================
Sub SetDefaultVal()

	arrParam = arrParent(1)

	frm1.txtApplicant.value = arrParam(0)
	frm1.txtApplicantNm.value = arrParam(1)
	frm1.txtSONo.Value = arrParam(2)
	frm1.txtCurrency.value = arrParam(3)	
	frm1.txtSalesGroup.value = arrParam(4)
	frm1.txtSalesGroupNm.value = arrParam(5)
	frm1.txtPayTerms.value = arrParam(6)
	frm1.txtPayTermsNm.value = arrParam(7)
	frm1.txtIncoTerms.value = arrParam(8)	
	frm1.txtIncoTermsNm.value = arrParam(9)	
		
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	Dim strRef
	Dim strChk
		
	Call CommonQueryRs("REFERENCE ","B_CONFIGURATION","MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if lgF0 <> "" Then
		strRef = Split(lgF0, Chr(11))
		strChk = strRef(0)
	else
		StrChk = ""
	End If
	IF strChk <> "Y" Then 
		Call SetZAdoSpreadSheet("S3112RA6","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")		
	Else		
		Call SetZAdoSpreadSheet("S3112RA602","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")				
	End IF					
		
	Call SetSpreadLock 		
	      
End Sub
'========================================================================================================
Sub SetSpreadLock()
	frm1.vspdData.OperationMode = 5
	ggoSpread.SpreadLockWithOddEvenRowColor
End Sub
'========================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 			
		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1

			frm1.vspdData.Row = intRowCnt + 1

			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
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
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
		frm1.txtItem.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	End If

End Function	
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenTrackingNo()
	Dim iCalledAspName
	Dim strRet

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

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

	arrTNParam(5) = "CS"

	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3135pa3")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa3", "x")
		lgIsOpenPop = False
		exit Function
	end if
	gblnWinEvent = True

	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrTNParam), _
		"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.value = strRet 
	End If		
		
	frm1.txtTrackingNo.focus
End Function	

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029				                                           
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
			
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	
	Call ggoOper.LockField(Document, "N")					
	Call InitVariables														    		
	Call SetDefaultVal		
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	
End Sub	
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)

	If Row = 0 Then Exit Function
	If frm1.vspdData.MaxRows = 0 Then Exit Function
	If Row > 0 Then Call OKClick()

End Function
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    		
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

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgPageNo <> "" Then		                                                    
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'==========================================================================================
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
'=======================================================================================================
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
'========================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    Err.Clear    
	
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
					
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call InitVariables	

    If DbQuery = False Then Exit Function							

    FncQuery = True									
        
End Function
'********************************************************************************************************
Function DbQuery()

	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)

	'--------- Developer Coding Part (Start) ----------------------------------------------------------		
	With frm1
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then        
	
			.txtHItem.value			=  Trim(frm1.txtItem.value)			
			.txtHFromDt.value		=  Trim(frm1.txtFromDt.text)
			.txtHToDt.value			=  Trim(frm1.txtToDt.text)
			.txtHApplicant.value	=  Trim(frm1.txtApplicant.value)
			.txtHSONo.value			=  Trim(frm1.txtSONo.value)
			.txtHCurrency.value		=  Trim(frm1.txtCurrency.value)
			.txtHSalesGroup.value	=  Trim(frm1.txtSalesGroup.value)
			.txtHPayTerms.value		=  Trim(frm1.txtPayTerms.value)
			.txtHIncoTerms.value	=  Trim(frm1.txtIncoTerms.value)						
			.txtHTrackingNo.value	=  Trim(frm1.txtTrackingNo.value)									
			
			.txtHlgSelectListDT.value	=  GetSQLSelectListDataType("A")			 
			.txtHlgTailList.value		=  MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		=  EnCoding(GetSQLSelectList("A"))	
		End If
		
		.txtHlgPageNo.value	= lgPageNo      
	End With
'--------- Developer Coding Part (End) ------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then
       DbQuery = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   
    
End Function
'========================================================================================
Function DbQueryOk()														<%%>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtItem.focus
	End If
End Function
'===========================================================================
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

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
							<TD CLASS=TD5 NOWRAP>수주일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3112ra6_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s3112ra6_fpDateTime2_txtToDt.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>주문처</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="주문처">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>수주번호</TD>
							<TD CLASS=TD6><INPUT NAME="txtSONo" ALT="수주번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="14XXXU" TABINDEX=-1></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>화폐</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="14XXXU" ALT="화폐"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="14XXXU" ALT="영업그룹">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5>결제방법</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=5 TAG="14XXXU" ALT="결제방법">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5>가격조건</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=5 TAG="14XXXU" ALT="가격조건">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="14"></TD>
						</TR>	
						<TR>
							<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
							<TD CLASS=TD6><INPUT NAME="txtTrackingNo" ALT="Tracking 번호" TYPE=TEXT MAXLENGTH=25 SIZE=30 TAG="11XXXU" TABINDEX=-1><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTrackingNo()"></TD>
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
							<script language =javascript src='./js/s3112ra6_vaSpread_vspdData.js'></script>
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
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>
							<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
						<TD WIDTH=30% ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX=-1 ></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHIncoTerms" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHTrackingNo" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHlgTailList" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHlgPageNo" TAG="24" TABINDEX=-1>	
		
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
