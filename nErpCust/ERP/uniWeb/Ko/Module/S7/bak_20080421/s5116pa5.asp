<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5116pa5
'*  4. Program Name         : 매출채권상세 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>매출채권상세</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

<!-- #Include file="../../inc/lgvariables.inc" -->

' External ASP File
'========================================

Const BIZ_PGM_ID 		= "s5116pb5.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 10                                          '☆: key count of SpreadSheet

' User-defind Variables
'========================================
Dim IsOpenPop  

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	

Dim arrPopupParent
Dim PopupParent

ArrPopupParent = window.dialogArguments
Set PopupParent  = ArrPopupParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================
Function InitVariables()
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
			
	ReDim arrReturn(0)
	Self.Returnvalue = arrReturn

End Function

'========================================
Sub SetDefaultVal()

	arrParam = ArrPopupParent(1)

'---------------------------	
' arrParam(0) : 조회From일자 
' arrParam(1) : 조회End일자 
' arrParam(2) : 사업장 
' arrParam(3) : 거래처 
' arrParam(4) : 매출형태 
' arrParam(5) : 확정여부 
' arrParam(6) : 사업장명 
' arrParam(7) : 거래처명 
' arrParam(8) : 매출형태명 
'---------------------------	
	With frm1
	
		.txtConFromDt.text = arrParam(0)
		.txtConToDt.text = arrParam(1)						

		.txtConBizArea.value = arrParam(2)		
		.txtConBpCd.value = arrParam(3)
		.txtConBillType.value = arrParam(4)
		.txtConRdoFlag.value = arrParam(5)		

		If arrParam(5) = "Y" Then
			.rdoConf.checked = True
			.txtConRdoFlag.value = .rdoConf.value
		ElseIf arrParam(5) = "N" Then
			.rdoNonConf.checked = True
			.txtConRdoFlag.value = .rdoNonConf.value
		End If

		.txtConBizAreaNm.value = arrParam(6)		
		.txtConBpNm.value = arrParam(7)
		.txtConBillTypeNm.value = arrParam(8)

	End With

End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("s5116pa5","S","A","V20030529", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock 

End Sub

'========================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 5
End Sub

'========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
		
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	'Call DbQuery()
	Call FncQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then Call DbQuery
	End If
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

    Call InitVariables 														'⊙: Initializes local global variables
    Call SetDefaultVal	
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	With frm1
		.txtMode.value = PopupParent.UID_M0001	

		.txtHConFromDt.value	= .txtConFromDt.text
		.txtHConToDt.value		= .txtConToDt.text	

		.txtHConBizArea.value		= .txtConBizArea.value			
		.txtHConBpCd.value			= .txtConBpCd.value	
		.txtHConBillType.value		= .txtConBillType.value

		.txtHConRdoFlag.value		= .txtConRdoFlag.value

		.txtHConBizAreaNm.value		= .txtConBizAreaNm.value			
		.txtHConBpNm.value			= .txtConBpNm.value	
		.txtHConBillTypeNm.value	= .txtConBillTypeNm.value
						
		.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
		.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
		.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))

        .txtHlgPageNo.value	= lgPageNo

	End With
      
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
	DbQuery = True														<%'⊙: Processing is NG%>
End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
						<TD CLASS="TD5" NOWRAP>매출채권일</TD>									
						<TD CLASS="TD6" NOWRAP>							        
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>
								<script language =javascript src='./js/s5116pa5_OBJECT1_txtConFromDt.js'></script>
								</TD>
								<TD>
								&nbsp;~&nbsp;
								</TD>
								<TD>
								<script language =javascript src='./js/s5116pa5_OBJECT2_txtConToDt.js'></script>
								</TD>
							</TR>
						</TABLE>							        
					    </TD>
					    <TD CLASS="TD5" NOWRAP>사업장</TD>
						<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizArea" SIZE=10 MAXLENGTH=10 tag="14NXXU" ALT="사업장">
												<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>거래처</TD>
						<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBpCd" SIZE=10 MAXLENGTH=10 tag="14NXXU" ALT="거래처">
												<INPUT TYPE=TEXT NAME="txtConBpNm" SIZE=20 tag="14" ALT="거래처명"></TD>
						<TD CLASS="TD5" NOWRAP>매출채권형태</TD>
						<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBillType" SIZE=10 MAXLENGTH=5 tag="14NXXU" ALT="매출채권형태">
												<INPUT TYPE=TEXT NAME="txtConBillTypeNm" SIZE=20 tag="14" ALT="매출채권형태명"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>확정여부</TD>
						<TD CLASS=TD6 NOWRAP>
							<input type=radio CLASS="RADIO" name="rdoConfFlag" id="rdoConf" value="Y" tag = "14X">
								<label for="rdoConf">확정</label>&nbsp;&nbsp;
							<input type=radio CLASS = "RADIO" name="rdoConfFlag" id="rdoNonConf" value="N" tag = "14X">
								<label for="rdoNonConf">미확정</label>
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
					<TD CLASS=TD5 NOWRAP>매출채권총자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s5116pa5_fpDoubleSingle2_txtTotalAmt.js'></script>
					</TD>															
					<TD CLASS=TD5 NOWRAP>매출채권자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s5116pa5_fpDoubleSingle2_txtBillAmt.js'></script>
					</TD>															

				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>VAT자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s5116pa5_fpDoubleSingle2_txtVatAmt.js'></script>
					</TD>									
					<TD CLASS=TD5 NOWRAP>VAT자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s5116pa5_fpDoubleSingle2_txtCollectAmt.js'></script>
					</TD>									
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>적립금자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s5116pa5_fpDoubleSingle2_txtDepositAmt.js'></script>
					</TD>									
					<TD CLASS=TD5 NOWRAP></TD>
					<TD CLASS=TD6 NOWRAP></TD>									
				</TR>

			
				<TR>				
					<TD HEIGHT="100%" colspan=4>
						<script language =javascript src='./js/s5116pa5_vaSpread_vspdData.js'></script>
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
						            	     <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG></TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" 
							 onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtConRdoFlag"		tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConBizArea"	tag="24" TABINDEX="-1">				
<INPUT TYPE=HIDDEN NAME="txtHConBpCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBillType"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConRdoFlag"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConBizAreaNm"	tag="24" TABINDEX="-1">				
<INPUT TYPE=HIDDEN NAME="txtHConBpNm"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBillTypeNm"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtMode" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
