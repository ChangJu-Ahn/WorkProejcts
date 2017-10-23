<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211PA3
'*  4. Program Name         : 판매계획기간 Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/12/27
'*  8. Modified date(Last)  : 2002/12/27
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

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

Option Explicit                              '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "S2211PB3.ASP"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 7                                            '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 

Dim lgArrParent
Dim lgStrInitQuery
Dim lgStrSpType

lgArrParent = window.dialogArguments
Set PopupParent = lgArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'========================================================================================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
End Function

'=======================================================================================================
Sub SetDefaultVal()
	Dim iArrParam
	Dim iArrReturn
	
	iArrParam = lgArrParent(1)
	
	With frm1
		.txtConSpPeriod.value = iArrParam(0)
		lgStrInitQuery = iArrParam(0)
		
		<%If Trim(Request("txtDisplayFlag")) = "Y" Then%>
		.txtConLastClosedSpPeriod.value		= iArrParam(1)
		.txtConLastClosedSpPeriodDesc.value = iArrParam(2)
		If iArrParam(3) = "" Then
			.txtFromDt.Text = EndDate
		Else
			.txtFromDt.Text = iArrParam(3)
		End If
		lgStrInitQuery = lgStrInitQuery & iArrParam(1)
		<%Else%>
		.txtFromDt.Text = EndDate
		<%End If%>
		If UBound(iArrParam) >= 4 Then
			lgStrSpType = iArrParam(4)
		Else
			lgStrSpType = "E"
		End If
		.vspdData.OperationMode = 3
	End With
	
	Redim iArrReturn(0)
	Self.Returnvalue = iArrReturn
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "PA") %>	
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S2211PA3","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	Call SetSpreadLock 	
	    
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
	Dim iIntCol
	Dim iArrReturn
	
	With frm1.vspdData
		If .ActiveRow > 0 Then	
			Redim iArrReturn(C_MaxKey)
			.Row = .ActiveRow

			For iIntCol = 0 To C_MaxKey - 1
				.Col = GetKeyPos("A",iIntCol + 1)
				iArrReturn(iIntCol) = .Text
			Next
			
			Self.Returnvalue = iArrReturn
		End If
	End With
	Err.Clear
	
	Self.Close()
End Function

'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
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

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
   
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	If Trim(lgStrInitQuery) <> "" Then DbQuery()
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=======================================================================================================
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

'=======================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'=======================================================================================================
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

'=======================================================================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")   
		frm1.txtFromDt.Focus
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

'=======================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    'If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
    '   Exit Function
    'End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================================================================================
Function DbQuery() 
	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtSpType=" & lgStrSpType
			strVal = strVal & "&txtSpPeriod=" & Trim(.txtHSpPeriod.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			
			<%If Trim(Request("txtDisplayFlag")) = "Y" Then%>
			strVal = strVal & "&txtLastClosedSpPeriod=" & Trim(.txtConLastClosedSpPeriod.value)
			<%Else%>
			strVal = strVal & "&txtLastClosedSpPeriod="
			<%End If%>

		Else
			' 처음 조회시 
			strVal = strVal & "&txtSpType=" & lgStrSpType
			strVal = strVal & "&txtSpPeriod=" & Trim(.txtConSpPeriod.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
			
			<%If Trim(Request("txtDisplayFlag")) = "Y" Then%>
			strVal = strVal & "&txtLastClosedSpPeriod=" & Trim(.txtConLastClosedSpPeriod.value)
			<%Else%>
			strVal = strVal & "&txtLastClosedSpPeriod="
			<%End If%>
		End If
		
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True    

End Function

'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
	Else
		frm1.txtConSpPeriod.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>계획기간</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConSpPeriod" ALT="계획기간" SIZE=15 MAXLENGTH=8 TAG="11XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtConSpPeriodDesc" SIZE=25 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>시작일</TD>
						<TD CLASS="TD6" NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s2211pa3_fpDateTime1_txtFromDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<%If Trim(Request("txtDisplayFlag")) = "Y" Then%>
					<TR>
						<TD CLASS=TD5>최종마감기간</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConLastClosedSpPeriod" ALT="최종마감기간" SIZE=15 MAXLENGTH=8 TAG="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtConLastClosedSpPeriodDesc" SIZE=25 TAG="14"></TD>
					</TR>
					<%End If%>
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
						<script language =javascript src='./js/s2211pa3_OBJECT1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											  <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG>			</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHSpPeriod" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
                                                                                                                    