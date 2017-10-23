<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111PA3
'*  4. Program Name         : 출하관리번호 팝업 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Date표준적용 
'*                            2002/12/17 Include 성능향상 강준구 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>출하번호</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop
Dim lgBlnDnFlgChecked

Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s4111pb3.asp"
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value
 
'=========================================
Sub InitVariables()
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1

End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtDlvyFrDt.text = StartDate
	frm1.txtDlvyToDt.text = EndDate
	lgBlnDnFlgChecked = False
End Sub

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

End Sub

'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4111pa1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "납품처"							
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.txtBpCd.value)				
	arrParam(3) = ""									
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
	arrParam(5) = "납품처"							
		
	arrField(0) = "BP_CD"								
	arrField(1) = "BP_NM"								
		
	arrHeader(0) = "납품처"							
	arrHeader(1) = "납품처명"						
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	frm1.txtBpCd.focus
	
	If arrRet(0) <> "" Then
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
	End If
End Function

'=========================================
Function OpenMinorCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "출하형태"						
	arrParam(1) = "b_minor A, I_MOVETYPE_CONFIGURATION B"	
	arrParam(2) = Trim(frm1.txtMovType.value)					
	arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
	arrParam(5) = "출하형태"						

	arrField(0) = "A.MINOR_CD"							
	arrField(1) = "A.MINOR_NM"							

	arrHeader(0) = "출하형태"						
	arrHeader(1) = "출하형태명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtMovType.focus

	If arrRet(0) <> "" Then
		frm1.txtMovType.value = arrRet(0)
		frm1.txtMovTypeNm.value = arrRet(1)
	End If
End Function

'=========================================
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "영업그룹"								
	arrParam(1) = "B_SALES_GRP"									
	arrParam(2) = Trim(frm1.txtSalesGroup.value)						
	arrParam(3) = ""											
	arrParam(4) = ""											
	arrParam(5) = "영업그룹"								

	arrField(0) = "SALES_GRP"									
	arrField(1) = "SALES_GRP_NM"										

	arrHeader(0) = "영업그룹"								
	arrHeader(1) = "영업그룹명"								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtSalesGroup.focus

	If arrRet(0) <> "" Then
		frm1.txtSalesGroup.Value = arrRet(0)
		frm1.txtSalesGroupNm.Value = arrRet(1)
	End If
End Function

'========================================
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

'========================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'============================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Sub ZbtnBpCd_OnClick()
	Call OpenBizPartner()
End Sub

'========================================
Sub ZbtnSalesGroup_OnClick()
	Call OpenSalesGroup()
End Sub

'========================================
Sub ZbtnMovType_OnClick()
	Call OpenMinorCd()
End Sub

'========================================
Sub txtDlvyFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyFrDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDlvyFrDt.Focus
    End If
End Sub

'========================================
Sub txtDlvyToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyToDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDlvyToDt.Focus
    End If
End Sub

'========================================
Sub txtDlvyFrDt_Keypress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 27 Then
        Call CancelClick()
     Elseif KeyAscii = 13 Then
        Call FncQuery()
     End if
End Sub

'========================================
Sub txtDlvyToDt_Keypress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 27 Then
        Call CancelClick()
     Elseif KeyAscii = 13 Then
        Call FncQuery()
     End if
End Sub

'========================================
Sub rdoDNFlg1_OnClick()
	If Not lgBlnDnFlgChecked Then
		lgBlnDnFlgChecked = True
		idDateTitle.innerHTML = "출고예정일"
	End If
End Sub

'========================================
Sub rdoDNFlg2_OnClick()
	If lgBlnDnFlgChecked Then
		lgBlnDnFlgChecked = False
		idDateTitle.innerHTML = "출고일"
	End If
End Sub

'========================================
Sub rdoDNFlg3_OnClick()
	If Not lgBlnDnFlgChecked Then
		lgBlnDnFlgChecked = True
		idDateTitle.innerHTML = "출고예정일"
	End If
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Sub
	
'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If CheckRunningBizProcess Then Exit Sub
		If lgStrPrevKey <> "" Then Call DbQuery
	End If		 
End Sub

'========================================
Function OKClick()
		
	dim arrReturn
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1) ' 1
		arrReturn = frm1.vspdData.Text

		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	If ValidDateCheck(frm1.txtDlvyFrDt, frm1.txtDlvyToDt) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
	If frm1.rdoDNFlg1.checked = True Then
		frm1.txtRadio.value = "A"
	ElseIf frm1.rdoDNFlg2.checked = True Then
		frm1.txtRadio.value = "Y"
	ElseIf frm1.rdoDNFlg3.checked = True Then
		frm1.txtRadio.value = "N"
	End If			   	

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=====================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)	
		strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
		strVal = strVal & "&txtMovType=" & Trim(.txtMovType.value)
		strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
		strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
		strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)										

    End With
    
    DbQuery = True


End Function

'=====================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBpCd.focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
						<TD CLASS=TD5 NOWRAP>납품처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnBpCd_OnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnSalesGroup_OnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>	
						<TD CLASS=TD5 NOWRAP>출하형태</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtMovType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU" ALT="출하형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnMovType_OnClick">&nbsp;
							<INPUT NAME="txtMovTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
						</TD>
						<TD CLASS=TD5 id="idDateTitle" NOWRAP>출고일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s4111pa3_fpDateTime2_txtDlvyFrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s4111pa3_fpDateTime2_txtDlvyToDt.js'></script>
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>출고여부</TD> 
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="A" ID="rdoDNFlg1"><LABEL FOR="rdoDNFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="Y" CHECKED ID="rdoDNFlg2"><LABEL FOR="rdoDNFlg2">출고</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="N"  ID="rdoDNFlg3"><LABEL FOR="rdoDNFlg3">미출고</LABEL>			
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
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s4111pa3_vaSpread_vspdData.js'></script>
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
							                  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
