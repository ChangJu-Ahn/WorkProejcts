<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : s3112pa1.asp
'*  4. Program Name         : 품목팝업(판매계획등록)
'*  5. Program Desc         : 품목팝업(판매계획등록)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
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
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<% ' 이미지 체인지 관련 자바스크립트 함수  %>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

Dim lgBlnFlgChgValue                                        <%'☜: Variable is for Dirty flag            %>
Dim lgStrPrevKey                                            <%'☜: Next Key tag                          %>
Dim lgSortKey                                               <%'☜: Sort상태 저장변수                     %> 
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

Dim lgSelectList                                            <%'☜: SpreadSheet의 초기  위치정보관련 변수 %>
Dim lgSelectListDT                                          <%'☜: SpreadSheet의 초기  위치정보관련 변수 %>

Dim lgTypeCD                                                <%'☜: 'G' is for group , 'S' is for Sort    %>
Dim lgFieldCD                                               <%'☜: 필드 코드값                           %>
Dim lgFieldNM                                               <%'☜: 필드 설명값                           %>
Dim lgFieldLen                                              <%'☜: 필드 폭(Spreadsheet관련)              %>
Dim lgFieldType                                             <%'☜: 필드 설명값                           %>
Dim lgDefaultT                                              <%'☜: 필드 기본값                           %>
Dim lgNextSeq                                               <%'☜: 필드 Pair값                           %>
Dim lgKeyTag                                                <%'☜: Key 정보                                %>
Dim lgNextSeq_T                                             <%'☜: 필드 Pair값                           %>
Dim lgKeyTag_T                                              <%'☜: Key 정보                              %>

Dim lgSortTitleNm                                           <%'☜: Orderby popup용 데이타(필드설명)      %>
Dim lgSortFieldCD1                                          <%'☜: Orderby popup용 데이타(필드코드)      %>

Dim lgPopUpR                                                <%'☜: Orderby default 값                    %>
Dim lgMark                                                  <%'☜: 마크                                  %>
Dim lgKeyPos                                                <%'☜: Key위치                               %>
Dim lgKeyPosVal                                             <%'☜: Key위치 Value                         %>
Dim IscookieSplit 

Dim arrReturn					<% '--- Return Parameter Group %>

Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3112pb1.asp"
Const C_SHEETMAXROWS    = 25                                   '☆: Spread sheet에서 보여지는 row
Const C_SHEETMAXROWS_D  = 30                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 7                                    '☆☆☆☆: Max key value
                                            '☆: Jump시 Cookie로 보낼 Grid value
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtItem.value = ArrParent(1)
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "PA") %> 
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3112pa1","S","A","V20021210", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub

'========================================================================================================= 
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================= 
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "품목"							
	arrParam(1) = "B_ITEM"								
	arrParam(2) = Trim(frm1.txtItem.value)				
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "품목"							
		
	arrField(0) = "ITEM_CD"								
	arrField(1) = "ITEM_NM"								
		
	arrHeader(0) = "품목"							
	arrHeader(1) = "품목명"							
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItem(arrRet)
	End If
End Function

'========================================================================================================= 
Function OpenJnlItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "품목계정"								
	arrParam(1) = "A_JNL_ITEM"									
	arrParam(2) = Trim(frm1.txtJnlItem.value)					
	arrParam(3) = ""											
	arrParam(4) = "JNL_TYPE = " & FilterVar("IA", "''", "S") & ""								
	arrParam(5) = "품목계정"								

	arrField(0) = "JNL_CD"										
	arrField(1) = "JNL_NM"										

	arrHeader(0) = "품목계정"								
	arrHeader(1) = "품목계정명"								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetJnlItem(arrRet)
	End If
End Function

'========================================================================================================= 
Function SetItem(arrRet)
	frm1.txtItem.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
End Function

'========================================================================================================= 
Function SetJnlItem(arrRet)
	frm1.txtJnlItem.value = arrRet(0)
	frm1.txtJnlItemNm.value = arrRet(1)
End Function

'========================================================================================================= 
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/zpConfig.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'========================================================================================================= 
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							<% '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
				DbQuery
			End If
		End If
	End With
End Sub
	
'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
   End if
    
End Sub

'========================================================================================================= 
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then   'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function


'========================================================================================================= 
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

'========================================================================================================= 
Function OKClick()
		
	Redim arrReturn(3)
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		'frm1.vspdData.Col = 1
		frm1.vspdData.Col = Getkeypos("A",1)
		arrReturn(0) = frm1.vspdData.Text
		'frm1.vspdData.Col = 2
		frm1.vspdData.Col = Getkeypos("A",2)
		arrReturn(1) = frm1.vspdData.Text
		'frm1.vspdData.Col = 6
		frm1.vspdData.Col = Getkeypos("A",6)			
		arrReturn(2) = frm1.vspdData.Text
		'frm1.vspdData.Col = 3
		frm1.vspdData.Col = Getkeypos("A",3)			
		arrReturn(3) = frm1.vspdData.Text
			
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'========================================================================================================= 
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================= 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'========================================================================================================= 
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)		<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtItemNm=" & Trim(frm1.txtItemNm.value)	<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtJnlItem=" & Trim(frm1.txtJnlItem.value)
		
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '☜: 한번에 가져올수 있는 데이타 건수 

		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
       	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function

'========================================================================================================= 
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtItem.focus
	End If

End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtItem" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="품목"></TD>
						<TD CLASS=TD5 NOWRAP>품목계정</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtJnlItem" TYPE="Text" MAXLENGTH="20" SIZE=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnJnlItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenJnlItem">&nbsp;
							<INPUT NAME="txtJnlItemNm" TYPE="Text" SIZE=20 tag="24">
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>품목명</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" TYPE="Text" SIZE=30 MAXLENGTH="50" ALT="품목명" tag="11"></TD>
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
						<script language =javascript src='./js/s3112pa1_vaSpread_vspdData.js'></script>
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
					<TD>&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"     ONCLICK="FncQuery()"     ></IMG>
									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG></TD>
					<TD ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"        ONCLICK="OkClick()"      ></IMG>
							        <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"    ONCLICK="CancelClick()"  ></IMG></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
