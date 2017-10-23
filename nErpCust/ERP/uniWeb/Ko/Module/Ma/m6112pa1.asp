<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Inventory															*
'*  2. Function Name        : DocumentNo Popup																*
'*  3. Program ID           :   i1111pa1.asp																	*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수불번호팝업																	*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2001/12/14																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Lee Seung Wook																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/02/29 : Coding Start													*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "m6112pb1.asp"							 '☆: 비지니스 로직 ASP명 
Const C_SHEETMAXROWS = 30								 '--- 한화면에 보일수 있는 최대 Row 수 
										'--- Index of Textbox Name 
Const C_DocumentNo = 1
Const C_Year = 2
Const C_DocumentDt = 3
Const C_MovType = 4
Const C_Plant = 5
Const C_DocumentText = 6


Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 
Dim lgDocumentNo
Dim lgYear
Dim lgFromDt
Dim lgToDt
Dim lgMovType
Dim lgTrnsType
Dim lgStrPrevKey


Dim hlgDocumentNo				 '--- Hidden View for Re Query 
Dim hlgYear
Dim hlgFromDt
Dim hlgToDt
Dim hlgMovType
Dim hlgTrnsType
Dim hlgPlantCd      
Dim hlgPlantNm  


Dim arrParam					 '--- First Parameter Group 
Dim arrReturn				 '--- Return Parameter Group 
Dim lgBlnFlgChgValue	

EndDate = GetSvrDate                                          '☆: 초기화면에 뿌려지는 시작 날짜 -----
StartDate = UniDateAdd("m", -1, EndDate,gServerDateFormat)    '☆: 초기화면에 뿌려지는 시작 날짜 -----
EndDate   = UniConvDateAToB(EndDate  ,gServerDateFormat,gDateFormat)
StartDate = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat) 

arrParam = window.dialogArguments

top.document.title = "수불번호팝업"
'=======================================================================================================
Function LoadInfTB19029()
	<!-- #Include file="../../ComAsp/ComLoadInfTB19029.asp" -->
End Function
'=======================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=EndDate%>",gServerDateFormat,gServerDateType,strYear,strMonth,strDay)
	frm1.txtDocumentNo.value = arrParam(0)
	
	If arrParam(1) = "" then
		frm1.hdnYear.value = strYear
	Else
		frm1.hdnYear.value = arrParam(1)
	End if
	
	frm1.hdnTrnsType.Value = arrParam(2)
	frm1.txtToDt.Text   = "<%=EndDate%>"
	frm1.txtFromDt.Text = "<%=StartDate%>"
	hlgPlantCd = arrParam(3)
	
	Self.Returnvalue = Array("")
End Sub
'=======================================================================================================
Sub InitSpreadSheet()

	frm1.vspdData.ReDraw = False
	frm1.vspdData.OperationMode = 3
	frm1.vspdData.MaxCols = C_DocumentText + 1
	frm1.vspdData.Col 	 = frm1.vspdData.MaxCols
	frm1.vspdData.ColHidden = True
	frm1.vspdData.MaxRows = 0
	    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit
	ggoSpread.SSSetEdit C_DocumentNo, "수불번호", 16
	ggoSpread.SSSetEdit C_Year, "년도", 8,2	
	ggoSpread.SSSetEdit C_DocumentDt, "수불일자", 10,2
	ggoSpread.SSSetEdit C_MovType, "이동유형", 10,2
	ggoSpread.SSSetEdit C_Plant, "공장", 8,2
	ggoSpread.SSSetEdit C_DocumentText, "비고", 27	

    Call SetSpreadLock 

	frm1.vspdData.ReDraw = True
End Sub

'==============================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=======================================================================================================	
Function OKClick()
	Dim intColCnt
	
	If frm1.vspdData.ActiveRow > 0 Then	
		Redim arrReturn(frm1.vspdData.MaxCols)
	
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
				
		For intColCnt = 0 To frm1.vspdData.MaxCols - 1
			frm1.vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next
		arrReturn(intColCnt) = hlgPlantNm
		Self.Returnvalue = arrReturn		
	End If
	
	Self.Close()
End Function
'=======================================================================================================
Function CancelClick()
	Self.Close()
End Function
'=======================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
'=======================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
	Call InitVariables
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, gComNum1000, gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, gDateFormat, gComNum1000, gComNumDec)
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call FncQuery()
	Call MM_preloadImages("../../image/Query.gif","../../image/OK.gif","../../image/Cancel.gif")
	
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
Sub txtFromDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("P")	
        frm1.txtToDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtToDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
    End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function
'=======================================================================================================
Function txtFromDt_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
	    Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function
'=======================================================================================================
Function txtToDt_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
	    Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

'=======================================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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
Function FncQuery()

    frm1.vspdData.MaxRows = 0

	lgQueryFlag = "1"
	lgDocumentNo = Trim(frm1.txtDocumentNo.Value)
	lgYear = Trim(frm1.hdnYear.value)
	lgFromDt = frm1.txtFromDt.Text
	lgToDt = frm1.txtToDt.Text
	lgStrPrevKey = ""
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If DbQuery() = False Then
		Exit Function
	End if
	
End Function
'=======================================================================================================
Function DbQuery()

	Dim strVal
	Dim txtMaxRows
	'Show Processing Bar
   	Call LayerShowHide(1)  
	DbQuery = False 
	txtMaxRows = frm1.vspdData.MaxRows
	if lgStrPrevKey <> "" Then
		strVal = BIZ_PGM_ID & "?txtDocumentNo=" & hlgDocumentNo
		strVal = strVal     & "&txtYear="       & hlgYear
		strVal = strVal     & "&txtFromDt="     & hlgFromDt
		strVal = strVal     & "&txtToDt="       & hlgToDt
		strVal = strVal     & "&txtMovType="    & "R40"
		strVal = strVal     & "&txtTrnsType="   & "PR"
		strVal = strVal     & "&txtPlantCd="    & hlgPlantCd
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal     & "&txtMaxRows"     & txtMaxRows		
	else
		strVal = BIZ_PGM_ID & "?txtDocumentNo=" & lgDocumentNo
		strVal = strVal     & "&txtYear="       & lgYear
		strVal = strVal     & "&txtFromDt="     & lgFromDt
		strVal = strVal     & "&txtToDt="       & lgToDt
		strVal = strVal     & "&txtMovType="    & "R40"
		strVal = strVal     & "&txtTrnsType="   & "PR"
		strVal = strVal     & "&txtPlantCd="    & hlgPlantCd
		strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal     & "&txtMaxRows"     & txtMaxRows	
	end if                                                        '⊙: Processing is NG
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True                                                          '⊙: Processing is NG
End Function
'=======================================================================================================
Function DbQueryOk()								'☆: 조회 성공후 실행로직 
	frm1.vspdData.Focus
End Function
'=======================================================================================================
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	

</HEAD>
<!--
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
-->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET CLASS="CLSFLD">
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP>수불번호</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtDocumentNo" SIZE=20 MAXLENGTH=16 tag="11xxxU" ></TD>
				<TD CLASS="TD5" NOWRAP>수불일자</TD>
				<TD CLASS="TD6" NOWRAP> <script language =javascript src='./js/m6112pa1_I913700675_txtFromDt.js'></script>
                                         &nbsp;~&nbsp;
                                        <script language =javascript src='./js/m6112pa1_I285492295_txtToDt.js'></script>
			</TR>
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/m6112pa1_I611981451_vspdData.js'></script>
	</TD></TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnYear" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrnsType" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

