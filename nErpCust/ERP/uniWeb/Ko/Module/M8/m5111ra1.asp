<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 구매 
'*  2. Function Name        : 매입참조 
'*  3. Program ID           : m5111ra1
'*  4. Program Name         : 매입참조 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/21
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Oh chang won
'* 10. Modifier (Last)      : Lee Eun Hee 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date 표준적용 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>매입참조</TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

Dim lgBlnFlgChgValue                                        <%'☜: Variable is for Dirty flag            %>
Dim lgStrPrevKey                                            <%'☜: Next Key tag                          %>
Dim lgSortKey                                               <%'☜: Sort상태 저장변수                     %> 
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim lgPageNo
Dim IscookieSplit 
Dim lgIntFlgMode		
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParam	
Const ivType = "ST"

Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName


Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


Const BIZ_PGM_ID        = "m5111rb1.asp"
Const C_MaxKey          = 32                                    '☆☆☆☆: Max key value		' ==== 2005.07.14 Lot No. Lot Sub No. 추가 ====
Const gstPaytermsMajor = "B9004"
 
'==========================================  2.1 InitVariables()  ======================================
Sub InitVariables()
	Dim arrParam
    
	arrParam = arrParent(1)
    
    frm1.hdnSupplierCd.value 	= arrParam(0)
	frm1.hdnCurr.value 		    = arrParam(1)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.hdnRefPoNo.value       = arrParam(3)
	'반품내역등록 - 외주가공여부 조건 추가 
	If UBound(arrparam) > 3 Then
		frm1.hdnSubcontraflg.value	= arrParam(4)
	End if
     
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo         = ""
    lgSortKey        = 1
    lgIntFlgMode = PopupParent.OPMD_CMODE	
	
	Redim arrReturn(0, 0)
	Self.Returnvalue = arrReturn
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtFrIvDt.text = StartDate
	frm1.txtToIvDt.text = EndDate
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
Sub InitSpreadSheet()
   Call SetZAdoSpreadSheet("M5111RA1","S","A","V20050714",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
   Call SetSpreadLock
   frm1.vspdData.OperationMode = 5
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	
	Dim strRet
'	Dim arrParam(0)
	Dim arrParam(1)
	Dim iCalledAspName
	
		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = ivType
		arrParam(1) = "Y"
		
		iCalledAspName = AskPRAspName("m5111pa1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m5111pa1", "X")
			lgIsOpenPop = False
			Exit Function
		End If

		strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False
	
		If strRet(0) = "" Then
			frm1.txtIvNo.focus
			Exit Function
		Else
			frm1.txtIvNo.value = strRet(0)
			frm1.txtIvNo.focus
		End If	
		Set gActiveElement = document.activeElement
End Function
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables							
	Call SetDefaultVal	
	Call InitSpreadSheet()	
		
	Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
	
End Sub

'------------------------------------------  OpenSortPopup()  -------------------------------------------------
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
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
'==================================== 3.2.23 txtToIvDt_DblClick()  =====================================
Sub txtToIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToIvDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToIvDt.focus
	End If
End Sub
'==================================== 3.2.23 txtFrIvDt_DblClick()  =====================================
Sub txtFrIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrIvDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrIvDt.focus
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrIvDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToIvDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
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
'======================================  3.3.3 vspdData_TopLeftChange()  ================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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

'======================================  3.3.4 vspdData_KeyPress()  ================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1

			frm1.vspdData.Row = intRowCnt + 1

			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next

				intInsRow = intInsRow + 1
					
			End IF
		Next
	End if			
		
	Self.Returnvalue = arrReturn
		
	Self.Close()
End Function
'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", PopupParent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If  

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables
  

	If ValidDateCheck(frm1.txtFrIvDt, frm1.txtToIvDt) = False Then Exit Function

    If DbQuery = False Then Exit Function

    FncQuery = True		
    Set gActiveElement = document.activeElement
End Function
'=====================================  DbQuery()  ==========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
       If lgIntFlgMode = PopupParent.OPMD_UMODE Then

		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
		    strVal = strVal & "&txtIvNo=" & Trim(frm1.hdnIvNo.value)				<%'☆: 조회 조건 데이타 %>
		    strVal = strVal & "&txtFrIvDt=" & Trim(frm1.hdnFrIvDt.value)
		    strVal = strVal & "&txtToIvDt=" & Trim(frm1.hdnToIvDt.value)
		    strVal = strVal & "&hdnRefPoNo= " & Trim(frm1.hdnRefPoNo.value)
		    strVal = strVal & "&hdnSupplierCd= " & Trim(frm1.hdnSupplierCd.value)
		    strVal = strVal & "&hdnGroupCd= " & Trim(frm1.hdnGroupCd.value)
			strVal = strVal & "&txtSubcontraflg=" & Trim(frm1.hdnSubcontraflg.value)		' 외주가공여부 추가 
		    strVal = strVal & "&hdnCurr= " & Trim(frm1.hdnCurr.value)
        Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
		    strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIvNo.value)				<%'☆: 조회 조건 데이타 %>
		    strVal = strVal & "&txtFrIvDt=" & Trim(frm1.txtFrIvDt.text)
		    strVal = strVal & "&txtToIvDt=" & Trim(frm1.txtToIvDt.text)
		    strVal = strVal & "&hdnRefPoNo= " & Trim(frm1.hdnRefPoNo.value)
		    strVal = strVal & "&hdnSupplierCd= " & Trim(frm1.hdnSupplierCd.value)
		    strVal = strVal & "&hdnGroupCd= " & Trim(frm1.hdnGroupCd.value)
			strVal = strVal & "&txtSubcontraflg=" & Trim(frm1.hdnSubcontraflg.value)		' 외주가공여부 추가 
		    strVal = strVal & "&hdnCurr= " & Trim(frm1.hdnCurr.value)
       End if   
            strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
            strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    End With
    
    DbQuery = True

End Function
'=====================================  DbQueryOk()  ==========================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtIvNo.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
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
						<TD CLASS=TD5>매입번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=20  MAXLENGTH=18 TAG="11NXXU" ALT="매입번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvNoPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIvNo()">
						<div STYLE="DISPLAY: none"><INPUT NAME="txtIvNo1" STYLE="BORDER-RIGHT: 0px solid;BORDER-TOP: 0px solid;BORDER-LEFT: 0px solid;BORDER-BOTTOM: 0px solid" TYPE="Text" SIZE=1 DISABLED=TRUE Tag="11"></div>
						</TD>
						<TD CLASS=TD5 NOWRAP>매입등록일</TD>
						<TD CLASS=TD6 NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/m5111ra1_fpDateTime1_txtFrIvDt.js'></script>
									</TD>
									<TD>
										~
									</TD>
									<TD>
										<script language =javascript src='./js/m5111ra1_fpDateTime2_txtToIvDt.js'></script>
									</TD>
								</TR>
							</TABLE>
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
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m5111ra1_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToIvDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrIvDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnCurr" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
