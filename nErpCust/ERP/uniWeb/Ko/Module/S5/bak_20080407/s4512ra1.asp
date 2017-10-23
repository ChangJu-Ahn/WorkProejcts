<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하요청내역 
'*  3. Program ID           : s4512ra1
'*  4. Program Name         : 출하요청내역 참조 
'*  5. Program Desc         : 출하내역등록에서 출하요청내역 참조 Popup
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/07
'*  8. Modified date(Last)  : 2002/05/07
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
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

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s4512rb1.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 81                                           '☆: key count of SpreadSheet
Const C_PopItemCd		= 1

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim lgArrReturn												'☜: Return Parameter Group
Dim lgBlnOpenedFlag
Dim	lgBlnItemCdChg
Dim gStrPlantCd

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
        
    lgBlnItemCdChg = False
End Function

'========================================
Sub SetDefaultVal()
	Dim arrRowSep, arrColValue

	With frm1
		arrRowSep = Split(arrParent(1), PopupParent.gRowSep)
		arrColValue = Split(arrRowSep(0),PopupParent.gColSep)

'		.txtFromDt.Text = UNIDateClientFormat(UniConvDateAToB(UniConvDateToYYYYMM(arrColValue(8), PopupParent.gDateFormat, "-") & "-01", PopupParent.gServerDateFormat ,PopupParent.gAPDateFormat))
'		.txtToDt.Text = arrColValue(8)	

		<% '수주번호 %>
		.txtSoNo.value			= arrColValue(0)
		<% '주문처 %>
		.txtPtnBpCd.value	= arrColValue(1)
		.txtPtnBpNm.value = arrColValue(2)
		<% '공장 %>
		gStrPlantCd = arrColValue(3)
		<% '영업그룹 %>
'		.txtSalesGrpCd.value	= arrColValue(3)
'		.txtSalesGrpNm.value	= arrColValue(4)
		<% '결제방법 %>
'		.txtPayTermsCd.value	= arrColValue(5)
'		.txtPayTermsNm.value	= arrColValue(6)
		<% '화폐 %>
'		.txtHCurrency.value		= arrColValue(7)

		<% '추가 %>
'		.txtHBillDt.value		= arrColValue(8)
'		.txtHVatRate.value		= arrColValue(9)
'		.txtHVatType.value		= arrColValue(10)
'		.txtHVatCalcType.value	= arrColValue(11)
'		.txtHVatIncflag.value	= arrColValue(12)
'		.txtHXchgRate.value		= arrColValue(13)
'		.txtHXchgOp.value		= arrColValue(14)		
		<% '매출채권형태 %>
'		.txtHBillTypeCd.value	= arrColValue(15)
			
	End With
	Redim lgArrReturn(0,0)
	Self.Returnvalue = lgArrReturn
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S4512RA1","S","A","V20051118",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
							C_MaxKey, "X","X")		
	Call SetSpreadLock 	 
	    
End Sub

'========================================
Sub SetSpreadLock()
'	ggoSpread.SpreadLock 1 , -1
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 5
End Sub	

'========================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow
	
	With frm1	
		If .vspdData.SelModeSelCount > 0 Then
			intInsRow = 0

			Redim lgArrReturn(.vspdData.SelModeSelCount, .vspdData.MaxCols)

			For intRowCnt = 1 To .vspdData.MaxRows

				.vspdData.Row = intRowCnt
				If .vspdData.SelModeSelected Then
					For intColCnt = 1 To .vspdData.MaxCols
						.vspdData.Col =  GetKeyPos("A", intColCnt)
						lgArrReturn(intInsRow, intColCnt - 1) = .vspdData.Text
					Next

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
	End With
		
	Self.Returnvalue = lgArrReturn
	Self.Close()

End Function

'========================================
Function CancelClick()
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
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
    
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedflag = True

	DbQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function OpenDnReqNo()
	Dim iCalledAspName
	Dim strRet
	If gblnWinEvent = True Then Exit Function
			
	gblnWinEvent = True

	iCalledAspName = AskPRAspName("S4511PA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4511PA1", "x")
		gblnWinEvent = False
		exit Function
	end if
		
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.PopupParent), _
		"dialogWidth=646px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	frm1.txtDnReqNo.focus
			
	If strRet <> "" Then
		frm1.txtDnReqNo.value = strRet
	End If	

End Function


'=============================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	Case 0		
		arrParam(0) = "수주번호"										' TextBox 명칭 
		arrParam(1) = "S_SO_HDR SO, B_BIZ_PARTNER BP"						' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtSoNo.value)								' Code Condition
		arrParam(4) = "SO.SOLD_TO_PARTY=BP.BP_CD AND SO.CFM_FLAG = " & FilterVar("Y", "''", "S") & " "		' Where Condition
		arrParam(5) = "수주번호"										' TextBox 명칭 
			
		arrField(0) = "SO.SO_NO"											' Field명(0)
		arrField(1) = "BP.BP_NM"											' Field명(1)
    
		arrHeader(0) = "수주번호"										' Header명(0)
		arrHeader(1) = "주문처"											' Header명(1)
		
	Case 1
		arrParam(0) = "주문처"		
		arrParam(1) = "B_BIZ_PARTNER"
		arrParam(2) = Trim(frm1.txtPtnBpCd.value)		
		arrParam(3) = ""					
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag = " & FilterVar("Y", "''", "S") & " "		
		arrParam(5) = "주문처"		
			
		arrField(0) = "BP_CD"							
		arrField(1) = "BP_NM"				
		
	    arrHeader(0) = "주문처"							
	    arrHeader(1) = "주문처명"						

	Case 2

		arrParam(0) = "영업그룹"						
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGrp.value)		
		arrParam(3) = ""									
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						

		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						

		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"						
	    
    
	Case 3

		If UCase(frm1.txtPlant.className) = PopupParent.UCN_PROTECTED Then 
			gblnWinEvent = False			
			Exit Function
		End IF
		
		arrParam(0) = "공장"				
		arrParam(1) = "B_PLANT"							
		arrParam(2) = Trim(frm1.txtPlant.value)		
		arrParam(4) = ""							
		arrParam(5) = "공장"				
		
		arrField(0) = "PLANT_CD"				
		arrField(1) = "PLANT_NM"				
	    
		arrHeader(0) = "공장"					
		arrHeader(1) = "공장명"				

	Case 4
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
		
	End Select


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'=============================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtSoNo.value = arrRet(0) 		
			.txtSoNo.focus
		Case 1
			.txtPtnBpCd.value = arrRet(0) 
			.txtPtnBpNm.value = arrRet(1)   
			.txtPtnBpCd.focus
		Case 2
			.txtSalesGrp.value = arrRet(0)
			.txtSalesGrpNm.value = arrRet(1)  
			.txtSalesGrp.focus
		Case 3
			.txtPlant.value = arrRet(0) 
			.txtPlantNm.value = arrRet(1) 
			.txtPlant.focus 		 
		Case 4
			.txtItem.value = arrRet(0) 
			.txtItemNm.value = arrRet(1)  
			.txtItem.focus		 
		End Select
	End With
End Function


'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================
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

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")
		frm1.txtFromDt.Focus
	End If
End Sub

'========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToDt.Focus
	End If
End Sub

'========================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

	End With
   
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call InitVariables 														'⊙: Initializes local global variables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	
	'원래는 Get방식이나 조건부가 많으면 POST방식으로 넘김 

    With frm1

		.txtHMode.Value = PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			.txtHFromDt.value = .txtHFromDt.value
			.txtHToDt.value = .txtHToDt.value
		Else
			' 처음 조회시 
			.txtHFromDt.value = .txtFromDt.Text
			.txtHToDt.value = .txtToDt.Text
		End If

		.txtHPtnBpCd.Value = .txtPtnBpCd.Value
		.txtHSoNo.Value = .txtSoNo.Value
		.txtHSalesGrp.value = .txtSalesGrp.value
		.txtHDnReqNo.Value = .txtDnReqNo.Value
		.txtHItem.Value = .txtItem.Value
		.txtHPlantCd.value = gStrPlantCd

		.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
		.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
		.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))

        .txtHlgPageNo.value	= lgPageNo

	End with

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True
    If Err.number = 0 Then
       DbQuery = True																
    End If   


End Function

'=========================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	With frm1
		If .vspdData.MaxRows > 0 Then
			If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
				lgIntFlgMode = PopupParent.OPMD_UMODE
				.vspdData.Row = 1	
				.vspdData.SelModeSelected = True
			End If
			.vspdData.Focus
		Else
			Call SetFocusToDocument("P")
			.txtFromDt.focus
		End If
	End With
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
						<TD CLASS=TD5>출고일</TD>
						<TD CLASS=TD6>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS="FPDTYYYYMMDD" tag="11X1" Alt="출고시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS="FPDTYYYYMMDD" tag="11X1" Alt="출고종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
						<TD CLASS=TD5 NOWRAP>납품처</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPtnBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPtnBpNm" SIZE=20 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="Vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>수주번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" SIZE=34 MAXLENGTH=18 TAG="11XXXU" ALT="S/O번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItem" SIZE=10 MAXLENGTH=20 TAG="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="Vbscript:OpenConSItemDC 4">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>출하요청번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtDnReqNo" SIZE=34 MAXLENGTH=20 TAG="11XXXU" ALT="출하요청번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoRef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDnReqNo">&nbsp;</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=YES NORESIZE framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPtnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHDnReqNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItem" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="24">


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
