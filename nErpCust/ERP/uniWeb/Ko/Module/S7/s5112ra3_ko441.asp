<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5112RA3
'*  4. Program Name         : 매출채권내역참조 
'*  5. Program Desc         : 세금계산서 내역등록에서 매출내역 참조 Popup
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/06/03
'*  8. Modified date(Last)  : 2002/06/03
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
<TITLE>매출채권내역참조</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5112rb3_ko441.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 19                                           '☆: key count of SpreadSheet
Const C_PopItemCd		= 1
Const C_PopSalesGrp		= 2

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

' User-defind Variables
'========================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim lgArrReturn												'☜: Return Parameter Group
Dim lgBlnSalesGrpChg
Dim lgBlnItemCdChg

Dim arrPopupParent
Dim PopupParent

ArrPopupParent = window.dialogArguments
Set PopupParent  = ArrPopupParent(0)
'20021228 kangjungu dynamic popup
top.document.title = PopupParent.gActivePRAspName

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
        
	lgBlnSalesGrpChg = False		' 영업그룹 변경여부 
	lgBlnItemCdChg	 = False		' 품목코드 변경여부 
End Function

'========================================
	Sub SetDefaultVal()
		Dim arrRowSep, arrColValue

		With frm1
			arrRowSep = Split(ArrPopupParent(1),PopupParent.gRowSep)
			arrColValue = Split(arrRowSep(0),PopupParent.gColSep)

			.txtFromDt.Text = UNIGetFirstDay(arrColValue(9), PopupParent.gDateFormat)
			.txtToDt.Text = arrColValue(9)	

			'발행처 
			.txtBilltoParty.value	= arrColValue(0)
			.txtBilltoPartyNm.value	= arrColValue(1)
	 	    '화폐단위 
			.txtCurrency.value		= arrColValue(2)
			'VAT 유형 
			.txtVatType.value		= arrColValue(3)
			.txtVatTypeNm.value		= arrColValue(4)
	    	'영업그룹 
			.txtSalesGrp.value	= arrColValue(5)
			.txtSalesGrpNm.value	= arrColValue(6)
			'매출채권번호 
			.txtBillNo.value		= arrColValue(7)
			'부가세 포함여부 
			.txtHVatIncFlag.value	= arrColValue(8)
			'발행일 
			.txtHIssueDt.value		= arrColValue(9)
		End With
		Redim lgArrReturn(0,0)
		Self.Returnvalue = ""
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGrp, "Q") 
        	frm1.txtSalesGrp.value = lgSGCd
	End If	

	End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("s5112ra3","S","A","V20030301", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	Call SetSpreadLock 
	    
End Sub

'========================================
Sub SetSpreadLock()
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
					For intColCnt = 1 To .vspdData.MaxCols - 1
						.vspdData.Col = GetKeyPos("A", intColCnt)
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
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopItemCd
		iArrParam(1) = "b_item A"									' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtItemCd.value)					' Code Condition
		iArrParam(3) = ""											' Name Cindition
		iArrParam(4) = ""											' Where Condition
		iArrParam(5) = "품목"									' TextBox 명칭 

		iArrField(0) = "ED15" & PopupParent.gColSep & "A.item_cd"	' Field명(0)
		iArrField(1) = "ED30" & PopupParent.gColSep & "A.item_nm"	

		iArrHeader(0) = "품목"									' Header명(0)
		iArrHeader(1) = "품목명"
		
		frm1.txtItemCd.focus

	Case C_PopSalesGrp
                If frm1.txtSalesGrp.className = "protected" Then
                	IsOpenPop = False
                        Exit Function
                End If
		iArrParam(1) = "dbo.B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"

		frm1.txtSalesGrp.focus
	End Select
 
	iArrParam(0) = iArrParam(5)

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
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
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopItemCd
		frm1.txtItemCd.value = pvArrRet(0) 
		frm1.txtItemNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
    Call InitVariables
        Call GetValue_ko441()											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	DbQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function

'========================================
Function txtItemCd_OnKeyDown()
	lgBlnItemCdChg = True
	lgBlnFlgChgValue = True
End Function

'	Description : 조회조건의 유효성을 Check한다.
'   주의사항 : 화면의 tab order 별로 기술한다. 
'========================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnItemCdChg Then
		iStrCode = Trim(frm1.txtItemCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("IT", "''", "S") & "", C_PopItemCd) Then
				Call DisplayMsgBox("970000", "X", frm1.txtItemCd.alt, "X")
				frm1.txtItemCd.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtItemNm.value = ""
		End If
		lgBlnItemCdChg = False
	End If

End Function

'	Name : GetCodeName()
'	Description : 코드값에 해당하는 명을 Display한다.
'========================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		'GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then
			If CheckRunningBizProcess Then Exit Sub
			Call DbQuery
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
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'========================================
Sub txtToDt_Keypress(KeyAscii)
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
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFromDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHIssueDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, .txtHIssueDt.alt & "(" & .txtHIssueDt.value & ")")
			.txtFromDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(.txtHIssueDt.value, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, .txtHIssueDt.alt & "(" & .txtHIssueDt.value & ")")	
			.txtToDt.Focus()
			Exit Function
		End If
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
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtHItemCd.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtHSalesGrp.value)
		Else
			' 처음 조회시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타 %>
			If Len(Trim(.txtToDt.text)) Then
				strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			Else
				strVal = strVal & "&txtToDt=" & Trim(.txtHIssueDt.value)
			End if
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
		End If
		strVal = strVal & "&txtBilltoParty=" & Trim(.txtBilltoParty.value)
		strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
		strVal = strVal & "&txtBillNo=" & Trim(.txtBillNo.value)
		strVal = strVal & "&txtVatType=" & Trim(.txtVatType.value)
		strVal = strVal & "&txtVatCalcType=" & Trim(.txtHVatCalcType.value)
		strVal = strVal & "&txtVatIncflag=" & Trim(.txtHVatIncFlag.value)

        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	End With    
                strVal = strVal & "&gBizArea=" & lgBACd 
                strVal = strVal & "&gPlant=" & lgPLCd 
                strVal = strVal & "&gSalesGrp=" & lgSGCd 
                strVal = strVal & "&gSalesOrg=" & lgSOCd     	
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
    DbQuery = True    

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
						<TD CLASS="TD5" NOWRAP>매출일</TD>
						<TD CLASS="TD6" NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s5112ra3_fpDateTime1_txtFromDt.js'></script>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s5112ra3_fpDateTime2_txtToDt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
						<TD CLASS=TD5>영업그룹</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSalesGrp" ALT="영업그룹" SIZE=10 MAXLENGTH=4 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup(C_PopSalesGrp)">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=23 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6"><INPUT NAME="txtItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup(C_PopItemCd)">&nbsp;<INPUT NAME="txtItemNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>발행처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBilltoParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="14XXXU" ALT="발행처">&nbsp;<INPUT NAME="txtBilltoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>VAT유형</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 ALT="VAT유형" tag="14XXXU">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=27 tag="14"></TD>
						<TD CLASS=TD5>화폐</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtCurrency" ALT="화폐" SIZE=10 MAXLENGTH=3 TAG="14XXXU"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>매출채권번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillNo" ALT="매출채권번호" TYPE="Text" MAXLENGTH="18" SIZE=30 tag="14XXXU"></TD>
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
						<script language =javascript src='./js/s5112ra3_OBJECT1_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="txtHItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHVatCalcType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHVatIncflag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHIssueDt" tag="14" alt="발행일">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
