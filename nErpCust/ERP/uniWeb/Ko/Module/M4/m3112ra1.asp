<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : M3112RA1																	*
'*  4. Program Name         : P/O Detail Reference ASP																*
'*  5. Program Desc         : L/C Amend 내역등록을 위한 발주 내역 참조 *
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/07/11																*
'*  9. Modifier (First)     : 																*
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2002/05/06 : ADO Conv.													*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>발주내역참조</TITLE>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">

Option Explicit		


Const BIZ_PGM_ID 		= "m3112rb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 15                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgSelectList                   
Dim lgSelectListDT                 

Dim lgSortFieldNm                  
Dim lgSortFieldCD                  

Dim lgPopUpR                       

Dim lgKeyPos                       
Dim lgKeyPosVal                    
Dim lgCookValue 

Dim IsOpenPop  
Dim gblnWinEvent

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
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

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	Dim arrParam
		
	arrParam = arrParent(1)
	
	With frm1
		.hdnPoNo.value 				= arrParam(0)
		'.txtPayMethCd.value 		= arrParam(1)
		'.txtPayMethNm.value 		= arrParam(2)
		.txtIncotermsCd.value 		= arrParam(3)
		.txtIncotermsNm.value 		= arrParam(4)
		.txtCurrency.value 			= arrParam(5)
		.txtBeneficiaryCd.value 	= arrParam(6)
		.txtBeneficiaryNm.value 	= arrParam(7)
		.txtGrpCd.value 			= arrParam(8)
		.txtGrpNm.value 			= arrParam(9)
			
		.vspdData.OperationMode = 5
	End With
		
End Sub

'=======================  LoadInfTB19029()  ==============================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>

End Sub
'=======================  InitSpreadSheet()  ==============================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3112RA1","S","A","V20030711",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock 	    
End Sub
'=======================  SetSpreadLock()  ==============================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
		If .vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(.vspdData.SelModeSelCount-1, .vspdData.MaxCols-1)

			For intRowCnt = 0 To .vspdData.MaxRows - 1

				.vspdData.Row = intRowCnt + 1

				If .vspdData.SelModeSelected Then
					For intColCnt = 0 To .vspdData.MaxCols - 1
						'.vspdData.Col = intColCnt + 1
						'arrReturn(intInsRow, intColCnt) = TRIM(.vspdData.Text)
						.vspdData.Col = GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = Trim(.vspdData.Text)
					Next

					intInsRow = intInsRow + 1

				End IF
			Next
		End if			
	end with
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	
	If gblnWinEvent = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	gblnWinEvent = True

	arrParam(0) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "30"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명							' Header명(1)

	iCalledAspName = AskPRAspName("B1B01PA2")

	If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "B1B01PA2", "X")
			IsOpenPop = False
			Exit Function
	End If
			
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrparam,arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		frm1.txtTrackingNo.focus
		lgBlnFlgChgValue = True
		Set gActiveElement = document.activeElement
	End If	

End Function
'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()

	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
   
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>

	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
		
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
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

'========================================  3.3.2 vspdData_KeyPress()  ===================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
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
'===================================   FncQuery()   ================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    						
	Call InitVariables												

    'frm1.vspdData.Maxrows = 0

    If DbQuery = False Then Exit Function							

    FncQuery = True	
    Set gActiveElement = document.activeElement								
        
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
		
	With frm1

	   If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
			strVal = strVal & "&txtGrpCd=" & Trim(.hdnGrpCd.value)
			strVal = strVal & "&txtBeneficiaryCd=" & Trim(.hdnBeneficiaryCd.value)
			strVal = strVal & "&txtCurrency=" & Trim(.hdnCurrency.value)
			'strVal = strVal & "&txtPayMethCd=" & Trim(.hdnPayMethCd.value)
			strVal = strVal & "&txtIncotermsCd=" & Trim(.hdnIncotermsCd.value)
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtItemCd="&Trim(.txtItemCd.value)
			strVal = strVal & "&txtGrpCd="&Trim(.txtGrpCd.value)
			strVal = strVal & "&txtBeneficiaryCd="&Trim(.txtBeneficiaryCd.value)
			strVal = strVal & "&txtCurrency="&Trim(.txtCurrency.value)
			'strVal = strVal & "&txtPayMethCd="&Trim(.txtPayMethCd.value)
			strVal = strVal & "&txtIncotermsCd="&Trim(.txtIncotermsCd.value)
		End If
	
	
		strVal = strVal & "&txtLcFlg="&Trim(.hdnLcFlg.Value)
		strVal = strVal & "&txtPoNo="&Trim(.hdnPoNo.Value)
		strVal = strVal & "&txtTrackingNo=" &Trim(.txtTrackingNo.value)
	
	End With	
		strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbQuery = True														<%'⊙: Processing is NG%>
End Function

'==========================  DbQueryOk()  =====================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	with frm1
		If .vspdData.MaxRows > 0 Then
			.vspdData.Focus
			.vspdData.Row = 1	
			.vspdData.SelModeSelected = True		
		Else
			.txtItemCd.focus
			Set gActiveElement = document.activeElement
		End If
	end with
End Function

'========================================================================================================
' Function Name : OpenOrderBy
'========================================================================================================
Function OpenOrderByPopup()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
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
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
					<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
						<TR>
							<TD CLASS="TD5" NOWRAP>품목</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목"  NAME="txtItemCd" MAXLENGTH=18 SIZE=10 tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
											 	   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
							<TD CLASS="TD5" NOWRAP>구매그룹</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtGrpCd" SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
											 	   <INPUT TYPE=TEXT ALT="구매그룹" NAME="txtGrpNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수출자</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수출자"  NAME="txtBeneficiaryCd" MAXLENGTH=18 SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
											 	   <INPUT TYPE=TEXT ALT="수출자" NAME="txtBeneficiaryNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
							<TD CLASS="TD5" NOWRAP>화폐</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="화폐"  NAME="txtCurrency" SIZE=10 tag="14NXXU" ></TD>
						</TR>
						<TR>							
							<TD CLASS="TD5" NOWRAP>가격조건</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="가격조건"  NAME="txtIncotermsCd" SIZE=10 tag="14NXXU" >&nbsp;&nbsp;&nbsp;&nbsp;
											 	   <INPUT TYPE=TEXT ALT="가격조건" NAME="txtIncotermsNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
							<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=Text ALT="Tracking번호" NAME="txtTrackingNo"   MAXLENGTH=25 SiZE=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>									 	   
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
						<script language =javascript src='./js/m3112ra1_A_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnLcFlg" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnBeneficiaryCd" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnCurrency" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" TAG="14">
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</Form>
</BODY>
</HTML>

