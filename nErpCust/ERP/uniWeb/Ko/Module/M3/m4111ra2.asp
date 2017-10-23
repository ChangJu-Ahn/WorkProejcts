<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m4111ra2
'*  4. Program Name         : 구매입고참조 
'*  5. Program Desc         : 구매입고참조 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBS">
Option Explicit					 '☜: indicates that All variables must be declared in advance 

'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID 		= "m4111rb2.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 27                                           '☆: key count of SpreadSheet

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 														    'Window가 여러 개 뜨는 것을 방지하기 위해 
Dim iDBSYSDate
Dim EndDate, StartDate
Dim lblnWinEvent
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName


iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

Const BIZ_PGM_QRY_ID = "m4111rb2.asp"			 '☆: 비지니스 로직 ASP명 

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	Dim arrParam
			
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
	lgIntGrpCount = 0										'⊙: Initializes Group View Size
	lgBlnFlgChgValue = False	                           'Indicates that no value changed
	lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgSortKey        = 1   
	        
	arrParam = arrParent(1)
			
	frm1.hdnMvmtType.value  	= arrParam(0)
	frm1.hdnSupplierCd.value 	= arrParam(1)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.hdnRefType.value 		= arrParam(3)
	frm1.hdnIvType.value 		= arrParam(4)
	frm1.hdnRefPONO.value 		= arrParam(5)
	frm1.hdnCurrency.value		= arrParam(6)
	'반품내역등록 - 외주가공여부 조건 추가 
	If UBound(arrparam) > 7 Then
		frm1.hdnSubcontraflg.value	= arrParam(7)
	End if
	frm1.vspdData.OperationMode = 5

	gblnWinEvent = False
	Redim arrReturn(0,0)        
	Self.Returnvalue = arrReturn     
End Function

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "RA") %>                                '☆: 
End Sub
	

'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M4111RA2","S","A","V20051201",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock 
End Sub
	

'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
	If .vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		'Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-2)
		Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-1)

		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt

			If frm1.vspdData.SelModeSelected Then
				'For intColCnt = 0 To frm1.vspdData.MaxCols - 2
				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
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
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
	
 
'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
Function OpenMvmtNo()
	
	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Or UCase(frm1.txtMvmtNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
	
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnMvmtType.Value)
	arrParam(1) = Trim(frm1.hdnSupplierCd.Value)
	arrParam(2) = Trim(frm1.hdnGroupCd.Value)
	arrParam(3) = ""
	
	iCalledAspName = AskPRAspName("M4111PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M4111PA3", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtMvmtNo.focus
		Exit Function
	Else
		frm1.txtMvmtNo.value = strRet(0)
		frm1.txtMvmtNo.focus
	End If	
End Function

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
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


'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'==========================================  2.2.1 SetDefaultVal()  ====================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)	
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtFrMvmtDt.text = StartDate
	frm1.txtToMvmtDt.text = EndDate
End Sub



'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
Sub txtFrMvmtDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToMvmtDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'==========================================================================================
'   Event Name : txtFrMvmtDt
'   Event Desc :
'==========================================================================================
Sub txtFrMvmtDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrMvmtDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtFrMvmtDt.Focus
	End if
End Sub


'==========================================================================================
'   Event Name : txtToMvmtDt
'   Event Desc :
'==========================================================================================
Sub txtToMvmtDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToMvmtDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtToMvmtDt.Focus
	End if
End Sub

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	
     If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
     End If

	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Function

Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
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
	
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFrMvmtDt, frm1.txtToMvmtDt) = False Then Exit Function
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
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
	If DbQuery = False Then Exit Function									

    FncQuery = True		
End Function

'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal

	
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
    		strVal = strVal & "&txtMvmtNo=" & .hdnMvmtNo.value
			strVal = strVal & "&txtFrMvmtDt=" & .hdnFrMvmtDt.value
			strVal = strVal & "&txtToMvmtDt=" & .hdnToMvmtDt.value
			strVal = strVal & "&txtRefPONO=" & .hdnRefPONO.value
			strval = strval & "&txtcur=" & .hdncurrency.value
			strVal = strVal & "&txtRefType=" & .hdnRefType.value
			strVal = strVal & "&txtIvType=" & .hdnIvType.value
			strVal = strVal & "&txtSppl=" & .hdnSupplierCd.value
			strVal = strVal & "&txtSubcontraflg=" & .hdnSubcontraflg.value				' 외주가공여부 추가 
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.Value)
			strVal = strVal & "&txtFrMvmtDt=" & Trim(.txtFrMvmtDt.text)
			strVal = strVal & "&txtToMvmtDt=" & Trim(.txtToMvmtDt.text)
			strVal = strVal & "&txtRefPONO=" & .hdnRefPONO.value
			strval = strval & "&txtcur=" & .hdncurrency.value
			strVal = strVal & "&txtRefType=" & .hdnRefType.value
			strVal = strVal & "&txtIvType=" & .hdnIvType.value
			strVal = strVal & "&txtSppl=" & .hdnSupplierCd.value
			strVal = strVal & "&txtSubcontraflg=" & .hdnSubcontraflg.value				' 외주가공여부 추가 
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function


'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtMvmtNo.focus
	End If
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
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
						<TD CLASS="TD5" NOWRAP>입고번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmt" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()">
											   <div style="Display:none"><input type=text name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>입고일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m4111ra2_fpDateTime1_txtFrMvmtDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m4111ra2_fpDateTime1_txtToMvmtDt.js'></script>
									</td>
								<tr>
							</table>
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
						<script language =javascript src='./js/m4111ra2_vaSpread1_vspdData.js'></script>
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
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrMvmtDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToMvmtDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefPONO" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnCurrency" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
