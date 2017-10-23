<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 구매 
'*  2. Function Name        : L/C관리 
'*  3. Program ID           : m3211ra2
'*  4. Program Name         : Local L/C참조 
'*  5. Program Desc         : Local L/C Amend등록을 위한 Local L/C참조 
'*  6. Comproxy List        : M32118ListLcHdrForAmendSvr
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : 	
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>LOCAL L/C참조</TITLE>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#									1.  Data Declaration Part
'########################################################################################################

Const BIZ_PGM_ID 		= "m3211rb2.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 1                                           '☆: key count of SpreadSheet
Const gstrPayTermsMajor = "B9004"


<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgSelectList                                            '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT                                          '☜: SpreadSheet의 초기  위치정보관련 변수 

Dim lgSortFieldNm                                           '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD                                           '☜: Orderby popup용 데이타(필드코드)      

Dim lgPopUpR                                                '☜: Orderby default 값                    

Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
   													    'Window가 여러 개 뜨는 것을 방지하기 위해 
   													    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
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

'========================================== 2.1.1 InitVariables()  ======================================
Function InitVariables()
	
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
				
    gblnWinEvent = False
   
    Self.Returnvalue = ""     

End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
						
	frm1.txtOpenFrDt.text = StartDate
	frm1.txtOpenToDt.text = EndDate
	frm1.vspdData.OperationMode = 3
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "RA") %>                                '☆: 

End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3211RA2","S","A","V20030402",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock 	    
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	
'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()

	Dim strReturn
		
	With frm1.vspdData 
		If .ActiveRow > 0 Then	
			Redim strReturn(.MaxCols - 1)
			
			.Row = .ActiveRow
			.Col =  GetKeyPos("A",1)
			strReturn = Trim(.Text)
					
			Self.Returnvalue = strReturn
		End If
	End With
			
	Self.Close()
			
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenConSItemDC
'========================================================================================================
Function OpenConSItemDC(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
		
	With frm1
		
	Select Case iWhere
	Case 0
		arrParam(0) = "수출자"										<%' 팝업 명칭 %>
		arrParam(1) = "B_BIZ_PARTNER"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtBeneficiary.value)					<%' Code Condition%>
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
		arrParam(5) = "수출자"										<%' TextBox 명칭 %>
			
	    arrField(0) = "BP_CD"										<%' Field명(0)%>
	    arrField(1) = "BP_NM"										<%' Field명(1)%>
		    
	    arrHeader(0) = "수출자"										<%' Header명(0)%>
	    arrHeader(1) = "수출자명"									<%' Header명(1)%>
		    
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 1
		arrParam(0) = "구매그룹"						<%' 팝업 명칭 %>
		arrParam(1) = "B_PUR_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtPurGrp.value)				<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "구매그룹"						<%' TextBox 명칭 %>
		
		arrField(0) = "PUR_GRP"								<%' Field명(0)%>
		arrField(1) = "PUR_GRP_NM"							<%' Field명(1)%>
		
		arrHeader(0) = "구매그룹"						<%' Header명(0)%>
		arrHeader(1) = "구매그룹명"						<%' Header명(1)%>
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	    		
	Case 2
		arrParam(0) = "결제방법"									<%' 팝업 명칭 %>
		arrParam(1) = "b_minor,b_configuration"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtPayTerms.Value)						<%' Code Condition%>
		arrParam(4) = "b_minor.Major_Cd= " & FilterVar(gstrPayTermsMajor, "''", "S") & " and  b_minor.major_cd=b_configuration.major_cd and b_minor.minor_cd=b_configuration.minor_cd AND b_configuration.REFERENCE <> " & FilterVar("M", "''", "S") & "  AND b_configuration.seq_no = 1 "
		arrParam(5) = "결제방법"									<%' TextBox 명칭 %>

		arrField(0) = "b_minor.Minor_CD"							<%' Field명(0)%>
		arrField(1) = "b_minor.Minor_NM"							<%' Field명(1)%>

		arrHeader(0) = "결제방법"								<%' Header명(0)%>
		arrHeader(1) = "결제방법명"								<%' Header명(1)%>
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	End Select
		
	End With
		
	arrParam(0) = arrParam(5)												' 팝업 명칭	

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
		
End Function

'-------------------------------------------------------------------------------------------------------
'	Name : SetConSItemDC()
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
With frm1
	Select Case iWhere
	Case 0
		.txtBeneficiary.value = arrRet(0)
		.txtBeneficiaryNm.value = arrRet(1)
		.txtBeneficiary.focus
	Case 1
		.txtPurGrp.value = arrRet(0)
		.txtPurGrpNm.value = arrRet(1)
		.txtPurGrp.focus
	Case 2
		.txtPayTerms.Value = arrRet(0)
		.txtPayTermsNm.Value = arrRet(1)
		.txtPayTerms.focus
	End Select
	Set gActiveElement = document.activeElement
End With
End Function
'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
	Call LoadInfTB19029													'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)		
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
	Call InitVariables											  '⊙: Initializes local global variables
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

'========================================================================================================
'   Event Name : OCX_DbClick()
'========================================================================================================
Sub txtOpenFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenFrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtOpenFrDt.focus		
	End If
End Sub

Sub txtOpenToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtOpenToDt.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtOpenFrDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
	   Call CancelClick()
	Elseif KeyAscii = 13 Then
	   Call FncQuery()
	End if
End Sub

Sub txtOpenToDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
	   Call CancelClick()
	Elseif KeyAscii = 13 Then
	   Call FncQuery()
	End if
End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 
    
	FncQuery = False                                                        '⊙: Processing is NG
	    
	Err.Clear                                                               '☜: Protect system from crashing
	With frm1
		if (UniConvDateToYYYYMMDD(.txtOpenFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtOpenToDt.text,gDateFormat,"")) and Trim(.txtOpenFrDt.text)<>"" and Trim(.txtOpenToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","개설일", "X")			
			.txtOpenToDt.Focus
			Exit Function
		End if   
	End with
	   
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData	
    Call InitVariables 

	If DbQuery = False Then Exit Function									

	FncQuery = True		
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : DbQuery
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
			 																					<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtBeneficiary=" 	& Trim(.txtHBeneficiary.value)				<%'수혜자 %>
			strVal = strVal & "&txtPurGrp=" 		& Trim(.txtHPurGrp.value)					<%'구매그룹 %>
			strVal = strVal & "&txtPayTerms=" 		& Trim(.txtHPayTerms.value)					<%'결제방법 %>
			strVal = strVal & "&txtOpenFrDt=" 		& Trim(.txtHOpenFrDt.value)					<%'개설일 %>
			strVal = strVal & "&txtOpenToDt=" 		& Trim(.txtHOpenToDt.value)					
			strVal = strVal & "&lgStrPrevKey="   	& lgStrPrevKey     
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			 																					<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtBeneficiary=" 	& Trim(.txtBeneficiary.value)				<%'수혜자 %>
			strVal = strVal & "&txtPurGrp=" 		& Trim(.txtPurGrp.value)					<%'구매그룹 %>
			strVal = strVal & "&txtPayTerms=" 		& Trim(.txtPayTerms.value)					<%'결제방법 %>
			strVal = strVal & "&txtOpenFrDt=" 		& Trim(.txtOpenFrDt.text)					<%'개설일 %>
			strVal = strVal & "&txtOpenToDt=" 		& Trim(.txtOpenToDt.text)					
			strVal = strVal & "&lgStrPrevKey="   	& lgStrPrevKey
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
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBeneficiary.focus
	End If
	Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : OpenOrderBy
'========================================================================================================
Function OpenOrderByPopup()
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

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
   					<TD CLASS=TD5>수혜자</TD>
   					<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="수혜자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
   					<TD CLASS=TD5 NOWRAP>구매그룹</TD>
   					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
   					</TR>
   				<TR>
   					<TD CLASS=TD5>결제방법</TD>
   					<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
   					<TD CLASS=TD5 NOWRAP>개설일</TD>						
   					<TD CLASS=TD6 NOWRAP>
   						<table cellspacing=0 cellpadding=0>
   							<tr>
   								<td>
   									<script language =javascript src='./js/m3211ra2_fpDateTime1_txtOpenFrDt.js'></script>
   								</td>
   								<td>~</td>
   								<td>
   									<script language =javascript src='./js/m3211ra2_fpDateTime2_txtOpenToDt.js'></script>
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
   					<script language =javascript src='./js/m3211ra2_vaSpread1_vspdData.js'></script>
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
   				<TD WIDTH=70% NOWRAP> <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
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
<INPUT TYPE=HIDDEN                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
