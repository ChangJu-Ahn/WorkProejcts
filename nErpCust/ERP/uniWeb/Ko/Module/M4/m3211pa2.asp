<%@<%@ LANGUAGE="VBSCRIPT"%>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211pa2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C No POPUP ASP														*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2003/05/20																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Lee Eun Hee																		*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--TITLE><%=Request("strASPMnuMnuNm")%></TITLE-->
<TITLE>LOCAL L/C 번호</TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
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
<SCRIPT LANGUAGE="VBScript">

Option Explicit

 Const BIZ_PGM_ID = "m3211pb2.asp"			<% '☆: 비지니스 로직 ASP명 %>
	
 Const C_SHEETMAXROWS_D = 30
 Const C_MaxKey          = 7                                           '☆: key count of SpreadSheet


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
Dim arrParam
Const Bpcd	=  1											'☜: 거래처를 기준으로 Sort하기 위한 변수 
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'=======================================================================================================
Function InitVariables()
	    
	Dim strReturn
		    
	lgPageNo         = ""
	lgBlnFlgChgValue = False	                           'Indicates that no value changed
	lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	               
	gblnWinEvent = False
	Redim strReturn(0)        
	strReturn = ""
	Self.Returnvalue = strReturn   

End Function

<!--'==========================================  2.2.1 SetDefaultVal()  =====================================-->
Sub SetDefaultVal()
	frm1.txtFrDt.text = StartDate
	frm1.txtToDt.text = EndDate
	frm1.vspdData.OperationMode = 3
End Sub

<!--'==========================================  2.2.2 LoadInfTB19029()  =====================================-->
Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "PA") %>                                '☆: 
End Sub
<!--'==========================================  2.2.3 InitSpreadSheet()  =====================================-->
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3211PA1","S","A","V20030428",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
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
	Dim intColCnt
	Dim strReturn
		
	strReturn = ""
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim strReturn(frm1.vspdData.MaxCols - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		strReturn = frm1.vspdData.Text			
	End If

	Self.Returnvalue = strReturn
	Self.Close()	
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Dim strReturn
					
	Redim strReturn(0)
	strReturn = ""

	Self.Returnvalue = strReturn
	Self.Close()
End Function
		
'==========================================  2.3.3 BpcdSort()  ======================================
Sub BpcdSort(ByVal SortCol, ByVal intKey)

	frm1.vspdData.BlockMode = True
	frm1.vspdData.Col = 0
	frm1.vspdData.Col2 = frm1.vspdData.MaxCols
	frm1.vspdData.Row = 1
	frm1.vspdData.Row2 = frm1.vspdData.MaxRows
	   
	'Row기준 Sort
	frm1.vspdData.SortBy = 0
	   
	'Sort기준 Column
	frm1.vspdData.SortKey(1) = SortCol
	   
	'정렬방법 
	frm1.vspdData.SortKeyOrder(1) = intKey					'0: 정렬None 1 :오름차순  2: 내림차순 
	frm1.vspdData.Action = 25								'SS_ACTION_SORT : VB number
	   
	frm1.vspdData.BlockMode = False
    
End Sub

'========================================================================================================
' Function Name : OpenConSItemDC
'========================================================================================================
Function OpenConSItemDC()
	
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True
						
	arrParam(0) = "수혜자"							<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtBeneficiary.value)				<%' Code Condition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "수혜자"								<%' TextBox 명칭 %>
			
	arrField(0) = "BP_CD"									<%' Field명(0)%>
	arrField(1) = "BP_NM"									<%' Field명(1)%>
		    
	arrHeader(0) = "수혜자"								<%' Header명(0)%>
	arrHeader(1) = "수혜자명"							<%' Header명(1)%>
		

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        
 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		frm1.txtBeneficiary.focus 
 		Set gActiveElement = document.activeElement
 		Exit Function
 	Else
 		frm1.txtBeneficiary.value = arrRet(0) 
 		frm1.txtBeneficiaryNm.value = arrRet(1)
 		frm1.txtBeneficiary.focus 
 		Set gActiveElement = document.activeElement
 	End If	
		
 End Function

<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
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
<!--
'=========================================  3.1.2 Form_QueryUnload()  ==========================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	   
End Sub

'========================================  3.3.2 vspdData_KeyPress()  ===================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
	
<!--
'*********************************************  3.3 Object Tag 처리  ************************************
-->
	
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
<!--
'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
-->
Sub txtFrDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
	
<!--
'==========================================================================================
'   Event Name : OCX_DbClick()
'==========================================================================================
-->
Sub txtFrDt_DblClick(Button)
 If Button = 1 Then
 	frm1.txtFrDt.Action = 7
 	Call SetFocusToDocument("P")
 	frm1.txtFrDt.focus
 End If
End Sub
Sub txtToDt_DblClick(Button)
 If Button = 1 Then
 	frm1.txtToDt.Action = 7
 	Call SetFocusToDocument("P")
 	frm1.txtToDt.focus
 End If
End Sub

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
<!--
'========================================================================================
' Function Name : FncQuery
'========================================================================================
-->
Function FncQuery() 
  
	FncQuery = False                                                        '⊙: Processing is NG
	   
	Err.Clear                                                               '☜: Protect system from crashing
	
	
	With frm1
		If (UniConvDateToYYYYMMDD(.txtFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToDt.text,gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","개설일", "X")			
			.txtToDt.Focus
			Exit Function
		End if   
	End with
	  
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
	Call InitVariables 														'⊙: Initializes local global variables	
			
	If DbQuery = False Then Exit Function									

	FncQuery = True		
    
End Function

<!--
'========================================================================================
' Function Name : DbQuery
'========================================================================================
-->
Function DbQuery() 

   Err.Clear														'☜: Protect system from crashing
   DbQuery = False													'⊙: Processing is NG
	
   If LayerShowHide(1) = False Then
   	Exit Function
   End If
	
   Dim strVal
	   
   With frm1
   	If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
   		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
   		strVal = strVal & "&txtBeneficiary=" & Trim(.hdnBeneficiary.value)
   		strVal = strVal & "&txtFrDt=" & Trim(.hdnFrDt.value)
   		strVal = strVal & "&txtToDt=" & Trim(.hdnToDt.value)
       Else
   		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
   		strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)
   		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
   		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
   	End If				
				
    strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
   	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		
    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
   	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

       Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
	       
   End With
    
DbQuery = True    
 
    
End Function	
<!--
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
-->
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtMvmtType.focus
	End If
	Set gActiveElement = document.activeElement
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

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
 <TR>
 	<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
 </TR>
 <TR>
 	<TD HEIGHT=20 WIDTH=100%>
 		<FIELDSET CLASS="CLSFLD">
 			<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
 				<TR>
 					<TD CLASS=TD5 NOWRAP>수혜자</TD>
 					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수혜자" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC()">
 										 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
 					<TD CLASS=TD5 NOWRAP>개설일</TD>
 					<TD CLASS=TD6 NOWRAP>
 						<table cellspacing=0 cellpadding=0>
 							<tr>
 								<td>
 									<script language =javascript src='./js/m3211pa2_FpDateFr_txtFrDt.js'></script>
 								</td>
 								<td>~</td>
 								<td>
 									<script language =javascript src='./js/m3211pa2_FpDateTo_txtToDt.js'></script>
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
 					<script language =javascript src='./js/m3211pa2_vaSpread1_vspdData.js'></script>
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
 				<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="SearchBtn" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
 									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>
 				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="OkBtn"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
 						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="CancelBtn"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
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
<INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<!--INPUT TYPE=HIDDEN NAME="hdnMode" tag="14"-->
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</FROM>
</BODY>
</HTML>






	
	

