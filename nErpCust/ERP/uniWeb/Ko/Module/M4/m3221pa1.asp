<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 구매 
'*  2. Function Name        : L/C관리 
'*  3. Program ID           : m3221pa1
'*  4. Program Name         : L/C AMEND 관리번호 
'*  5. Program Desc         : L/C Amend 등록을 위한 L/C AMEND 관리번호 
'*  6. Comproxy List        : M32218ListLcAmendHdrSvr
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/26
'*  9. Modifier (First)     : Byun Jee Hyun	
'* 10. Modifier (Last)      : Kang Su-hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>L/C AMEND 관리번호</TITLE>
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
<SCRIPT LANGUAGE="VBScript">

Option Explicit                        

'########################################################################################################
'#									1.  Data Declaration Part
'########################################################################################################

Const BIZ_PGM_ID 		= "m3221pb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 1                                           '☆: key count of SpreadSheet

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
		Dim arrParent
		'Dim arrParam
		
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
				
        gblnWinEvent = False
        Redim arrReturn(0)        
        arrReturn = ""
        Self.Returnvalue = arrReturn     

	End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
	Sub SetDefaultVal()
		frm1.txtFrDt.Text = StartDate
		frm1.txtToDt.Text = EndDate
		frm1.vspdData.OperationMode = 3
	End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

		<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "PA") %>                                '☆: 
	End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
	Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("M3221PA1","S","A","V20030402",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
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
		
		DIM strReturn
		Dim intColCnt
		
		If frm1.vspdData.ActiveRow > 0 Then	
			Redim strReturn(0)	
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col =  GetKeyPos("A",1)
			
			strReturn = frm1.vspdData.Text
		End If
		
		Self.Returnvalue = strReturn
		Self.Close()
	
	End Function

'=========================================  2.3.2 CancelClick()  ========================================
	Function CancelClick()
		DIM strReturn
		
		Redim strReturn(0)
		strReturn = ""
		Self.Returnvalue = strReturn
		Self.Close()
	End Function

'========================================================================================================
' Function Name : OpenConSItemDC
'========================================================================================================
Function OpenConSItemDC()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	arrParam(0) = "수출자"										<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtBeneficiary.value)				<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtBeneficiaryNm.value)				<%' Name Cindition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
	arrParam(5) = "수출자"										<%' TextBox 명칭 %>
		
    arrField(0) = "BP_CD"										<%' Field명(0)%>
    arrField(1) = "BP_NM"										<%' Field명(1)%>
	    
    arrHeader(0) = "수출자"										<%' Header명(0)%>
    arrHeader(1) = "수출자명"									<%' Header명(1)%>
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	arrParam(0) = arrParam(5)												' 팝업 명칭	

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

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtFrDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToDt_KeyPress(KeyAscii)
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
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToDt.text,gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","신청일", "X")			
			.txtToDt.Focus
			Exit Function
		End if   
	End with
   

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables
	
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
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						'☜: 비지니스 처리 ASP의 상태	
																				'☆: 조회 조건 데이타 
		    strVal = strVal & "&txtBeneficiary=" & Trim(.hdnBeneficiary.value)	'수출자 
		    strVal = strVal & "&txtFrDt=" & Trim(.hdnFrDt.Value)				'신청일 
		    strVal = strVal & "&txtToDt=" & Trim(.hdnToDt.Value)		
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)	'수출자 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)					'신청일 
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
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
						<TD CLASS=TD5>수출자</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수출자" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC()">
									  <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD6 NOWRAP>
						<TD CLASS=TD6 NOWRAP>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>AMEND 신청일</TD>
						<TD CLASS=TD6 NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3221pa1_fpDateTime1_txtFrDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3221pa1_fpDateTime2_txtToDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS=TD6 NOWRAP>
						<TD CLASS=TD6 NOWRAP>
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
						<script language =javascript src='./js/m3221pa1_vaSpread1_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
