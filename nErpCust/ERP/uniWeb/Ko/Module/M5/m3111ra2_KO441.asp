<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 구매 
'*  2. Function Name        : B/L관리 
'*  3. Program ID           : m3111ra2.asp
'*  4. Program Name         : 발주참조 
'*  5. Program Desc         : B/L 등록을 위한 발주참조 
'*  6. Comproxy List        : M31118ListPoHdrForBlSvr
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : 	
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
<TITLE>발주참조</TITLE>
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "m3111rb2_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 1                                           '☆: key count of SpreadSheet

Const gstrpaymethMajor 	= "B9004"										'결제방법 
Const gstrIncotermsMajor= "B9006"										'가격조건 


<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgPopUpR                                                '☜: Orderby default 값                    
Dim IscookieSplit 
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
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
    
    gblnWinEvent = False
    Self.Returnvalue = ""     
End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
	frm1.txtPOFrDt.text = StartDate
	frm1.txtPOToDt.text = EndDate
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrp, "Q") 
		frm1.txtPurGrp.Tag = left(frm1.txtPurGrp.Tag,1) & "4" & mid(frm1.txtPurGrp.Tag,3,len(frm1.txtPurGrp.Tag))
        frm1.txtPurGrp.value = lgPGCd
	End If
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","RA") %>                                '☆: 
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("M3111RA2","S","A","V20030402",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
		Call SetSpreadLock 
		frm1.vspdData.OperationMode = 3  
    
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	

'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()

	Dim intColCnt
	Dim rtnValue
	
	If frm1.vspdData.ActiveRow > 0 Then		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		rtnValue = frm1.vspdData.Text		
	End If
	
	Self.Returnvalue = rtnValue
	Self.Close()

End Function
'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn(0)
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
	Case 0	<%'OpenBizPartner%>
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

	Case 1	<%'OpenPurGrp%>
	    If frm1.txtPurGrp.className = "protected" Then Exit Function
		arrParam(0) = "구매그룹"							<%' 팝업 명칭 %>
		arrParam(1) = "B_PUR_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtPurGrp.value)					<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "구매그룹"							<%' TextBox 명칭 %>
	
		arrField(0) = "PUR_GRP"								<%' Field명(0)%>
		arrField(1) = "PUR_GRP_NM"							<%' Field명(1)%>
	
		arrHeader(0) = "구매그룹"						<%' Header명(0)%>
		arrHeader(1) = "구매그룹명"						<%' Header명(1)%>
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	Case 2	<%'OpenPOType%>
		arrParam(0) = "발주형태"									<%' 팝업 명칭 %>
		arrParam(1) = "m_config_process"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtPOType.Value)							<%' Code Condition%>
		arrParam(4) = "Import_flg=" & FilterVar("Y", "''", "S") & " " 								<%' Where Condition%>
		arrParam(5) = "발주형태"									<%' TextBox 명칭 %>
	
		arrField(0) = "po_type_cd"									<%' Field명(0)%>
		arrField(1) = "po_type_Nm"									<%' Field명(1)%>
	
		arrHeader(0) = "발주형태"								<%' Header명(0)%>
		arrHeader(1) = "발주형태명"								<%' Header명(1)%>
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 3	<%'OpenPayMeth%>
		arrParam(0) = "결제방법"								<%' 팝업 명칭 %>
		arrParam(1) = "b_minor,b_configuration"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtpaymeth.Value)						<%' Code Condition%>
'		arrParam(3) = Trim(.txtpaymethNm.Value)						<%' Name Cindition%>
		arrParam(4) = "b_minor.Major_Cd= " & FilterVar(gstrpaymethMajor, "''", "S") & " and  b_minor.major_cd=b_configuration.major_cd and b_minor.minor_cd=b_configuration.minor_cd AND b_configuration.REFERENCE <> " & FilterVar("M", "''", "S") & "  AND b_configuration.seq_no = 1 "
	'	arrParam(4) = "MAJOR_CD='" & gstrpaymethMajor & "'" & " AND REFERENCE = 'M'"		<%' Where Condition%>
		arrParam(5) = "결제방법"								<%' TextBox 명칭 %>

		arrField(0) = "b_minor.Minor_CD"									<%' Field명(0)%>
		arrField(1) = "b_minor.Minor_NM"									<%' Field명(1)%>

		arrHeader(0) = "결제방법"								<%' Header명(0)%>
		arrHeader(1) = "결제방법명"								<%' Header명(1)%>
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 4	<%'OpenIncoterms%>
		arrParam(0) = "가격조건"									<%' 팝업 명칭 %>
		arrParam(1) = "B_Minor"										<%' TABLE 명칭 %>
		arrParam(2) = Trim(.txtIncoterms.Value)						<%' Code Condition%>
		arrParam(4) = "MAJOR_CD= " & FilterVar(gstrIncotermsMajor, "''", "S") & "" 		<%' Where Condition%>
		arrParam(5) = "가격조건"									<%' TextBox 명칭 %>
	
		arrField(0) = "Minor_CD"									<%' Field명(0)%>
		arrField(1) = "Minor_NM"									<%' Field명(1)%>
	
		arrHeader(0) = "가격조건"								<%' Header명(0)%>
		arrHeader(1) = "가격조건명"								<%' Header명(1)%>
	
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
'	Description : OpenConSItemDC Popup에서 Return되는 값 setting
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
		Select Case iWhere
		Case 0		<%'SetBizPartner%>
			frm1.txtBeneficiary.value 	= arrRet(0)
			frm1.txtBeneficiaryNm.value = arrRet(1)
			frm1.txtBeneficiary.focus
		Case 1		<%'SetPurGrp%>
			frm1.txtPurGrp.value 	= arrRet(0)
			frm1.txtPurGrpNm.value 	= arrRet(1)
			frm1.txtPurGrp.focus
		Case 2		<%'SetPOType%>
			frm1.txtPOType.Value 	= arrRet(0)
			frm1.txtPOTypeNm.Value 	= arrRet(1)
			frm1.txtPOType.focus
		Case 3		<%'SetPayMeth%>
			frm1.txtpaymeth.Value 	= arrRet(0)
			frm1.txtpaymethNm.Value = arrRet(1)
			frm1.txtpaymeth.focus
		Case 4		<%'SetIncoterms%>
			frm1.txtIncoterms.Value 	= arrRet(0)
			frm1.txtIncotermsNm.Value 	= arrRet(1)
			frm1.txtIncoterms.focus
		End Select
		Set gActiveElement = document.activeElement
End Function

'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	
	Call ggoOper.LockField(Document, "N")                                              '⊙: Lock  Suitable  Field
	Call InitVariables											  '⊙: Initializes local global variables
    Call GetValue_ko441()
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
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
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
Sub txtPOFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPOFrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtPOFrDt.focus
	End if
End Sub

Sub txtPOToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPOToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtPOToDt.focus
	End if
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtPOFrDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtPOToDt_KeyPress(KeyAscii)
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
	
	If ValidDateCheck(frm1.txtPOFrDt, frm1.txtPOToDt) = False Then Exit Function
   
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
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
																							<%'☜: 조회 조건 데이타 %>
		    strVal = strVal & "&txtBeneficiary=" 	& Trim(.hdnBeneficiary.value)			<%'수출자-Biz.Partner(Hidden)%>
			strVal = strVal & "&txtPurGrp=" 		& Trim(.hdnPurGrp.value)				<%'구매그룹(Hidden)%>
			strVal = strVal & "&txtPOType=" 		& Trim(.hdnPOType.value)				<%'발주형태(Hidden)%>
			strVal = strVal & "&txtpaymeth=" 		& Trim(.hdnPayMeth.value)				<%'결제방법(Hidden)%>
			strVal = strVal & "&txtIncoterms=" 		& Trim(.hdnIncoterms.value)				<%'가격조건(Hidden)%>
		    strVal = strVal & "&txtPOFrDt=" 		& Trim(.hdnFrDt.Value)					<%'발주일(시작)(Hidden)%>
		    strVal = strVal & "&txtPOToDt=" 		& Trim(.hdnToDt.Value)					<%'발주일(종료)(Hidden)%>
			strVal = strVal & "&lgStrPrevKey=" 		& lgStrPrevKey							<%'☜: Next key tag%>
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>
																							<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtBeneficiary=" 	& Trim(.txtBeneficiary.value)			<%'수출자-Biz.Partner%>
			strVal = strVal & "&txtPurGrp=" 		& Trim(.txtPurGrp.value)				<%'구매그룹 %>
			strVal = strVal & "&txtPOType=" 		& Trim(.txtPOType.value)				<%'발주형태 %>
			strVal = strVal & "&txtpaymeth=" 		& Trim(.txtpaymeth.value)				<%'결제방법 %>
			strVal = strVal & "&txtIncoterms=" 		& Trim(.txtIncoterms.value)				<%'가격조건 %>
			strVal = strVal & "&txtPOFrDt=" 		& Trim(.txtPOFrDt.text)					<%'발주일(시작)%>
			strVal = strVal & "&txtPOToDt=" 		& Trim(.txtPOToDt.text)					<%'발주일(종료)%>
			strVal = strVal & "&lgStrPrevKey=" 		& lgStrPrevKey							<%'☜: Next key tag%>
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

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
		frm1.txtDnType.focus
	End If

End Function

'========================================================================================================
' Function Name : OpenOrderByPopup
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
				<TD CLASS=TD5>수출자</TD>
				<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="수출자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 0" >&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
				<TD CLASS=TD5>구매그룹</TD>
				<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="11XXXU" ALT="수입담당"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="14"></TD>
			</TR>
			<TR>
				<TD CLASS=TD5>발주형태</TD>
				<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPOType" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPOType" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtPOTypeNm" SIZE=20 TAG="14"></TD>
				<TD CLASS=TD5>발주일</TD>
				<TD CLASS=TD6 NOWRAP>
					<table cellspacing=0 cellpadding=0>
						<tr>
							<td NOWRAP>
								<script language =javascript src='./js/m3111ra2_fpDateTime1_txtPOFrDt.js'></script>
							</td>
							<td NOWRAP>
								~
							</td>
							<td NOWRAP>
								<script language =javascript src='./js/m3111ra2_fpDateTime2_txtPOToDt.js'></script></TD>
							</td>
						</tr>
					</table>
				</TD>		 	
			</TR>
			<TR>
				<TD CLASS=TD5>결제방법</TD>
				<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPayMeth" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnpaymeth" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 3">
							  <INPUT TYPE=TEXT NAME="txtpaymethNm" SIZE=20 TAG="14"></TD>
				<TD CLASS=TD5>가격조건</TD>
				<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="가격조건"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 4">
							  <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="14"></TD>

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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/m3111ra2_vspdData_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPurGrp" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPOType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayMeth" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIncoterms" tag="14">
</FORM>

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
