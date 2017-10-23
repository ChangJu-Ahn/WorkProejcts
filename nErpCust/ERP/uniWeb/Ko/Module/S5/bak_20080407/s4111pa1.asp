<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111PA1
'*  4. Program Name         : 출하관리번호 팝업 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2003/06/12
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Date표준적용 
'*                            2002/12/13 Include 성능향상 강준구 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>출하번호</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop

Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

Const BIZ_PGM_ID        = "s4111pb1.asp"
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value
 
'=========================================
Sub InitVariables()
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          
	lgIsOpenPop = False
End Sub

'=========================================
Sub SetDefaultVal()

	Dim arrParam

	frm1.txtDNFrDt.text = StartDate
	frm1.txtDNToDt.text = EndDate
	
	<%If Request("txtExceptFlag") = "N" Then%>
		frm1.txtDlvyFrDt.text = StartDate
		frm1.txtDlvyToDt.text = EndDate
	<%End If%>

	frm1.txtHExceptFlag.value = "<%=Request("txtExceptFlag")%>"
	Self.Returnvalue = ""

End Sub

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
End Sub

'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4111pa1","S","A","V20060320", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "납품처"							
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.txtBpCd.value)				
	arrParam(3) = ""									
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
	arrParam(5) = "납품처"							
		
	arrField(0) = "BP_CD"								
	arrField(1) = "BP_NM"								
		
	arrHeader(0) = "납품처"							
	arrHeader(1) = "납품처명"						
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False

	frm1.txtBpCd.focus
			
	If arrRet(0) <> "" Then
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
	End If
End Function

'=========================================
Function OpenMinorCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"				
	arrParam(2) = Trim(frm1.txtMovType.value)
	arrParam(3) = ""
	arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
	arrParam(5) = "출하형태"

	arrField(0) = "A.MINOR_CD"
	arrField(1) = "A.MINOR_NM"

	arrParam(0) = arrParam(5)
		
	arrHeader(0) = "출하형태"						
	arrHeader(1) = "출하형태명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtMovType.focus
	
	If arrRet(0) <> "" Then
		frm1.txtMovType.value = arrRet(0)
		frm1.txtMovTypeNm.value = arrRet(1)
	End If
End Function

'=========================================
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "영업그룹"								
	arrParam(1) = "B_SALES_GRP"									
	arrParam(2) = Trim(frm1.txtSalesGroup.value)						
	arrParam(3) = ""											
	arrParam(4) = ""											
	arrParam(5) = "영업그룹"								

	arrField(0) = "SALES_GRP"									
	arrField(1) = "SALES_GRP_NM"										

	arrHeader(0) = "영업그룹"								
	arrHeader(1) = "영업그룹명"								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtSalesGroup.focus
	
	If arrRet(0) <> "" Then
		frm1.txtSalesGroup.Value = arrRet(0)
		frm1.txtSalesGroupNm.Value = arrRet(1)
	End If
End Function

'=========================================
Function OpenSONo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주번호"											
	arrParam(1) = "S_SO_HDR A, B_BIZ_PARTNER B, B_SALES_GRP C"				
	arrParam(2) = Trim(frm1.txtSONo.value)									
	arrParam(3) = ""														
	arrParam(4) = "A.SOLD_TO_PARTY = B.BP_CD AND A.SALES_GRP = C.SALES_GRP" 
	arrParam(5) = "수주번호"											

	arrField(0) = "ED12" & PopupParent.gColSep & "A.SO_NO"								
	arrField(1) = "ED10" & PopupParent.gColSep & "A.SOLD_TO_PARTY"						
	arrField(2) = "ED15" & PopupParent.gColSep & "B.BP_NM"								
	arrField(3) = "DD10" & PopupParent.gColSep & "A.SO_DT"								
	arrField(4) = "ED15" & PopupParent.gColSep & "C.SALES_GRP_NM"						
	arrField(5) = "ED10" & PopupParent.gColSep & "A.PAY_METH"							
		
	arrHeader(0) = "수주번호"											
	arrHeader(1) = "주문처"												
	arrHeader(2) = "주문처명"											
	arrHeader(3) = "수주일"												
	arrHeader(4) = "영업그룹명"											
	arrHeader(5) = "결제방법"											

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtSONo.focus
	
	If arrRet(0) <> "" Then
		frm1.txtSONo.Value = arrRet(0)
	End If
End Function

'=========================================
Function OpenCCNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "통관관리번호"							
	arrParam(1) = "S_CC_HDR A, B_BIZ_PARTNER B"					
	arrParam(2) = Trim(frm1.txtCCNo.value)						
	arrParam(3) = ""											
	arrParam(4) = "A.APPLICANT = B.BP_CD"						
	arrParam(5) = "통관관리번호"							

	arrField(0) = "ED15" & PopupParent.gColSep & "A.CC_NO"					
	arrField(1) = "ED12" & PopupParent.gColSep & "A.APPLICANT"				
	arrField(2) = "ED20" & PopupParent.gColSep & "B.BP_NM"					
	arrField(3) = "ED12" & PopupParent.gColSep & "A.IV_NO"					
	arrField(4) = "DD12" & PopupParent.gColSep & "A.IV_DT"					


	arrHeader(0) = "통관관리번호"							
	arrHeader(1) = "수입자"									
	arrHeader(2) = "수입자명"								
	arrHeader(3) = "송장번호"								
	arrHeader(4) = "작성일"									

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=646px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtCCNo.focus
	
	If arrRet(0) <> "" Then
		frm1.txtCCNo.Value = arrRet(0)
	End If
End Function


'========================================
Function OpenDnReqNo
	Dim iCalledAspName
	Dim strRet
	If lgIsOpenPop = True Then Exit Function
			
	lgIsOpenPop = True

	iCalledAspName = AskPRAspName("S4511PA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4511PA1", "x")
		lgIsOpenPop = False
		exit Function
	end if
		
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.PopupParent), _
		"dialogWidth=646px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtDnReqNo.focus
			
	If strRet <> "" Then
		frm1.txtDnReqNo.value = strRet		
	End If	
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
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

'========================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Sub ZbtnBpCd_OnClick()
	Call OpenBizPartner()
End Sub

'========================================
Sub ZbtnSalesGroup_OnClick()
	Call OpenSalesGroup()
End Sub

'========================================
Sub ZbtnMovType_OnClick()
	Call OpenMinorCd()
End Sub

'========================================
Sub txtDNFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDNFrDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDNFrDt.Focus
    End If
End Sub

'========================================
Sub txtDNToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDNToDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDNToDt.Focus
    End If
End Sub

'========================================
Sub txtDlvyFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyFrDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDlvyFrDt.Focus
    End If
End Sub

'========================================
Sub txtDlvyToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDlvyToDt.Action = 7 
		Call SetFocusToDocument("P")
		frm1.txtDlvyToDt.Focus
    End If
End Sub

'========================================
Sub txtDNFrDt_KeyDown(KeyCode,Shift)
     On Error Resume Next
     If KeyCode = 27 Then
        Call CancelClick()
     Elseif KeyCode = 13 Then
        Call FncQuery()
     End If
End Sub

'========================================
Sub txtDNToDt_Keypress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 27 Then
        Call CancelClick()
     Elseif KeyAscii = 13 Then
        Call FncQuery()
     End if
End Sub

'========================================
Sub txtDlvyFrDt_Keypress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 27 Then
        Call CancelClick()
     Elseif KeyAscii = 13 Then
        Call FncQuery()
     End if
End Sub

'========================================
Sub txtDlvyToDt_Keypress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 27 Then
        Call CancelClick()
     Elseif KeyAscii = 13 Then
        Call FncQuery()
     End if
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Sub
	
'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If CheckRunningBizProcess Then Exit Sub
		If lgPageNo <> "" Then Call DbQuery
	End If		 
End Sub

'========================================
Function OKClick()
		
	dim arrReturn
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1) ' 1
		arrReturn = frm1.vspdData.Text
			
		Self.Returnvalue = arrReturn
	End If
	Self.Close()
End Function

'========================================
Function CancelClick()
	Self.Close()
End Function

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	Dim IntRetCD

	If ValidDateCheck(frm1.txtDlvyFrDt, frm1.txtDlvyToDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtDNFrDt, frm1.txtDNToDt) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
	If frm1.rdoDNFlg1.checked = True Then
		frm1.txtRadio.value = "A"
	ElseIf frm1.rdoDNFlg2.checked = True Then
		frm1.txtRadio.value = "Y"
	ElseIf frm1.rdoDNFlg3.checked = True Then
		frm1.txtRadio.value = "N"
	End If			   	

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=====================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
		 
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then	
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtBpCd=" & Trim(.HBpCd.value)	
			strVal = strVal & "&txtSalesGroup=" & Trim(.HSalesGroup.value)
			strVal = strVal & "&txtMovType=" & Trim(.HMovType.value)
			strVal = strVal & "&txtRadio=" & Trim(.HRadio.value)
			strVal = strVal & "&txtSONO=" & Trim(.HSONO.value)
			strVal = strVal & "&txtDnReqNo=" & Trim(.HDnReqNo.value)
			strVal = strVal & "&txtCCNO=" & Trim(.HCCNo.value)
			strVal = strVal & "&txtDlvyFrDt=" & Trim(.HDlvyFrDt.value)
			strVal = strVal & "&txtDlvyToDt=" & Trim(.HDlvyToDt.value)
			strVal = strVal & "&txtDNFrDt=" & Trim(.HDNFrDt.value)
			strVal = strVal & "&txtDNToDt=" & Trim(.HDNToDt.value)
			strVal = strVal & "&txtExceptFlag=" & Trim(.txtHExceptFlag.value)
		Else			
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)	
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtMovType=" & Trim(.txtMovType.value)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
			strVal = strVal & "&txtSONO=" & Trim(.txtSONO.value)
			strVal = strVal & "&txtDnReqNo=" & Trim(.txtDnReqNo.value)
			strVal = strVal & "&txtCCNO=" & Trim(.txtCCNo.value)
			strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
			strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
			strVal = strVal & "&txtDNFrDt=" & Trim(.txtDNFrDt.text)
			strVal = strVal & "&txtDNToDt=" & Trim(.txtDNToDt.text)
			strVal = strVal & "&txtExceptFlag=" & Trim(.txtHExceptFlag.value)
		End If
		
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        strVal = strVal & "&lgPageNo="		 & lgPageNo						
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True


End Function

'=====================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBpCd.focus
	End If
	
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
						<TD CLASS=TD5 NOWRAP>납품처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnBpCd_OnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnSalesGroup_OnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>	
						<TD CLASS=TD5 NOWRAP>출하형태</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtMovType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU" ALT="출하형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:ZbtnMovType_OnClick">&nbsp;
							<INPUT NAME="txtMovTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>출고예정일</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDNFrDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="출고예정시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDNToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="출고예정종료일"></OBJECT>');</SCRIPT>
						</TD>
					</TR>	
					<TR>
						<%If Request("txtExceptFlag") = "Y" Then%>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE="Hidden" NAME="txtSONo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="수주번호"></TD>
						<%Else%>
							<TD CLASS=TD5 NOWRAP>수주번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSONo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="수주번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSONo" align=top TYPE="BUTTON" OnClick="vbscript:OpenSONo"></TD>
						<%End If%>
						<%If Request("txtExceptFlag") = "Y" Then%>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyFrDt" style="HEIGHT: 0px; WIDTH: 0px" tag="11X1" Title="FPDATETIME" ALT="납기시작일"></OBJECT>');</SCRIPT>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyToDt" style="HEIGHT: 0px; WIDTH: 0px" tag="11X1" Title="FPDATETIME" ALT="납기종료일"></OBJECT>');</SCRIPT>
							</TD>
						<%Else%>
							<TD CLASS=TD5 NOWRAP>납기일</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyFrDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="납기시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="납기종료일"></OBJECT>');</SCRIPT>
							</TD>
						<%End If%>
					</TR>
					<TR>
						<%If Request("txtExceptFlag") = "Y" Then%>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE="Hidden" NAME="txtCCNo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="통관관리번호"></TD>
						<%Else%>
							<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="통관관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" align=top TYPE="BUTTON" OnClick="vbscript:OpenCCNo"></TD>
						<%End If%>
						<TD CLASS=TD5 NOWRAP>출고여부</TD> 
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="A" ID="rdoDNFlg1"><LABEL FOR="rdoDNFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="Y" ID="rdoDNFlg2"><LABEL FOR="rdoDNFlg2">출고</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDNFlg" TAG="11X" VALUE="N" CHECKED ID="rdoDNFlg3"><LABEL FOR="rdoDNFlg3">미출고</LABEL>			
						</TD>
					</TR>	
					<TR>
						<%If Request("txtExceptFlag") = "Y" Then%>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE="Hidden" NAME="txtDnReqNo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="출하요청번호"></TD>
						<%Else%>
							<TD CLASS=TD5 NOWRAP>출하요청번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDnReqNo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="출하요청번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnReqNo" align=top TYPE="BUTTON" OnClick="vbscript:OpenDnReqNo"></TD>
						<%End If%>
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
							                  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHExceptFlag" tag="14">

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="HBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HMovType" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio" tag="24">
<INPUT TYPE=HIDDEN NAME="HSONO" tag="24">
<INPUT TYPE=HIDDEN NAME="HDnReqNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HCCNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HDlvyFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HDlvyToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HDNFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HDNToDt" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
