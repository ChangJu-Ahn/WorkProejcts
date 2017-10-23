<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3111pa1.asp	
'*  4. Program Name         : 수주번호팝업 
'*  5. Program Desc         : 수주번호팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
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

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

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
Dim lgMark                                                 
Dim IscookieSplit

Dim StoFlag
DIm PopFlag

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3111pb1.asp"
Const C_MaxKey          = 13                           
Const gstPaytermsMajor = "B9004"                                       
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'==================================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                                 
    lgSortKey        = 1
	lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE              'Indicates that current mode is Create mode
End Sub

'==================================================================================================================
Sub SetDefaultVal()	
	frm1.txtSOFrDt.text = StartDate
	frm1.txtSOToDt.text = EndDate	
	Select case arrParent(1)
		case ""
			StoFlag = ""
			
		case "SO_REG"
			StoFlag = "N"
			
		case "STO"
			StoFlag = "Y"
			
		case "ALLOCATION"	
			frm1.rdoComfirmFlg2.checked = True
			Call ggoOper.SetReqAttr(frm1.rdoComfirmFlg1, "Q")
			Call ggoOper.SetReqAttr(frm1.rdoComfirmFlg2, "Q")
			Call ggoOper.SetReqAttr(frm1.rdoComfirmFlg3, "Q")
			PopFlag = "alloc"
			
		case "invoice"
			PopFlag = "invoice"
			
	End Select
End Sub

'==================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>		
End Sub

'==================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3111pa1","S","A","V20030320", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock      
End Sub

'==================================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'==================================================================================================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "주문처"						
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.txtBpCd.value)				
	arrParam(3) = ""									
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
	arrParam(5) = "주문처"							
		
	arrField(0) = "BP_CD"								
	arrField(1) = "BP_NM"								
		
	arrHeader(0) = "주문처"							
	arrHeader(1) = "주문처명"						
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function

'==================================================================================================================
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = strPopPos								
	arrParam(1) = "B_Minor"								
	arrParam(2) = Trim(strMinorCD)						
	arrParam(3) = ""						            
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		
	arrParam(5) = strPopPos								

	arrField(0) = "Minor_CD"							
	arrField(1) = "Minor_NM"							
	arrHeader(0) = strPopPos							
	arrHeader(1) = strPopPos & "명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(strMajorCd, arrRet)
	End If
End Function

'==================================================================================================================
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

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function

'==================================================================================================================
Function OpenSOType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주형태"					
	arrParam(1) = "S_SO_TYPE_CONFIG"				
	arrParam(2) = Trim(frm1.txtSo_Type.value)		
	arrParam(3) = ""								
	
	Select case arrParent(1)
		case ""
			arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "
			
		case "SO_REG"
			arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  and STO_FLAG = " & FilterVar("N", "''", "S") & " "							
			
		case "STO"
			arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  and STO_FLAG = " & FilterVar("Y", "''", "S") & " "				
			
		case "ALLOCATION"	
			arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  and RET_ITEM_FLAG = " & FilterVar("N", "''", "S") & " "		
			
	End Select
	
	arrParam(5) = "수주형태"					
		
    arrField(0) = "SO_TYPE"							
    arrField(1) = "SO_TYPE_NM"						
	    
    arrHeader(0) = "수주형태"					
    arrHeader(1) = "수주형태명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSOType(arrRet)
	End If	
End Function

'==================================================================================================================
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

'==================================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtBpCd.value = arrRet(0)
	frm1.txtBpNm.value = arrRet(1)
	frm1.txtBpCd.focus
End Function

'==================================================================================================================
Function SetMinorCd(strMajorCd, arrRet)
	frm1.txtPay_terms.value = arrRet(0)
	frm1.txtPay_terms_nm.value = arrRet(1)
	frm1.txtPay_terms.focus
End Function

'==================================================================================================================
Function SetSOType(arrRet)
	frm1.txtSo_Type.value = arrRet(0)
	frm1.txtSo_TypeNm.value = arrRet(1)
	frm1.txtSo_Type.focus
End Function

'==================================================================================================================
Function SetSalesGroup(arrRet)
	frm1.txtSalesGroup.Value = arrRet(0)
	frm1.txtSalesGroupNm.Value = arrRet(1)
	frm1.txtSalesGroup.focus
End Function	

'==================================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'==================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==================================================================================================================
Sub btnBpCdOnClick()
	Call OpenBizPartner()
End Sub

'==================================================================================================================
Sub btnSalesGroupOnClick()
	Call OpenSalesGroup()
End Sub

'==================================================================================================================
Sub btnSoTypeOnClick()
	Call OpenSOType()
End Sub

'==================================================================================================================
Sub btnPayTermsOnClick()
	Call OpenMinorCd(frm1.txtPay_terms.value, frm1.txtPay_terms_nm.value, "결제방법", gstPaytermsMajor)
End Sub

'==================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'==================================================================================================================
Sub txtSOFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOFrDt.Action = 7 
        Call SetFocusToDocument("P")
		frm1.txtSOFrDt.Focus
    End If
End Sub

'==================================================================================================================
Sub txtSOToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOToDt.Action = 7 
        Call SetFocusToDocument("P")
		frm1.txtSOToDt.Focus
    End If
End Sub

'==================================================================================================================
Sub txtSOFrDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtSOToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'==================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then  
        Exit Sub
    End If
	
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
	

'==================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    					
    	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub

'==================================================================================================================
Function OKClick()
		
	dim arrReturn
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		arrReturn = frm1.vspdData.Text

		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'==================================================================================================================
Function CancelClick()
	Self.Close()
End Function

'==================================================================================================================
Function FncQuery() 
	Dim IntRetCD
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
  
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function

	If frm1.rdoComfirmFlg1.checked = True Then
		frm1.txtRadio.value = "A"
	ElseIf frm1.rdoComfirmFlg2.checked = True Then
		frm1.txtRadio.value = "Y"
	ElseIf frm1.rdoComfirmFlg3.checked = True Then
		frm1.txtRadio.value = "N"
	End If			   	

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'==================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'==================================================================================================================
Function FncExcel() 
	Call parent.FncExport(C_MULTI)
End Function

'==================================================================================================================
Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                    
End Function

'==================================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'==================================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If

    With frm1

		If lgIntFlgMode = PopupParent.OPMD_UMODE Then	
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtBpCd=" & Trim(.HBpCd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.HSalesGroup.value)
			strVal = strVal & "&txtSo_Type=" & Trim(.HSo_Type.value)
			strVal = strVal & "&txtPay_terms=" & Trim(.HPay_terms.value)
			strVal = strVal & "&txtRadio=" & Trim(.HRadio.value)
			strVal = strVal & "&txtSOFrDt=" & Trim(.HSOFrDt.value)
			strVal = strVal & "&txtSoToDt=" & Trim(.HSoToDt.value)
			strVal = strVal & "&txtSTOFlag=" & StoFlag
			strVal = strVal & "&txtPopFlag=" & PopFlag

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)    	
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtSo_Type=" & Trim(.txtSo_Type.value)
			strVal = strVal & "&txtPay_terms=" & Trim(.txtPay_terms.value)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
			strVal = strVal & "&txtSOFrDt=" & Trim(.txtSOFrDt.text)
			strVal = strVal & "&txtSoToDt=" & Trim(.txtSoToDt.text)
			strVal = strVal & "&txtSTOFlag=" & StoFlag
			strVal = strVal & "&txtPopFlag=" & PopFlag
		End If
		       
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function

'==================================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus		
    Else
       frm1.txtBpCd.focus
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
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnBpCdOnClick()">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnSalesGroupOnClick()">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>	
						<TD CLASS=TD5 NOWRAP>수주형태</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtSo_Type" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoType" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnSoTypeOnClick()">&nbsp;
							<INPUT NAME="txtSo_TypeNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24">
						</TD>
						<TD CLASS=TD5 NOWRAP>수주일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s3111pa1_fpDateTime2_txtSOFrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s3111pa1_fpDateTime2_txtSoToDt.js'></script>
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>결제방법</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtPay_terms" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnPayTermsOnClick()">&nbsp;
							<INPUT NAME="txtPay_terms_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24">
						</TD>
						<TD CLASS=TD5 NOWRAP>확정여부</TD> 
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="A" ID="rdoComfirmFlg1"><LABEL FOR="rdoComfirmFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="Y" ID="rdoComfirmFlg2"><LABEL FOR="rdoComfirmFlg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="N" CHECKED ID="rdoComfirmFlg3"><LABEL FOR="rdoComfirmFlg3">미확정</LABEL>			
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
						<script language =javascript src='./js/s3111pa1_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                     <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
<INPUT TYPE=HIDDEN NAME="HBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HSo_Type" tag="24">
<INPUT TYPE=HIDDEN NAME="HPay_terms" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadio" tag="24">
<INPUT TYPE=HIDDEN NAME="HSOFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="STOFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="PopFlag" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
