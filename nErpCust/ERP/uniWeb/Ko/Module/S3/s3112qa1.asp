<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3112qa1
'*  4. Program Name         : 수주상세조회 
'*  5. Program Desc         : 수주상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                              

Dim lgMark                                                  
Dim IscookieSplit

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3112qb1.asp"
Const BIZ_PGM_JUMP_ID	= "s3112ma1"
Const C_MaxKey          = 32                                       

Const C_PopItemCd		= 1
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'=============================================================================================================
Sub InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE
    lgPageNo         = ""
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgSortKey        = 1
    
End Sub

'=============================================================================================================
Sub SetDefaultVal()	
	'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	frm1.txtSOFrDt.text = StartDate
	frm1.txtSOToDt.text = EndDate
	frm1.txtRadio.value = frm1.rdoQueryFlg1.value 
	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "MA") %>
End Sub


'=============================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3112QA1","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock
End Sub


'=============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub


'=============================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'=============================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	on Error Resume Next
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("s2210pa1")

	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function


'============================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
			Case C_PopItemCd
				.txtItem_cd.value = pvArrRet(0) 
				.txtItem_nm.value = pvArrRet(1)   
		End Select
	End With
	
	SetConPopup = True

End Function


'=============================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		arrParam(4) = "BP_TYPE in ('C','CS')"				
		arrParam(5) = "주문처"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "주문처"							
		arrHeader(1) = "주문처명"						

	Case 1
		arrParam(1) = "S_SO_TYPE_CONFIG"					
		arrParam(2) = Trim(frm1.txtSOType.Value)			
		arrParam(4) = "USAGE_FLAG = 'Y'"									
		arrParam(5) = "수주형태"						
	
		arrField(0) = "SO_TYPE"							
		arrField(1) = "SO_TYPE_NM"							
    
		arrHeader(0) = "수주형태"						
		arrHeader(1) = "수주형태명"						

	Case 2
		arrParam(1) = "B_SALES_ORG"							
		arrParam(2) = Trim(frm1.txtSalesOrg.Value)			
		arrParam(4) = ""									
		arrParam(5) = "영업조직"						
	
		arrField(0) = "SALES_ORG"							
		arrField(1) = "SALES_ORG_NM"						
    
		arrHeader(0) = "영업조직"						
		arrHeader(1) = "영업조직명"						

	Case 3
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"						

	Case 4
		Call OpenConItemPopup(C_PopItemCd, frm1.txtItem_cd.value)
		frm1.txtItem_cd.focus
		Exit Function		
	
	Case 5
		'2002-10-07 s3135pa1.asp 추가 
		Dim strRet
		
		Dim arrTNParam(5), i

		For i = 0 to UBound(arrTNParam)
			arrTNParam(i) = ""
		Next	

		'20021227 kangjungu dynamic popup
		iCalledAspName = AskPRAspName("s3135pa1")	
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
			lgIsOpenPop = False
			exit Function
		end if

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrTNParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If strRet = "" Then
			Exit Function
		Else
			frm1.txtTrackingNo.value = strRet 
		End If		
		
		frm1.txtTrackingNo.focus
		Exit Function
	
	End Select

	arrParam(0) = arrParam(5)								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False


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
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 1
			.txtSOType.value = arrRet(0) 
			.txtSOTypeNm.value = arrRet(1)   
		Case 2
			.txtSalesOrg.value = arrRet(0)
			.txtSalesOrgNm.value = arrRet(1)  
		Case 3
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   		
		End Select
	End With
End Function


'===========================================================================
Function OpenSoNo(strSoNo)
	Dim iCalledAspName
	Dim strRet

	If lgIsOpenPop = True Then Exit Function
			
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		lgIsOpenPop = False
		exit Function
	end if

	lgIsOpenPop = True

	strRet = window.showModalDialog(iCalledAspName,Array(Window.parent, "SO_REG"), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		strSoNo.value = strRet
	End If
	
	frm1.txtConSo_no.focus 	

End Function

'=============================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						

	If Kubun = 1 Then		

		If frm1.vspdData.ActiveRow > 0 Then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1)
			WriteCookie CookieSplit , frm1.vspdData.Text
		Else
			WriteCookie CookieSplit , ""
		End If
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							<% 'Jump로 화면이 이동해 왔을경우 %>

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		
		Dim iniSep

		'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------	
		frm1.txtconBp_cd.value =  arrVal(0)
		frm1.txtconBp_Nm.value =  arrVal(1)
		frm1.txtBillType.value =  arrVal(2)
		frm1.txtBillTypeNm.value = arrVal(3) 
		frm1.txtSalesOrg.value =  arrVal(4)
		frm1.txtSalesOrgNm.value = arrVal(5) 
		frm1.txtSalesGroup.value =  arrVal(6)
		frm1.txtSalesGroupNm.value = arrVal(7) 
		frm1.txtItem_cd.value =  arrVal(8)
		frm1.txtItem_Nm.value = arrVal(9)
		'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'=============================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
  	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------   
	Call CookiePage(0)
	frm1.txtconBp_cd.focus	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'=============================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If    
        
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	IscookieSplit = ""	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
End Sub


'=============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub

	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
		If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If    
End Sub

'=============================================================================================================
Sub rdoQueryFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg1.value
End Sub

Sub rdoQueryFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg2.value
End Sub

Sub rdoQueryFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoQueryFlg3.value
End Sub

'=============================================================================================================
Sub txtSOFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSOFrDt.Action = 7
	End If
End Sub

Sub txtSOToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtSOToDt.Action = 7
	End If
End Sub

'=============================================================================================================
Sub txtSOFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtSOToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'=============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    lgIntFlgMode = parent.OPMD_CMODE
	
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables       

	With frm1
		If .rdoQueryFlg1.checked = True Then
			.txtRadio.value = .rdoQueryFlg1.value
		ElseIf .rdoQueryFlg2.checked = True Then

			.txtRadio.value = .rdoQueryFlg2.value
		ElseIf .rdoQueryFlg3.checked = True Then
			.txtRadio.value = .rdoQueryFlg3.value
		End If		
	End With

    Call DbQuery															

    FncQuery = True		
End Function

'=============================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=============================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=============================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     
End Function


'=============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=============================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'=============================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False

	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1   

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		.txtFlgMode.value = lgIntFlgMode			
		.OPMD_UMODE.value = parent.OPMD_UMODE
		.lgPageNo.value = lgPageNo 
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        
		.lgSelectListDT.value = GetSQLSelectListDataType("A")
        .lgTailList.value = MakeSQLGroupOrderByList("A")
		.lgSelectList.value = EnCoding(GetSQLSelectList("A"))         
		          
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
        
    End With
    
    DbQuery = True

End Function

'=============================================================================================================
Function DbQueryOk()														

	lgIntFlgMode = parent.OPMD_UMODE
  
	Call SetToolbar("11000000000111")

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus   
    End if 
    
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수주상세조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="주문처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 4">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS=TD5 NOWRAP>수주일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSOFrDt" style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME" ALT="수주시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtSOToDt" style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME" ALT="수주종료일"></OBJECT>');</SCRIPT>
									</TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5>수주형태</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSOType" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSORef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtSOTypeNm" SIZE=20 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>수주번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSo_no" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="11XXXU"  STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo frm1.txtConSo_no"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 5"></TD>	
									<TD CLASS=TD5>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="A" CHECKED ID="rdoQueryFlg1"><LABEL FOR="rdoQueryFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="Y" ID="rdoQueryFlg2"><LABEL FOR="rdoQueryFlg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="N" ID="rdoQueryFlg3"><LABEL FOR="rdoQueryFlg3">미확정</LABEL>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage(1)">수주내역등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">

<INPUT TYPE=HIDDEN NAME="HtxtconBp_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="HtxtSoType" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtSOFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtSOToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtRadio" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtTrackingNo" tag="24">        
<INPUT TYPE=HIDDEN NAME="HtxtConSo_no" tag="24">    

<INPUT TYPE=HIDDEN NAME="OPMD_UMODE" tag="24"> 
<INPUT TYPE=HIDDEN NAME="lgPageNo" tag="24">  
<INPUT TYPE=HIDDEN NAME="lgSelectListDT" tag="24">  
<INPUT TYPE=HIDDEN NAME="lgTailList" tag="24">  
<INPUT TYPE=HIDDEN NAME="lgSelectList" tag="24">  
	

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>