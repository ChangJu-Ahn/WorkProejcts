<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212ra4.asp																*
'*  4. Program Name         : local l/c 내역참조(입고등록)					   					    	*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/05/19																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "m3212rb4_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 25                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam

Dim EndDate, StartDate
Dim iDBSYSDate
'================================================================================================================================ 
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================ 
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================ 
Sub SetDefaultVal()

	Err.Clear
	
	frm1.txtFrLCDt.text = StartDate
    frm1.txtToLCDt.text = EndDate
    
	frm1.hdnSupplierCd.value 	= arrParam(0)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.txtGroupCd.value 		= arrParam(2)
	frm1.hdnGroupNm.value 		= arrParam(3)
	frm1.txtGroupNm.value 		= arrParam(3)
	frm1.hdnPlantCd.Value		= arrParam(4)

	frm1.txtPlantCd.value		=  PopupParent.gPlant
	frm1.txtPlantNm.value		=  PopupParent.gPlantNm

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
		frm1.txtGroupCd.Tag = left(frm1.txtGroupCd.Tag,1) & "4" & mid(frm1.txtGroupCd.Tag,3,len(frm1.txtGroupCd.Tag))
        frm1.txtGroupCd.value = lgPGCd
	End If
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub
'================================================================================================================================ 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================ 
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("m3212ra4","S","A","V20030528",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 5 
End Sub
'================================================================================================================================ 
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End IF
End Sub
'================================================================================================================================ 
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount -1, frm1.vspdData.MaxCols -2)


		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt

			If frm1.vspdData.SelModeSelected Then
				
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next
				intInsRow = intInsRow + 1
			End IF
		Next
	End if			
	
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	
'================================================================================================================================ 
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================ 
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	arrParam(4) = ""			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)	
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If gblnWinEvent = True Then Exit Function
	
	gblnWinEvent = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	gblnWinEvent = False

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
'================================================================================================================================ 
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
	
	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M3111PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If

End Function

'================================================================================================================================ 
Function OpenSortPopup()
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
'================================================================================================================================ 
Function OpenPlant()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	lgIsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function	
'================================================================================================================================ 
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitVariables														    '⊙: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub
'================================================================================================================================ 
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
    
End Sub
'================================================================================================================================ 
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
'================================================================================================================================ 
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================ 
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
'================================================================================================================================ 
Sub txtFrLCDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================ 
Sub txtToLCDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================ 
Sub txtFrLCDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrLCDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrLCDt.Focus
	End if
End Sub
'================================================================================================================================ 
Sub txtToLCDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToLCDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToLCDt.Focus
	End if
End Sub
'================================================================================================================================ 
Function FncQuery() 
    Err.Clear                                                        
	
	FncQuery = False                                                 
	ggoSpread.Source = frm1.vspdData
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrLCDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToLCDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrLCDt.text) <> "" And Trim(.txtToLCDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","L/C개설일", "X")	
			.txtToLCDt.Focus()
			Exit Function
		End if   
	End with
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
    
	Call InitVariables												
	
	If CheckRunningBizProcess = True Then Exit Function
	If DbQuery = False Then Exit Function
    
    FncQuery = True									
        
End Function
'================================================================================================================================ 
Function DbQuery()
	Dim strVal
	
	Err.Clear															<%'☜: Protect system from crashing%>
	
	DbQuery = False														<%'⊙: Processing is NG%>
	
	If LayerShowHide(1) = False Then Exit Function 
    
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPoNo=" &	Trim(frm1.hdnPoNo.Value)
		strVal = strVal & "&txtFrLCDt=" & Trim(frm1.hdnFrLCDt.value)
		strVal = strVal & "&txtToLCDt=" & Trim(frm1.hdnToLCDt.value)
		strVal = strVal & "&txtGroup=" & frm1.hdnGroupCd.value
		strVal = strVal & "&txtPlant=" & Trim(frm1.hdnPlantCd.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
		strVal = strVal & "&txtFrLCDt=" & Trim(frm1.txtFrLCDt.text)
		strVal = strVal & "&txtToLCDt=" & Trim(frm1.txtToLCDt.text)
		strVal = strVal & "&txtGroup=" & frm1.txtGroupCd.value
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlantCd.value)
	End if 
		strVal = strVal & "&txtTrackingNo="	& Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplierCd.value)
		
		strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

		Call RunMyBizASP(MyBizASP, strVal)								<%'☜: 비지니스 ASP 를 가동 %>
		
	DbQuery = True														<%'⊙: Processing is NG%>
End Function
'================================================================================================================================ 
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPoNo.focus
	End If
End Function
'================================================================================================================================ 
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
						<TD CLASS="TD5" NOWRAP>발주번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"><div style="Display:none"><input type=text name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>L/C개설일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3212ra4_fpDateTime1_txtFrLCDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3212ra4_fpDateTime1_txtToLCDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>구매그룹</TD> 
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
						<TD CLASS=TD5 NOWRAP>공장</TD> 
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
						<INPUT TYPE=TEXT AlT="공장" ID="txtPlantNm" tag="14X">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
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
						<script language =javascript src='./js/m3212ra4_vspdData_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="hdnFrLCDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToLCDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
