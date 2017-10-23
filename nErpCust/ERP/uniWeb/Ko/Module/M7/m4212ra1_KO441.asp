<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4212ra1.asp																*
'*  4. Program Name         : 통관내역참조(입고등록ADO)																			*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/05/28																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Kim Jin Ha
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "m4212rb1_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 27                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent
Dim arrReturn					
Dim arrParent
Dim arrParam
Dim EndDate, StartDate

'================================================================================================================================
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
	
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                     'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	Err.Clear
 Call GetValue_ko441()	
	frm1.txtFrCcDt.text = StartDate
    frm1.txtToCCDt.text = EndDate
    
	frm1.hdnSupplierCd.value 	= arrParam(0)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.txtGroupCd.value 		= arrParam(2)
	frm1.hdnGroupNm.value 		= arrParam(3)
	frm1.txtGroupNm.value 		= arrParam(3)
	frm1.hdnPlantCd.Value		= arrParam(4)
	'frm1.hdnSubcontraFlg.Value	= arrParam(5)
	frm1.hdnRcptType.Value		= arrParam(5)

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

	Call SetZAdoSpreadSheet("m4212ra1","S","A","V20030528",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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

			Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols - 2)
			
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
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
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
Function OpenCcNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
			
	If gblnWinEvent = True Or UCase(frm1.txtCcNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
	
	gblnWinEvent = True
		
 	iCalledAspName = AskPRAspName("M4211PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M4211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If strRet = "" Then	
		frm1.txtCcNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtCcNo.value = strRet 
		frm1.txtCcNo.focus	
		Set gActiveElement = document.activeElement
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
Sub txtFrCcDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToCcDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrCcDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrCcDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrCcDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToCcDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToCcDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToCcDt.Focus
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    FncQuery = False                                                 
    
    Err.Clear                                                        

	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrCcDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToCcDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrCcDt.text) <> "" And Trim(.txtToCcDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","통관일", "X")	
			.txtFrCcDt.Focus()
			Exit Function
		End if   
	End with
	
	Call InitVariables												
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
    
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
		strVal = strVal & "&txtCcNo=" & Trim(frm1.hdnCcNo.value)
		strVal = strVal & "&txtFrCcDt=" & Trim(frm1.hdnFrCcDt.value)
		strVal = strVal & "&txtToCcDt=" & Trim(frm1.hdnToCcDt.value)
		strVal = strVal & "&txtGroup=" & Trim(frm1.hdnGroupCd.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.hdnPlantCd.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
		strVal = strVal & "&txtCcNo=" & Trim(frm1.txtCcNo.Value)
		strVal = strVal & "&txtFrCcDt=" & Trim(frm1.txtFrCcDt.text)
		strVal = strVal & "&txtToCcDt=" & Trim(frm1.txtToCcDt.text)
		strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlantCd.value)
	End if 
		strVal = strVal & "&txtTrackingNo="	& Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplierCd.value)
		strVal = strVal & "&txtRcptType=" & Trim(frm1.hdnRcptType.value)
		'strVal = strVal & "&txtSubcontraFlg=" & Trim(frm1.hdnSubcontraFlg.value)
		
		strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  
 
		Call RunMyBizASP(MyBizASP, strVal)								

		DbQuery = True														
End Function
'================================================================================================================================
Function DbQueryOk()														

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	':	frm1.vspdData.SelModeSelected = True		2005.07.13 Tracking No. 9913
	Else
		frm1.txtSoNo.focus
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
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
						<TD CLASS="TD5" NOWRAP>통관일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=통관일 NAME="txtFrCcDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td>~</td>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=통관일 NAME="txtToCcDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>				   
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>통관번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCcNo" SIZE=32 MAXLENGTH=18 ALT="통관번호" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCcNo()"></TD>						
						<TD CLASS=TD5 NOWRAP>구매그룹</TD> 
						<TD CLASS=TD6 colspan=3 NOWRAP>
						<INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
                    </TR>		
                    <TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
											   <INPUT TYPE=TEXT AlT="공장" ID="txtPlantNm" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=Text ALT="Tracking번호" NAME="txtTrackingNo"   MAXLENGTH=25 SiZE=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></td>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% ID=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnCcNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrCcDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToCcDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
