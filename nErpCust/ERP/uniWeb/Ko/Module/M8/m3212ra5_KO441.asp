<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212ra5.asp																*
'*  4. Program Name         : local l/c 내역참조(매입내역등록)				   					    	*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2003-03-13																*
'*  8. Modified date(Last)  : 2003-09-19																			*
'*  9. Modifier (First)     : Lee Eun Hee																*
'* 10. Modifier (Last)      : 																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--TITLE>LOCAL L/C내역참조</TITLE-->
<TITLE></TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
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
<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "m3212rb5_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 28                                           '☆: key count of SpreadSheet


<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

Dim EndDate, StartDate

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	
	Err.Clear

	frm1.txtFrLCDt.text = StartDate
    frm1.txtToLCDt.text = EndDate
    
	frm1.hdnSupplierCd.value 	= arrParam(0)
	'frm1.hdnGroupCd.value 		= arrParam(1)
	'frm1.txtGroupCd.value 		= arrParam(1)
	frm1.hdnIvType.value 		= arrParam(3)
	frm1.hdnPoNo.value 			= arrParam(4)
	frm1.hdnPoCur.value 		= arrParam(5)
	'***수정(2003.03.26)****
	frm1.hdnLcKind.value 		= arrParam(6)
	frm1.hdnPayMeth.value 		= arrParam(7)


	frm1.txtPlantCd.value		=  PopupParent.gPlant
	frm1.txtPlantNm.value		=  PopupParent.gPlantNm
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'============================================  LoadInfTB19029()  ========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
	
End Sub
'============================================  InitSpreadSheet()  ========================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("m3212ra5","S","A","V20030919",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 5 
End Sub
'============================================  SetSpreadLock()  ========================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End IF
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount -1, frm1.vspdData.MaxCols -2)


		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt

			If frm1.vspdData.SelModeSelected Then
				
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					'frm1.vspdData.Col = intColCnt + 1
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

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenSortPopup
'========================================================================================================
Function OpenSortPopup()
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

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLLCNo()  ++++++++++++++++++++++++++++++++++++++++
-->
Function OpenLLCNo()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtLLCNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3211PA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3211PA2", "X")
		gblnWinEvent = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtLLCNo.focus
		Exit Function
	Else
		frm1.txtLLCNo.value = strRet
		frm1.txtLLCNo.focus
	End If
	Set gActiveElement = document.activeElement	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.Value= arrRet(1)	
		frm1.txtPlantCd.focus	
	End If	
	Set gActiveElement = document.activeElement
	
End Function


'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description :
'---------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	if Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit Function
	end if
	
	IsOpenPop = True

	arrParam(0) = "품목"		
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"
	
	arrParam(2) = Trim(frm1.txtitemCd.Value)	

	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " "  
	End if 
	
	arrParam(5) = "품목"						
	
    arrField(0) = "B_Item.Item_Cd"				
    arrField(1) = "B_Item.Item_NM"	
    arrField(2) = "B_Plant.Plant_Cd"			
    arrField(3) = "B_Plant.Plant_NM"			
    
    arrHeader(2) = "공장"					
    arrHeader(3) = "공장명"					
    
    arrHeader(0) = "품목"					
    arrHeader(1) = "품목명"					

	arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(PopupParent, arrParam, arrField, arrHeader), _
		"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)	
		frm1.txtItemNm.Value    = arrRet(1)	
		frm1.txtItemCd.focus	
	End If	
	Set gActiveElement = document.activeElement
End Function
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

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
'========================================================================================
' Function Name : changeItemPlant()
'========================================================================================
Function changeItemPlant()

	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                               
    
    if Trim(frm1.txtPlantCd.Value) = "" or Trim(frm1.txtItemCd.Value) = "" then
    	exit Function
    End if
    
    changeItemPlant = False                 
    
    If LayerShowHide(1) = False Then Exit Function
        
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemPlant"
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
    Call RunMyBizASP(MyBizASP, strVal)
	
    changeItemPlant = True                  

End Function	

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitVariables														    '⊙: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'=========================================  vspdData_Click()  ===============================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	'Call SetPopupMenuItemInf("0001111111")		
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
'=========================================  3.3.1 vspdData_DblClick()  ==================================
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
    
    If CheckRunningBizProcess = True Then
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

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrLCDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToLCDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'=============================   OCX_EVENT  ================================================
Sub txtFrLCDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrLCDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrLCDt.focus
	End if
End Sub

Sub txtToLCDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToLCDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToLCDt.focus
	End if
End Sub
'=============================  FncQuery  ==============================================
Function FncQuery() 
    
    Err.Clear                                                        
	
	FncQuery = False                                                 
	ggoSpread.Source = frm1.vspdData
	
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrLCDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToLCDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrLCDt.text) <> "" And Trim(.txtToLCDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","L/C개설일", "X")	
			.txtToLCDt.Focus()
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

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Dim strVal
	
	Err.Clear															<%'☜: Protect system from crashing%>
	
	DbQuery = False														<%'⊙: Processing is NG%>
	
	If LayerShowHide(1) = False Then Exit Function 
	
	With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then

			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtLLCNo=" & .hdnLLCNo.value
			strVal = strVal & "&txtFrLCDt=" & Trim(frm1.hdnFrLCDt.value)
			strVal = strVal & "&txtToLCDt=" & Trim(frm1.hdnToLCDt.value)
		    strVal = strVal & "&txtItemCd=" & Trim(frm1.hdnItemCd.Value)
		    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hdnPlantCd.Value)    
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtLLCNo=" & Trim(frm1.txtLLCNo.Value)
			strVal = strVal & "&txtFrLCDt=" & Trim(frm1.txtFrLCDt.text)
			strVal = strVal & "&txtToLCDt=" & Trim(frm1.txtToLCDt.text)
		    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
		    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
		End if 
	    
	    strVal = strVal & "&txtTrackingNo=" &Trim(frm1.txtTrackingNo.value)    
		strVal = strVal & "&txtIvType=" & .hdnIvType.value
		strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value
		strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
		strVal = strVal & "&txtPoCur=" & .hdnPoCur.value	      
		strVal = strVal & "&txtLcKind=" & .hdnLcKind.value
		strVal = strVal & "&txtPayMeth=" & .hdnPayMeth.value

		strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

		Call RunMyBizASP(MyBizASP, strVal)								<%'☜: 비지니스 ASP 를 가동 %>
	End With		
	DbQuery = True														<%'⊙: Processing is NG%>
End Function

'========================  DbQueryOk()  =================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPoNo.focus
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
						<TD CLASS="TD5" NOWRAP>L/C번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="L/C번호" NAME="txtLLCNo" MAXLENGTH=18 SIZE=32 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLLc" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenLLCNo()">
											  <div style="Display:none"><input type=text name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>L/C개설일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3212ra5_fpDateTime1_txtFrLCDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3212ra5_fpDateTime1_txtToLCDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU" ALT="공장" ONCHANGE="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="공장"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemcd" SIZE=10 MAXLENGTH=18 tag="11NXXU" ONCHANGE="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
											   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" SIZE=20 tag="14X"></TD>
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
						<script language =javascript src='./js/m3212ra5_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hdnLLcNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoCur" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnLcKind" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayMeth" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                               
