<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4111ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open Po Ref Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : park no yeol														*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--<TITLE>입출고참조</TITLE> -->
<TITLE></TITLE>
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

Option Explicit                              '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "m4111rb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 38                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IscookieSplit 
dim lblnWinEvent
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 

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
	Dim arrParam
	   
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
		
	
	arrParam = arrParent(1)
		
	frm1.hdnMvmtType.value  	= arrParam(0)
	frm1.hdnSupplierCd.value 	= arrParam(1)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.hdnRefType.value 		= arrParam(3)
	frm1.hdnIvType.value 		= arrParam(4)
	frm1.hdnPoNo.value 			= arrParam(5)
	frm1.hdnPoCur.value 		= arrParam(6)
	'수정(2003.03.24)
	frm1.hdnLcKind.value 		= arrParam(7)
	frm1.hdnPayMeth.value 		= arrParam(8)
	'추가(2005.10.28)
	frm1.hdnIvDt.value			= arrParam(9)

	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn

End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
	
	frm1.txtFrMvmtDt.Text = StartDate
	frm1.txtToMvmtDt.Text = EndDate
		
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M4111RA1","S","A","V20030921",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")
	Call SetSpreadLock 
	frm1.vspdData.OperationMode = 5
End Sub
'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	
'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
	If .vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(.vspdData.SelModeSelCount - 1, .vspdData.MaxCols - 1)

		For intRowCnt = 0 To .vspdData.MaxRows - 1

			.vspdData.Row = intRowCnt + 1
			If .vspdData.SelModeSelected Then
				For intColCnt = 0 To .vspdData.MaxCols - 1
					frm1.vspdData.Col = GetKeyPos("A", (intColCnt + 1))
					arrReturn(intInsRow, intColCnt) = .vspdData.Text
				Next

				intInsRow = intInsRow + 1

			End IF
		Next
	End if			
	end with

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
'------------------------------------------  OpenMvmtNo()  -------------------------------------------------
Function OpenMvmtNo()

	Dim strRet
	Dim arrParam(3)
	
	If lblnWinEvent = True Or UCase(frm1.txtMvmtNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
	
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnMvmtType.Value)
	arrParam(1) = Trim(frm1.hdnSupplierCd.Value)
	arrParam(2) = ""  'Trim(frm1.hdnGroupCd.Value)
	arrParam(3) = ""
	
	strRet = window.showModalDialog("M4111pa4.asp",  Array(PopupParent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtMvmtNo.focus
		Exit Function
	Else
		frm1.txtMvmtNo.value = strRet(0)
		frm1.txtMvmtNo.focus	
		Set gActiveElement = document.activeElement
	End If	
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
		Set gActiveElement = document.activeElement
		'lgBlnFlgChgValue = True
		'Call changeItemPlant()
	End If	
	
End Function
'------------------------------------------  OpenItem()  -------------------------------------------------
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
		Set gActiveElement = document.activeElement
		'lgBlnFlgChgValue = True
		'Call changeItemPlant()
	End If	
End Function
'------------------------------------------  OpenVatType()  -----------------------------------------------
Function OpenVatType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	arrHeader(0) = "VAT형태"									' Header명(0)
    arrHeader(1) = "VAT형태명"									' Header명(1)
    arrHeader(2) = "VAT율"									    ' Header명(2)
    
    arrField(0) = "b_minor.MINOR_CD"					            ' Field명(0)
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"					    ' Field명(1)
    
	arrParam(0) = "VAT"	            							' 팝업 명칭 
	arrParam(1) = "B_MINOR,b_configuration"
	arrParam(2) = Trim(frm1.txtVatType.Value)						    ' Code Condition
	'arrParam(3) = Trim(frm1.txtVatNm.Value)						' Name Cindition
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd and b_configuration.seq_no=1 and b_minor.major_cd=b_configuration.major_cd"
	arrParam(5) = "VAT"										    ' TextBox 명칭 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    If arrRet(0) = "" Then
		frm1.txtVatType.focus
		Exit Function
	Else
		frm1.txtVatType.Value = arrRet(0)
		frm1.txtVatNm.Value = arrRet(1)	
		frm1.txtVatType.focus	
		Set gActiveElement = document.activeElement
	End If	
	
    IsOpenPop = False
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
'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'발주단가 
		ggoSpread.SSSetFloatByCellOfCur C_PoPrc,-1, .hdnPoCur.value, PopupParent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
		'발주금액 
		ggoSpread.SSSetFloatByCellOfCur C_PoAmt,-1, .hdnPoCur.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub
'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables	
	frm1.txtPlantCd.value = PopupParent.gPlant
	frm1.txtPlantNm.value = PopupParent.gPlantNm
							
	Call SetDefaultVal	
	Call InitSpreadSheet()

	Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================  OpenSortPopup()  ==============================================
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
'================================  vspdData_KeyPress()  ===========================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

		
Sub txtFrMvmtDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToMvmtDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'==================================   OCX_EVENT    =======================================
Sub txtFrMvmtDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrMvmtDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrMvmtDt.focus
	End if
End Sub

Sub txtToMvmtDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToMvmtDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToMvmtDt.focus
	End if
End Sub
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	If ValidDateCheck(frm1.txtFrMvmtDt, frm1.txtToMvmtDt) = False Then Exit Function
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
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtMvmtNo=" & .hdnMvmtNo.value
			strVal = strVal & "&txtFrDt=" & .hdnFrMvmtDt.value
			strVal = strVal & "&txtToDt=" & .hdnToMvmtDt.value
		    strVal = strVal & "&txtItemCd=" & Trim(frm1.hdnItemCd.Value)
		    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hdnPlantCd.Value)
			strVal = strVal & "&txtVatType=" & Trim(frm1.hdnVatType.Value)
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtMvmtNo=" & Trim(frm1.txtMvmtNo.Value)
			strVal = strVal & "&txtFrDt=" & Trim(frm1.txtFrMvmtDt.text)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtToMvmtDt.text)
		    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
		    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
		    strVal = strVal & "&txtVatType=" & Trim(frm1.txtVatType.Value)
		End if 
		strVal = strVal & "&txtTrackingNo=" &Trim(frm1.txtTrackingNo.value)	        
		strVal = strVal & "&txtRefType=" & .hdnRefType.value
		strVal = strVal & "&txtIvType=" & .hdnIvType.value
		strVal = strVal & "&txtSppl=" & .hdnSupplierCd.value
		strVal = strVal & "&txtPoNo=" & .hdnPoNo.value
		strVal = strVal & "&txtPoCur=" & .hdnPoCur.value	        
		strVal = strVal & "&txtLcKind=" & .hdnLcKind.value
		strVal = strVal & "&txtPayMeth=" & .hdnPayMeth.value
		strVal = strVal & "&txtIvDt=" & .hdnIvdt.value

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
		frm1.txtDnType.focus
	End If

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
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
						<TD CLASS="TD5" NOWRAP>입고번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmt" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()">
											  <div style="Display:none"><input type=text name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>입고일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입고일 NAME="txtFrMvmtDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td>~</td>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입고일 NAME="txtToMvmtDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
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
						<TD CLASS="TD5" NOWRAP>VAT</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="VAT" NAME="txtVatType" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="VAT"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenVatType()">
											   <INPUT TYPE=TEXT ALT="VAT" NAME="txtVatNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="VAT"></TD>
						<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD >&nbsp;&nbsp; <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>&nbsp;
									  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT>  <IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>&nbsp;
							          <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>&nbsp;&nbsp;</TD>
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

<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrMvmtDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToMvmtDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoCur" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnLcKind" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPayMeth" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvDt" tag="14">

<INPUT TYPE=HIDDEN NAME="lgSelectListDT" tag="14">
<INPUT TYPE=HIDDEN NAME="lgTailList" tag="14">
<INPUT TYPE=HIDDEN NAME="lgSelectList" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
