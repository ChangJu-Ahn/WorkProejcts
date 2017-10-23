<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : 매입관리																	*
'*  3. Program ID           : m3112ra6																	*
'*  4. Program Name         : 발주내역참조																*
'*  5. Program Desc         : 매입내역등록을 위한 발주내역참조 (Business Logic Asp)						*
'*  6. Comproxy List        : M31228ListPoDtlRefInIvSvr													*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>발주내역참조</TITLE>
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
<Script Language="VBScript">

Option Explicit		

	
Const BIZ_PGM_ID 		= "m3112rb6_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 32                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgCookValue 
Dim IsOpenPop  
Dim gblnWinEvent
Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)


Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
			
	if Trim(frm1.hdnPoNo.value) <> "" Then
	   frm1.txtPoNo.value =  Trim(frm1.hdnPoNo.value) 
	   Call ggoOper.SetReqAttr(frm1.txtPoNo,"Q")
	end if 

	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	Dim arrParam

	arrParam = arrParent(1)
		
	frm1.hdnSupplierCd.value 	= arrParam(0)		
	'frm1.txtSupplierNm.value 	= arrParam(1)
	frm1.hdnGroupCd.value 		= arrParam(2)		
	frm1.hdnPlant.value 		= arrParam(3)		
	frm1.hdnClsflg.value 		= arrParam(4)		
	frm1.hdnReleaseflg.value 	= arrParam(5)		
	frm1.hdnRcptflg.value 		= arrParam(6)
	frm1.hdnRetflg.value 		= arrParam(7)
	frm1.hdnRefType.value 		= arrParam(8)
	frm1.hdnRcptType.value 		= arrParam(9)
	frm1.hdnIvflg.value 		= arrParam(10)
	frm1.hdnIvType.value 		= arrParam(11)
	frm1.hdnPoType.value 		= arrParam(12)	
	frm1.hdnPoNo.value 			= arrParam(13)
	frm1.hdnPoCur.value 		= arrParam(14)

	frm1.txtFrPoDt.Text = StartDate
	frm1.txtToPoDt.Text = EndDate

	frm1.txtPlantCd.value = PopupParent.gPlant
	frm1.txtPlantNm.value = PopupParent.gPlantNm

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'==================================  LoadInfTB19029()  ==================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
End Sub
'==================================  InitSpreadSheet()  ==================================================
Sub InitSpreadSheet()
	    
    Call SetZAdoSpreadSheet("M3112RA6","S","A","V20031007",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")
	Call SetSpreadLock
	frm1.vspdData.OperationMode = 5
	      
End Sub
'==================================  SetSpreadLock()  ==================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()

	Dim intColCnt, intRowCnt, intInsRow
	ReDim arrReturn(0,0)
	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		ReDim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

		For intRowCnt = 0 To frm1.vspdData.MaxRows -1 

			frm1.vspdData.Row = intRowCnt + 1

			If frm1.vspdData.SelModeSelected Then

				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = GetKeyPos("A", (intColCnt + 1))
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
'++++++++++++++++++++++++++++++++++++++++++++  OpenPoNo()  ++++++++++++++++++++++++++++++++++++++++++++++
Function OpenPoNo()
	
	Dim strRet
	Dim lblnWinEvent
	Dim iCalledAspName
	Dim arrParam(2)
	
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
	
	iCalledAspName = AskPRAspName("m3111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m3111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'++++++++++++++++++++++++++++++++++++++++++++  OpenPlant()  ++++++++++++++++++++++++++++++++++++++++++++++
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
	End If	
	
End Function

'++++++++++++++++++++++++++++++++++++++++++++  OpenItem()  ++++++++++++++++++++++++++++++++++++++++++++++
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	if Trim(frm1.txtPlantCd.Value) = "" then
		'Call DisplayMsgBox("17A002", "X", "공장", "X")
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
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & ""  
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
    				
    iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.PopupParent, arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)	
		frm1.txtItemCd.focus	
		Set gActiveElement = document.activeElement	
	End If	
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
'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables							
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

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
'   Event Name : OCX_DbClick()
'==========================================================================================
Sub txtFrPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrPoDt.focus
	End If
End Sub
Sub txtToPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToPoDt.focus
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToPoDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'+++++++++++++++++++++++++++++++++++++++++++  OpenVatType()  ++++++++++++++++++++++++++++++++++++++++++++++++
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
	End If	
	Set gActiveElement = document.activeElement
    IsOpenPop = False
End Function	
'==================================  FncQuery()  ===========================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	If ValidDateCheck(frm1.txtFrPoDt, frm1.txtToPoDt) = False Then Exit Function

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	Call InitVariables												

    If DbQuery = False Then Exit Function							

    FncQuery = True									
    Set gActiveElement = document.activeElement    
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
		
	With frm1
		
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtPoNo=" & .hdnPoNo.Value					'발주번호 
		strVal = strVal & "&txtFrPoDt=" & .hdnFrPoDt.value				'발주일 
		strVal = strVal & "&txtToPoDt=" & .hdnToPoDt.value	
        strVal = strVal & "&txtItemCd=" & Trim(frm1.hdnItemCd.Value)	'품목 
        strVal = strVal & "&txtPlantCd=" & Trim(frm1.hdnPlantCd.Value)	'공장 
		strVal = strVal & "&txtVatType=" & Trim(frm1.hdnVatType.Value)      			
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		if Trim(.hdnPoNo.Value) = "" then
           strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.Value)			'발주번호 
		Else
           strVal = strVal & "&txtPoNo=" & Trim(.hdnPoNo.Value)
        End if
		strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)			'발주일 
		strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)		
	    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)	'품목 
        strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)	'공장 
		strVal = strVal & "&txtVatType=" & Trim(frm1.txtVatType.Value)				
	End If
	strVal = strVal & "&txtTrackingNo=" &Trim(frm1.txtTrackingNo.value)

	strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value			
	strVal = strVal & "&txtGroup=" & .hdnGroupCd.value
	strVal = strVal & "&txtClsflg=" & .hdnClsflg.value
	strVal = strVal & "&txtreleaseflg=" & .hdnReleaseflg.value
	strVal = strVal & "&txtRcptflg=" & .hdnRcptflg.value
	strVal = strVal & "&txtRetflg=" & .hdnRetflg.value
	strVal = strVal & "&txtRefType=" & .hdnRefType.value
	strVal = strVal & "&txtRcptType=" & .hdnRcptType.value
	strVal = strVal & "&txtIvflg=" & .hdnIvflg.value
	strVal = strVal & "&txtIvType=" & .hdnIvType.value
	strVal = strVal & "&txtPoType=" & .hdnPoType.value
	strVal = strVal & "&txtPlant=" & .hdnPlant.value
	strVal = strVal & "&txtPoCur=" & Trim(.hdnPoCur.value)
		
	End With
		
    strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
    strVal =	 strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
	strVal =	 strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal =	 strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True														<%'⊙: Processing is NG%>
End Function

'====================================  DbQueryOk()  ======================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPoNo.focus
	End If

End Function
'========================================================================================
' Function Name : changeItemPlant()
'========================================================================================
Function changeItemPlant()
   Dim strVal
    
   If gLookUpEnable = False Then Exit Function
   Err.Clear
    
   If CheckRunningBizProcess = True Then
   		Exit Function
   End If                               
    
   if Trim(frm1.txtPlantCd.Value) = "" or Trim(frm1.txtItemCd.Value) = "" then
   		Exit Function
   End if
    
   changeItemPlant = False                 
    
   If LayerShowHide(1) = False Then Exit Function
        
   strVal = BIZ_PGM_ID & "?txtMode=" & "changeItemPlant"
   strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.Value)
   strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.Value)
    
   Call RunMyBizASP(MyBizASP, strVal)
	
   changeItemPlant = True                  

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
						<TD CLASS="TD5" NOWRAP>발주번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"><div style="Display:none"><input type="text" name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>발주일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m3112ra6_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m3112ra6_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU" ALT="공장" onChange="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X" ALT="공장"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemcd" SIZE=10 MAXLENGTH=18 tag="11NXXU" onChange="vbscript:changeItemPlant()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
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
						<script language =javascript src='./js/m3112ra6_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>

		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoCur" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnVatType" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
