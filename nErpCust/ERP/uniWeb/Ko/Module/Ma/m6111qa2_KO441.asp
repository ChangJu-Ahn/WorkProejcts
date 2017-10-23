<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M6111QA2
'*  4. Program Name         :
'*  5. Program Desc         : 경비상세 조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sin jin Hyun
'* 10. Modifier (Last)      : Park JIn Uk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003/05/20
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'						1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit			

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                         
Dim IscookieSplit 
Dim lgSaveRow  

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID 		= "m6111qb2_KO441.asp"                     
Const BIZ_PGM_JUMP_ID1 	= "m6111ma2"                     
Const BIZ_PGM_JUMP_ID2 	= ""                             
Const C_MaxKey          = 20					         
'==============================================================================================================================
Sub InitVariables()
    lgPageNo     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = parent.OPMD_CMODE 
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtChargeFrDt.Text	= StartDate
	frm1.txtChargeToDt.Text	= EndDate
	If lgBACd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtBizArea, "Q") 
		frm1.txtBizArea.Tag = left(frm1.txtBizArea.Tag,1) & "4" & mid(frm1.txtBizArea.Tag,3,len(frm1.txtBizArea.Tag))
        frm1.txtBizArea.value = lgBACd
	End If
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA")%>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M6111QA2","S","A","V20030520",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock 
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'==============================================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtBizArea.className = "protected" Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "사업장"			
	arrParam(1) = "B_BIZ_AREA"	
	arrParam(2) = Trim(frm1.txtBizArea.Value)	
'	arrParam(3) = Trim(frm1.txtBizAreaNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "사업장"					

    arrField(0) = "BIZ_AREA_CD"					
    arrField(1) = "BIZ_AREA_NM"					
    
    arrHeader(0) = "사업장"					
    arrHeader(1) = "사업장명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizArea.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBizArea.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		frm1.txtBizArea.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
Function OpenChargeType()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "경비항목"			    
	arrParam(1) = "A_JNL_ITEM,b_trade_charge"		
	arrParam(2) = Trim(frm1.txtChargeType.Value)	
'	arrParam(3) = Trim(frm1.txtChargeTypeNm.Value)	
	arrParam(4) = "b_trade_charge.charge_cd=A_JNL_ITEM.JNL_CD And A_JNL_ITEM.JNL_TYPE=" & FilterVar("EC", "''", "S") & " and b_trade_charge.module_type=" & FilterVar("M", "''", "S") & " "
	arrParam(5) = "경비항목"						
	
    arrField(0) = "JNL_CD"					
    arrField(1) = "JNL_NM"					
    
    arrHeader(0) = "경비항목"			
    arrHeader(1) = "경비항목명"			
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtChargeType.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtChargeType.Value = arrRet(0)
		frm1.txtChargeTypeNm.Value = arrRet(1)	
		frm1.txtChargeType.focus	
		Set gActiveElement = document.activeElement	
	End If	
End Function
'==============================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "지급처"				
	arrParam(1) = "B_Biz_Partner"			
	arrParam(2) = Trim(frm1.txtBpCd.Value)	
'	arrParam(3) = Trim(frm1.txtBpNm.Value)	
	arrParam(4) = ""				
	arrParam(5) = "지급처"		
	
    arrField(0) = "BP_CD"			
    arrField(1) = "BP_NM"			
    
    arrHeader(0) = "지급처"		
    arrHeader(1) = "지급처명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
Function Openprocessstep()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "진행구분"				         
	arrParam(1) = "B_minor"					         
	arrParam(2) = Trim(frm1.txtprocessstep.value)	 
'	arrParam(3) = trim(frm1.txtprocessstepNm.value)	
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""				
	arrParam(5) = "진행구분"			
	
    arrField(0) = "minor_cd"					
    arrField(1) = "minor_nm"					
    
    arrHeader(0) = "진행구분"				
    arrHeader(1) = "진행구분명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtprocessstep.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtprocessstep.Value = arrRet(0)
		frm1.txtprocessstepNm.Value = arrRet(1)	
		frm1.txtprocessstep.focus	
		Set gActiveElement = document.activeElement	
	End If	

End Function
'==============================================================================================================================
Function OpenCostCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "비용집계처"				
	arrParam(1) = "B_COST_CENTER"				
	arrParam(2) = Trim(frm1.txtCostCd.Value)	
'	arrParam(3) = Trim(frm1.txtCostNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "비용집계처"				
	
    arrField(0) = "COST_CD"  					
    arrField(1) = "COST_NM"	    				
        
    arrHeader(0) = "비용집계처"				
    arrHeader(1) = "비용집계처명"			
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCostCd.focus	
		Set gActiveElement = document.activeElement			
		Exit Function
	Else
		frm1.txtCostCd.Value = arrRet(0)
		frm1.txtCostNm.Value = arrRet(1)
		frm1.txtCostCd.focus	
		Set gActiveElement = document.activeElement				
	End If	
End Function
'==============================================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	Dim iLoop
	Dim tmpPopUpR
	
	On Error Resume Next
	
	ReDim arrParam(parent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = gMethodText
	
	tmpPopUpR = GetPopUpR("A")
	
	For iLoop = 0 to parent.C_MaxSelList * 2 - 1 Step 2
      arrParam(iLoop + 0 ) = tmpPopUpR(iLoop / 2  , 0)
      arrParam(iLoop + 1 ) = tmpPopUpR(iLoop / 2  , 1)
    Next  
      
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(GetSQLSortFieldCD("A"),GetSQLSortFieldNm("A"),arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   Call SetPopUpR("A",arrRet,frm.vspdData)   
	   
       Call InitVariables
       Call InitSpreadSheet
       
   End If
   
End Function
'==============================================================================================================================
Sub openBasNo()
	Dim strChargeType
		
	strChargeType = Trim(frm1.txtProcessStep.value)
		
	If strChargeType <> "" Then
		Select Case UCase(strChargeType)
			Case "PO"							'Count Offer
				Call OpenBasNoPop("m3111pa1")
			Case "VL"							'수입 L/C
				Call OpenBasNoPop("m3211pa1")	
			Case "VA"							'수입 L/C Amend
				Call OpenBasNoPop("m3221pa1")			
			Case "VO"							'수입 Local L/C
				Call OpenBasNoPop("m3211pa2")
			Case "VF"							'수입 Local L/C Amend
				Call OpenBasNoPop("m3211pa2")		
			Case "VD"							'수입통관 
				Call OpenBasNoPop("m4211pa1")			
			Case "VB"							'수입선적 
				Call OpenBasNoPop("m5211pa1")		
			Case Else
				Call DisplayMsgBox("17A003","X" , "진행구분","X")
		End Select
		frm1.txtBasNo.focus	
	Else
		Call DisplayMsgBox("17A002","X" , "진행구분","X")
		frm1.txtProcessStep.focus	
	End If
	Set gActiveElement = document.activeElement
End Sub
'==============================================================================================================================
Function OpenBasNoPop(ByVal strPath)
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
		
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
		
	iCalledAspName = AskPRAspName(strPath)
		
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, strPath, "X")
		lgIsOpenPop = False
		Exit Function
	End If
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
		
	if UCase(Trim(frm1.txtProcessStep.value)) = "PO" then
			
		If strRet(0) = "" Then
			Call SetBasNo()
			Exit Function
		Else
			frm1.txtBasNo.value = strRet(0)
		End If
	else
		If strRet = "" Then
			Call SetBasNo()
			Exit Function
		Else
		   frm1.txtBasNo.value = strRet
		End If
	End if
	Call SetBasNo()
End Function
'==============================================================================================================================
Function SetBasNo()
	frm1.txtBasNo.focus	
	Set gActiveElement = document.activeElement
End Function	
'==============================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877				

	If Kubun = 1 Then						

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , IsCookieSplit			
		Call PgmJump(BIZ_PGM_JUMP_ID2)

	ElseIf Kubun = 0 Then							

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		Dim iniSep
	
		If Len(ReadCookie ("tBizArea")) Then
			frm1.txtBizArea.Value	=  ReadCookie ("tBizArea")
			WriteCookie "tBizArea",""
		Else
			frm1.txtBizArea.Value	=  arrVal(2)
		End If	
		
		frm1.txtBizAreaNm.value		=  arrVal(3)
		
		If Len(ReadCookie ("ChargeType")) Then
			frm1.txtChargeType.Value	=  ReadCookie ("ChargeType")
			WriteCookie "ChargeType",""
		Else
			frm1.txtChargeType.Value	=  arrVal(0)			
		End If
		
		frm1.txtChargeTypeNm.value	=  arrVal(1)
						
		If Len(ReadCookie ("BpCd")) Then
			frm1.txtBpCd.Value	=  ReadCookie ("BpCd")
			WriteCookie "BpCd",""
		Else
			frm1.txtBpCd.Value		=  arrVal(6)
		End If
		
		frm1.txtBpNm.value			=  arrVal(7)
		
		If arrVal(8) = "" or arrVal(8) = Null Then
			frm1.txtChargeFrDt.Text		=  ReadCookie ("ChargeFrDt")
			WriteCookie "ChargeFrDt",""
		Else
			frm1.txtChargeFrDt.Text		=  arrVal(8)			
		End If
				
		If arrVal(8) = "" or arrVal(8) = Null Then
			frm1.txtChargeToDt.Text		=  ReadCookie ("ChargeToDt")
			WriteCookie "ChargeToDt",""
		Else
			frm1.txtChargeToDt.Text		=  arrVal(8)			
		End If
				
		If Len(ReadCookie ("tCostCd")) Then
			frm1.txtCostCd.Value	=  ReadCookie ("tCostCd")
			WriteCookie "tCostCd",""
		Else
			frm1.txtCostCd.Value	=  arrVal(4)
		End If
		
		frm1.txtCostNm.Value	    =  arrVal(5)
						
		If Len(ReadCookie ("ProcessStep")) Then
			frm1.txtProcessStep.Value	=  ReadCookie ("ProcessStep")
			WriteCookie "ProcessStep",""
		Else
			frm1.txtProcessStep.Value	=  arrVal(9)
		End If
				
		frm1.txtProcessStepNm.Value	=  arrVal(10)
				
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 2 Then
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)		
		WriteCookie CookieSplit , IscookieSplit		
		If frm1.vspdData.ActiveRow > 0 Then			
			strTemp = ReadCookie(CookieSplit)
			If strTemp = "" Then Exit Function
			arrVal = Split(strTemp, parent.gRowSep)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow 			
			WriteCookie "Process_Step" , arrVal(0)
			WriteCookie "Po_No" , arrVal(1)
			WriteCookie "Pur_Grp" ,arrVal(2)			
			WriteCookie CookieSplit , ""			
		End if 					
				
		Call PgmJump(BIZ_PGM_JUMP_ID1)
	End If
		
End Function
'==============================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       

	Call InitVariables							
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")	
    Call CookiePage(0)
    
    frm1.txtBizArea.focus
    Set gActiveElement = document.activeElement
End Sub
'==============================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub
'==============================================================================================================================
Function OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"), gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function
'==============================================================================================================================
Sub txtChargeFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtChargeFrDt.Focus
    End If
End Sub
'==============================================================================================================================
Sub txtChargeToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtChargeToDt.Focus
    End If
End Sub
'==============================================================================================================================
Sub txtChargeFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Sub txtChargeToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'==============================================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
    End If
		If frm1.vspdData.MaxRows > 0 Then
			If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			End If
		End If
End Function
'==============================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Set gActiveSpdSheet = frm1.vspdData
    SetPopupMenuItemInf("00000000001")
	
	gMouseClickStatus = "SPC"
	
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
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    
	IsCookieSplit=""
    With frm1.vspddata
		.Row = Row
		.Col = GetKeyPos("A", 2)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", 5)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", 15)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
	End With
	
End Sub
'==============================================================================================================================
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
'==============================================================================================================================
Function FncQuery() 

    FncQuery = False                                        
    
    Err.Clear                                               
    
    With frm1
          If CompareDateByFormat(.txtChargeFrDt.text,.txtChargeToDt.text,.txtChargeFrDt.Alt,.txtChargeToDt.Alt, _
                   "970025",.txtChargeFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtChargeFrDt.text) <> "" And Trim(.txtChargeToDt.text) <> "" Then
			Call DisplayMsgBox("17a003","X","발생일자","X")	
			Exit Function
		End if   
	End with
	
	Call ggoOper.ClearField(Document, "2")	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData				
    Call InitVariables
    
    Call DbQuery											

    FncQuery = True											
	Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)  
    Set gActiveElement = document.activeElement                  
End Function
'==============================================================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                               
    If  LayerShowHide(1) = False Then
      	Exit Function
    End If

    
    With frm1
	    If lgIntFlgMode = parent.OPMD_UMODE Then
	        strVal = BIZ_PGM_ID & "?txtBizArea=" & Trim(.hdnBizArea.value)
	        strVal = strVal & "&txtChargeType=" & Trim(.hdnChargeType.value)
	        strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.value)
    	    strVal = strVal & "&txtChargeFrDt=" & Trim(.hdnChargeFrDt.value)
    	    strVal = strVal & "&txtChargeToDt=" & Trim(.hdnChargeToDt.value)
    	    strVal = strVal & "&txtCostCd=" & Trim(.hdnCostCd.value)    	
    	    strVal = strVal & "&txtProcessStep=" & Trim(.hdnProcessStep.value)
		    strVal = strVal & "&txtBasNo=" & Trim(.hdnBasNo.value)
    	    strVal = strVal & "&txtBasDocNo=" & Trim(.hdnBasDocNo.value)
        Else  
	        strVal = BIZ_PGM_ID & "?txtBizArea=" & Trim(.txtBizArea.value)
	        strVal = strVal & "&txtChargeType=" & Trim(.txtChargeType.value)
	        strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
    	    strVal = strVal & "&txtChargeFrDt=" & Trim(.txtChargeFrDt.Text)
    	    strVal = strVal & "&txtChargeToDt=" & Trim(.txtChargeToDt.Text)
    	    strVal = strVal & "&txtCostCd=" & Trim(.txtCostCd.value)    	
    	    strVal = strVal & "&txtProcessStep=" & Trim(.txtProcessStep.value)
		    strVal = strVal & "&txtBasNo=" & Trim(.txtBasNo.value)
    	    strVal = strVal & "&txtBasDocNo=" & Trim(.txtBasDocNo.value)
        End If    
            strVal = strVal & "&lgPageNo="		 & lgPageNo   
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

		
        Call RunMyBizASP(MyBizASP, strVal)								

    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")									

End Function
'==============================================================================================================================
Function DbQueryOk()													
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = parent.OPMD_UMODE
    
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtBizArea.focus
	End If
	Set gActiveElement = document.activeElement	
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경비상세</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
								    <TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="사업장" NAME="txtBizArea" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea() ">
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>경비항목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="경비항목" NAME="txtChargeType" SIZE=10 MAXLENGTH=20 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChargeType()">
														   <INPUT TYPE=TEXT NAME="txtChargeTypeNm" SIZE=20 tag="14"></TD>					   
								</TR>
								<TR>						   
									<TD CLASS="TD5" NOWRAP>지급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="지급처" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m6111qa2_fpDateTime2_txtChargeFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m6111qa2_fpDateTime2_txtChargeToDt.js'></script>
												</td>
											<tr>
										</table>
									</TD>
	                            </TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>비용집계처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="비용집계처" NAME="txtCostCd" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCd() ">
														   <INPUT TYPE=TEXT NAME="txtCostNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>진행구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="진행구분" NAME="txtProcessStep" SIZE=10 MAXLENGTH=5  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcessStep" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProcessStep()">
														   <INPUT TYPE=TEXT NAME="txtProcessStepNm" SIZE=20 tag="14"></TD>					   
								</TR>	
								<TR>
								    <TD CLASS="TD5" NOWRAP>발생근거관리번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발생근거관리번호" NAME="txtBasNo" SIZE=32  MAXLENGTH=35 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnbas_no" align=top TYPE="BUTTON" ONCLICK="vbscript:openBasNo()"></TD>
									<TD CLASS="TD5" NOWRAP>발생근거번호</TD>
                                     <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발생근거번호" NAME="txtBasDocNo" SIZE=34  MAXLENGTH=35 tag="11XXXU">
									<TD CLASS="TD6" NOWRAP></TD>
 								</TR>     													   
															
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m6111qa2_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
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
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(2)">경비등록</a></a></TD>
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
<INPUT TYPE=HIDDEN NAME="hdnBizArea" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnChargeToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCostCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProcessStep" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBasNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBasDocNo" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
