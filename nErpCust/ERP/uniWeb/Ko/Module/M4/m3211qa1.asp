<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3211qa1
'*  4. Program Name         : LC집계조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/12
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : park jin uk
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
                          
Dim lgIsOpenPop                         
Dim IscookieSplit 
Dim lgSaveRow                           


Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat) 

Const BIZ_PGM_ID 		= "m3211qb1.asp"   
Const BIZ_PGM_JUMP_ID 	= "m3211qa2"                     
Const C_MaxKey          = 11				

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
	lgIntFlgMode = Parent.OPMD_CMODE 
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub

'===========================================  LoadInfTB19029()  ==============================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M3211QA1","G","A","V20030410", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_PLANT"	
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
	arrParam(4) = ""							
	arrParam(5) = "공장"					

    arrField(0) = "PLANT_CD"					
    arrField(1) = "PLANT_NM"					
    
    arrHeader(0) = "공장"					
    arrHeader(1) = "공장명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++  OpenPurGrp() ++++++++++++++++++++++++++++++++++++++++
Function OpenPurGrp()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True

	arrParam(0) = "구매그룹"			    
	arrParam(1) = "B_Pur_Grp"		
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)	
	'	arrParam(3) = Trim(frm1.txtChargeTypeNm.Value)	
	arrParam(4) = ""
	arrParam(5) = "구매그룹"			
		
	arrField(0) = "Pur_Grp"			
	arrField(1) = "Pur_Grp_NM"		
	    
	arrHeader(0) = "구매그룹"				
	arrHeader(1) = "구매그룹명"				
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		

	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)	
		frm1.txtPurGrpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function	

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수출자"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBeneficiary.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "수출자"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"				
    
    arrHeader(0) = "수출자"			
    arrHeader(1) = "수출자명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBeneficiary.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBeneficiary.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBeneficiary.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPayMeth()  -------------------------------------------------
Function OpenPayMeth()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "결제방법"			
	arrParam(1) = "B_MINOR,B_CONFIGURATION"				
	arrParam(2) = Trim(frm1.txtPayMeth.value)
'	arrParam(3) = trim(frm1.txtPayMethNm.value)	
	arrParam(4) = "B_MINOR.MAJOR_CD=" & FilterVar("B9004", "''", "S") & " AND B_MINOR.MINOR_CD =B_CONFIGURATION.MINOR_CD AND B_CONFIGURATION.REFERENCE=" & FilterVar("M", "''", "S") & " "				
	arrParam(5) = "결제방법"			
	
    arrField(0) = "b_minor.minor_cd"			
    arrField(1) = "b_minor.minor_nm"			
    
    arrHeader(0) = "결제방법"		
    arrHeader(1) = "결제방법명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPayMeth.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtPayMeth.Value = arrRet(0)
		frm1.txtPayMethNm.Value = arrRet(1)	
		frm1.txtPayMeth.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function

'------------------------------------------  OpenIncoterms()  -------------------------------------------------
Function OpenIncoterms()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "가격조건"				
	arrParam(1) = "B_MINOR"				
	arrParam(2) = Trim(frm1.txtIncoterms.Value)	
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9006", "''", "S") & ""							
	arrParam(5) = "가격조건"				
	
    arrField(0) = "minor_cd"  					
    arrField(1) = "minor_nm"	    				
        
    arrHeader(0) = "가격조건"				
    arrHeader(1) = "가격조건명"			
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtIncoterms.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtIncoterms.Value = arrRet(0)
		frm1.txtIncotermsNm.Value = arrRet(1)	
		frm1.txtIncoterms.focus	
		Set gActiveElement = document.activeElement			
	End If	
End Function

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'------------------------------------  OpenGroupPopup()  ----------------------------------------------
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function
'==========================================   CookiePage()  ======================================
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877					

	If Kubun = 1 Then							

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)		
		WriteCookie CookieSplit , IsCookieSplit	
	
		If Len(Trim(frm1.txtPlantCd.value)) Then			
			WriteCookie "tPlantCd",Trim(frm1.txtPlantCd.value) 
		Else
			WriteCookie "tPlantCd",""
		End If
		
		If Len(Trim(frm1.txtPurGrpCd.value)) Then
			WriteCookie "tPurGrpCd",Trim(frm1.txtPurGrpCd.value) 
		Else
			WriteCookie "tPurGrpCd",""
		End If
		
		If Len(Trim(frm1.txtBeneficiary.value)) Then
			WriteCookie "tBeneficiary",Trim(frm1.txtBeneficiary.value) 
		Else
			WriteCookie "tBeneficiary",""
		End If
		
		If Len(Trim(frm1.txtFrDt.text)) Then
			WriteCookie "tFrDt",Trim(frm1.txtFrDt.text) 
		Else
			WriteCookie "tFrDt",""
		End If
		
		If Len(Trim(frm1.txtToDt.text)) Then
			WriteCookie "tToDt",Trim(frm1.txtToDt.text) 
		Else
			WriteCookie "tToDt",""
		End If
				
		If Len(Trim(frm1.txtPayMeth.value)) Then
			WriteCookie "tPayMeth",Trim(frm1.txtPayMeth.value) 
		Else
			WriteCookie "tPayMeth",""
		End If
		
		If Len(Trim(frm1.txtIncoterms.value)) Then
			WriteCookie "tIncoterms",Trim(frm1.txtIncoterms.value) 
		Else
			WriteCookie "tIncoterms",""
		End If
		
				
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then						

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		'If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF
	
End Function
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	
    Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
    
	Call InitVariables							
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")	
    Call InitComboBox()

End Sub
'==========================================  Form_QueryUnload()  ======================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub
'==========================================  vspdData_MouseDown()  ======================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==========================================  FncSplitColumn()  ======================================
Function FncSplitColumn()
    
   If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function
'==========================================================================================
'   Event Name : txtFrDt
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtFrDt.focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt
'==========================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtToDt.focus
	End If
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'==========================================================================================
'   Event Name : vspdData_DblClick
'==========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
      Exit Function
    End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
	
'======================================================================================================
'   Event Name : vspdData_Click
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("00000000001")

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If

	Const C_L_PlantCd		= 1
	Const C_L_PlantNm		= 2
	Const C_L_Beneficiary	= 3
	Const C_L_BpNm			= 4
	Const C_L_ItemCd		= 5
	Const C_L_ItemNm		= 6
	Const C_L_Spec			= 7
	Const C_L_Unit			= 8
	Const C_L_Currency		= 9


	If Row < 1 Then Exit Sub
	
	IscookieSplit = ""	
	'====
	With frm1.vspddata
		.Row = Row
		
		.Col = GetKeyPos("A", C_L_PlantCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PlantNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_Beneficiary)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_BpNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_ItemCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_ItemNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_Spec)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
		
		.Col = GetKeyPos("A", C_L_Unit)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
		
		.Col = GetKeyPos("A", C_L_Currency)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
	
	End With

End Sub
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'=======================================================================================================
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
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'==============================  FncQuery()  ================================================
Function FncQuery() 

    FncQuery = False                                        
    
    Err.Clear                                               

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    					
    Call InitVariables 										
    
    With frm1
		If (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","개설일","X")	
			Exit Function
		End If   
	End With

    Call DbQuery											

    FncQuery = True	
    Set gActiveElement = document.activeElement										

End Function

'==============================  FncSave()  ================================================
Function FncSave()     
End Function
'==============================  FncPrint()  ================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'==============================  FncExcel()  ================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'==============================  FncFind()  ================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)   
    Set gActiveElement = document.activeElement                 
End Function
'==============================  FncExit()  ================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                               
		if LayerShowHide(1) =false then
		    exit Function
		end if
    
    With frm1
	  	
		If lgIntFlgMode = Parent.OPMD_UMODE Then	

		strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.hdnPlantCd.value)
		strVal = strVal	& "&txtPurGrpCd=" &	Trim(.hdnPurGrpCd.value)
		strVal = strVal	& "&txtBeneficiary="     &	Trim(.hdnBeneficiary.value)
		strVal = strVal	& "&txtFrDt="	  & Trim(.hdnFrDt.value)
		strVal = strVal	& "&txtToDt="	  & Trim(.hdnToDt.value)
		strVal = strVal	& "&txtPayMeth="	  & Trim(.hdnPayMeth.value)		
		strVal = strVal	& "&txtIncoterms=" & Trim(.hdnIncoterms.value)
        strVal = strVal & "&lgPageNo="   & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D) 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		Else
		
		strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal	& "&txtPurGrpCd=" &	Trim(.txtPurGrpCd.value)
		strVal = strVal	& "&txtBeneficiary="     &	Trim(.txtBeneficiary.value)
		strVal = strVal	& "&txtFrDt="	  & Trim(.txtFrDt.Text)
		strVal = strVal	& "&txtToDt="	  & Trim(.txtToDt.Text)
		strVal = strVal	& "&txtPayMeth="	  & Trim(.txtPayMeth.value)		
		strVal = strVal	& "&txtIncoterms=" & Trim(.txtIncoterms.value)
        strVal = strVal & "&lgPageNo="   & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		End If

        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()												

    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE 

    If Len(Trim(frm1.txtPurGrpNm.value)) Then
			WriteCookie "PurGrpCdNm",Trim(frm1.txtPurGrpNm.value) 
	Else
			WriteCookie "PurGrpCdNm",""
	End If
	
	If Len(Trim(frm1.txtPayMethNm.value)) Then
			WriteCookie "PayMethNm",Trim(frm1.txtPayMethNm.value) 
	Else
			WriteCookie "PayMethNm",""
	End If
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtPlantCd.focus
	End If				

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C집계</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right>&nbsp;<!--<button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenGroupPopup()">집계순서</button>--></td>
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
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="BtnPlantPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=20 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>					   
								</TR>
								<TR>						   
									<TD CLASS="TD5" NOWRAP>수출자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수출자" NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>개설일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m3211qa1_fpDateTime2_txtFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m3211qa1_fpDateTime2_txtToDt.js'></script>
												</td>
											<tr>
										</table>
									</TD>
	                            </TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>결제방법</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="결제방법" NAME="txtPayMeth" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayMeth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayMeth() ">
														   <INPUT TYPE=TEXT NAME="txtPayMethNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>가격조건</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="가격조건" NAME="txtIncoterms" SIZE=10 MAXLENGTH=5  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIncoterms()">
														   <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 tag="14"></TD>					   
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
									<script language =javascript src='./js/m3211qa1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">L/C상세</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMeth" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncoterms" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
