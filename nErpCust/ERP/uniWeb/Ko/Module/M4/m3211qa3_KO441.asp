<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3211qa3
'*  4. Program Name         : Local LC집계조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/13
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
<!-- '******************************************  1.1 Inc 선언   *******************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ====================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
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

Const BIZ_PGM_ID 		= "m3211qb3_KO441.asp"   
Const BIZ_PGM_JUMP_ID 	= "m3211qa4"                    
Const C_MaxKey          = 13				

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
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
	Set gActiveElement = document.activeElement
End Sub

'=======================================  LoadInfTB19029()  =====================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M3211QA3","G","A","V20030410", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'------------------------------------------  OpenBizArea()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_PLANT"	
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	
'	arrParam(3) = Trim(frm1.txtBizAreaNm.Value)	
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
	If frm1.txtPurGrpCd.className = "protected" Then Exit Function
	
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

	arrParam(0) = "수혜자"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBeneficiary.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "수혜자"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"				
    
    arrHeader(0) = "수혜자"			
    arrHeader(1) = "수혜자명"		
    
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
	arrParam(4) = "B_MINOR.MAJOR_CD=" & FilterVar("B9004", "''", "S") & " AND B_MINOR.MINOR_CD =B_CONFIGURATION.MINOR_CD AND B_CONFIGURATION.REFERENCE=" & FilterVar("L", "''", "S") & " "				
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

	arrParam(0) = "근거서류유형"				
	arrParam(1) = "B_MINOR"				
	arrParam(2) = Trim(frm1.txtOLcKind.Value)	
'	arrParam(3) = Trim(frm1.txtIncotermsNm.Value)	
	arrParam(4) = "MAJOR_CD=" & FilterVar("S9002", "''", "S") & ""							
	arrParam(5) = "근거서류유형"				
	
    arrField(0) = "minor_cd"  					
    arrField(1) = "minor_nm"	    				
        
    arrHeader(0) = "근거서류유형"				
    arrHeader(1) = "근거서류유형명"			
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtOLcKind.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtOLcKind.Value = arrRet(0)
		frm1.txtOLcKindNm.Value = arrRet(1)	
		frm1.txtOLcKind.focus	
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
		
		If Len(Trim(frm1.txtOLcKind.value)) Then
			WriteCookie "tOLcKind",Trim(frm1.txtOLcKind.value) 
		Else
			WriteCookie "tOLcKind",""
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
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")	
    Call InitComboBox()

End Sub
'========================================  Form_QueryUnload()  ======================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub
  
'========================================  FncSplitColumn()  ==================================
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
'========================================  vspdData_GotFocus()  ==================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'========================================  vspdData_DblClick()  ==================================
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
'========================================  vspdData_Click()  ==================================	
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

	If Row < 1 Then Exit Sub
	
	Const C_L_PlantCd		= 1
	Const C_L_PlantNm		= 2
	Const C_L_Beneficiary	= 3
	Const C_L_BpNm			= 4
	Const C_L_ItemCd		= 5
	Const C_L_ItemNm		= 6
	Const C_L_Spec			= 7
	Const C_L_Unit			= 8
	Const C_L_Currency		= 10
	
	IscookieSplit = ""	
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
'========================================  vspdData_TopLeftChange()  ==================================		
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
'========================================  vspdData_MouseDown()  ======================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub  
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'================================  FncQuery()  ==========================================
Function FncQuery() 

    FncQuery = False                                        
    
    Err.Clear                                               

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData				
    Call InitVariables 										
    
    with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","개설일","X")	
			Exit Function
		End if   
	End with

    Call DbQuery											

    FncQuery = True	
    Set gActiveElement = document.activeElement										

End Function
'================================  FncSave()  ==========================================	
Function FncSave()     
End Function
'================================  FncPrint()  ==========================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'================================  FncExcel()  ==========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'================================  FncFind()  ==========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False) 
    Set gActiveElement = document.activeElement                   
End Function
'================================  FncExit()  ==========================================
Function FncExit()
	
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'================================  DbQuery()  ==========================================
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
		strVal = strVal	& "&txtBeneficiary=" & Trim(.hdnBeneficiary.value)
		strVal = strVal	& "&txtFrDt=" & Trim(.hdnFrDt.value)
		strVal = strVal	& "&txtToDt=" & Trim(.hdnToDt.value)
		strVal = strVal	& "&txtPayMeth=" & Trim(.hdnPayMeth.value)		
		strVal = strVal	& "&txtOLcKind=" & Trim(.hdnOLcKind.value)

		Else
		
		strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal	& "&txtPurGrpCd=" &	Trim(.txtPurGrpCd.value)
		strVal = strVal	& "&txtBeneficiary="     &	Trim(.txtBeneficiary.value)
		strVal = strVal	& "&txtFrDt="	  & Trim(.txtFrDt.Text)
		strVal = strVal	& "&txtToDt="	  & Trim(.txtToDt.Text)
		strVal = strVal	& "&txtPayMeth="	  & Trim(.txtPayMeth.value)		
		strVal = strVal	& "&txtOLcKind=" & Trim(.txtOLcKind.value)

		End If
		
		strVal = strVal & "&lgPageNo="   & lgPageNo         
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

'================================  DbQueryOk()  ==========================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Local L/C집계</font></td>
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
									<TD CLASS="TD5" NOWRAP>수혜자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수혜자" NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>개설일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m3211qa3_fpDateTime2_txtFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m3211qa3_fpDateTime2_txtToDt.js'></script>
												</td>
											<tr>
										</table>
									</TD>
	                            </TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>결제방법</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="결제방법" NAME="txtPayMeth" SIZE=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayMeth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayMeth() ">
														   <INPUT TYPE=TEXT NAME="txtPayMethNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>근거서류유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="근거서류유형" NAME="txtOLcKind" SIZE=10 MAXLENGTH=5  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIncoterms()">
														   <INPUT TYPE=TEXT NAME="txtOLcKindNm" SIZE=20 tag="14"></TD>					   
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
									<script language =javascript src='./js/m3211qa3_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">LOCAL L/C상세</a></TD>
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
<INPUT TYPE=HIDDEN NAME="hdnOLcKind" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
