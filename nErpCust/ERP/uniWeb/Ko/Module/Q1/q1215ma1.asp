<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1215MA1
'*  4. Program Name         : 선별형검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_QRY_ID	= "q1215mb1.asp"				'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "q1215mb2.asp"				'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID	= "q1215mb3.asp"				'☆: 비지니스 로직 ASP명 

Const BIZ_PGM_JUMP_ID = "q1211ma1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop						' Popup

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size

    IsOpenPop = False							'☆: 사용자 변수 초기화 
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	frm1.cboInspClassCd.value = "R"
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
	End If	
	
	If ReadCookie("txtItemNm") <> "" Then
		frm1.txtItemNm.Value = ReadCookie("txtItemNm")
	End If	
	
	If ReadCookie("txtInspClassCd") <> "" Then
		frm1.cboInspClassCd.Value = ReadCookie("txtInspClassCd")
	End If	
		
	If ReadCookie("txtInspItemCd") <> "" Then
		frm1.txtInspItemCd.Value = ReadCookie("txtInspItemCd")
	End If	
		
	If ReadCookie("txtInspItemNm") <> "" Then
		frm1.txtInspItemNm.Value = ReadCookie("txtInspItemNm")
	End If	
	
	If ReadCookie("txtInspMthdCd") <> "" Then
		frm1.txtInspMthdCd.Value = ReadCookie("txtInspMthdCd")
	End If	
		
	If ReadCookie("txtInspMthdNm") <> "" Then
		frm1.txtInspMthdNm.Value = ReadCookie("txtInspMthdNm")
	End If
	
	If ReadCookie("txtInspClassCd") = "P" Then
		If ReadCookie("txtRoutNo") <> "" Then
			frm1.txtRoutNo.Value = ReadCookie("txtRoutNo")
		End If
		
		If ReadCookie("txtRoutNoDesc") <> "" Then
			frm1.txtRoutNoDesc.Value = ReadCookie("txtRoutNoDesc")
		End If
		
		If ReadCookie("txtOprNo") <> "" Then
			frm1.txtOprNo.Value = ReadCookie("txtOprNo")
		End If
		
		If ReadCookie("txtOprNoDesc") <> "" Then
			frm1.txtOprNoDesc.Value = ReadCookie("txtOprNoDesc")
		End If
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtInspClassCd", ""
	WriteCookie "txtInspItemCd", ""
	WriteCookie "txtInspItemNm", ""
	WriteCookie "txtInspMthdCd", ""
	WriteCookie "txtInspMthdNm", ""
	WriteCookie "txtRoutNo", ""
	WriteCookie "txtRoutNoDesc", ""
	WriteCookie "txtOprNo", ""
	WriteCookie "txtOprNoDesc", ""	
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Err.Clear
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboInspClassCd ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0019", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLotQualityIndex ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	Dim TmplgF0
	Dim TmplgF1
	Dim i

	TmplgF0 = split(lgF0,Chr(11))
	TmplgF1 = split(lgF1,Chr(11))	
	lgF0 = ""
	lgF1 = ""
	
	For i = 0 To UBound(TmplgF0) - 1
        lgF0 = lgF0 & uniConvNumAToB(TmplgF0(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next

	For i = 0 To UBound(TmplgF1) - 1
        lgF1 = lgF1 & uniConvNumAToB(TmplgF1(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
    Call SetCombo2(frm1.cboAOQL ,lgF0  ,lgF1  ,Chr(11))
    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0021", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	
	TmplgF0 = split(lgF0,Chr(11))
	TmplgF1 = split(lgF1,Chr(11))	
	lgF0 = ""
	lgF1 = ""
	
	For i = 0 To UBound(TmplgF0) - 1
        lgF0 = lgF0 & uniConvNumAToB(TmplgF0(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
	For i = 0 To UBound(TmplgF1) - 1
        lgF1 = lgF1 & uniConvNumAToB(TmplgF1(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
    Call SetCombo2(frm1.cboLTPD ,lgF0  ,lgF1  ,Chr(11))
End Sub

'=======================================================================================================
'   Event Name : LockAOQLLTPD()
'   Event Desc : 
'=======================================================================================================
Sub LockAOQLLTPD(Byval vLotQualityIndex)
	With frm1
		Select Case vLotQualityIndex
			Case "A"
				Call ggoOper.SetReqAttr(.cboAOQL, "N")
				Call ggoOper.SetReqAttr(.cboLTPD, "Q")
			Case "L"
				Call ggoOper.SetReqAttr(.cboLTPD, "N")
				Call ggoOper.SetReqAttr(.cboAOQL, "Q")
			Case Else
				Call ggoOper.SetReqAttr(.cboLTPD, "Q")
				Call ggoOper.SetReqAttr(.cboAOQL, "Q")
		End Select 
	End With
End Sub

'=======================================================================================================
'   Event Name : ClearDataAsLotQualityIndex()
'   Event Desc : 
'=======================================================================================================
Sub ClearDataAsLotQualityIndex(Byval vLotQualityIndex)
		With frm1
		Select Case vLotQualityIndex
			Case "A"
				.cboLTPD.value=""
			Case "L"
				.cboAOQL.value=""
			Case Else
				.cboAOQL.value=""
				.cboLTPD.value=""
		End Select 
	End With
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "공장코드"		
	arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

 '------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	frm1.txtItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
		frm1.txtItemCd.Focus		
	End If	

	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

 '------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : Inspection Item By Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	OpenInspItem = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		Exit Function
	End If
	'검사분류가 있는 지 체크 
	If Trim(frm1.cboInspClassCd.Value) = "" then 
		Call DisplayMsgBox("229915", "X", "X", "X") 		'검사분류정보가 필요합니다 
		Exit Function
	End If
	'품목코드가 있는 지 체크 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916", "X", "X", "X") 		'품목정보가 필요합니다 
		Exit Function
	End If
	
	If Trim(frm1.cboInspClassCd.Value) = "P" then 
		'RoutNo가 있는 지 체크 
		If Trim(frm1.txtRoutNo.Value) = "" then 
			Call DisplayMsgBox("220735", "X", "X", "X") 		'라우팅정보가 필요합니다 
			frm1.txtRoutNo.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
		
		'OprNo가 있는 지 체크 
		If Trim(frm1.txtOprNo.Value) = "" then 
			Call DisplayMsgBox("220736", "X", "X", "X") 		'공정정보가 필요합니다 
			frm1.txtOprNo.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
	IsOpenPop = True
	
	With frm1
		Param1 = Trim(.txtPlantCd.Value)
		Param2 = Trim(.txtPlantNm.Value)
		Param3 = Trim(.txtItemCd.Value)
		Param4 = Trim(.txtItemNm.Value)
		Param5 = Trim(.cboInspClassCd.Value)
		Param6 = Trim(.cboInspClassCd.Options(.cboInspClassCd.SelectedIndex).Text)
		Param7 = Trim(.txtRoutNo.Value)
		Param8 = Trim(.txtRoutNoDesc.Value)
		Param9 = Trim(.txtOprNo.Value)
		Param10 = Trim(.txtInspItemCd.value)
		Param11 = ""
		Param12 = "0200"	'선별형 
	End With
	
	iCalledAspName = AskPRAspName("q1211pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9, Param10, Param11, Param12), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtInspItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspItemCd.Value = arrRet(1)
		frm1.txtInspItemNm.Value = arrRet(2)	
		frm1.txtInspMthdCd.Value = arrRet(3)
		frm1.txtInspMthdNm.Value = arrRet(4)
		frm1.txtInspItemCd.Focus
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspItem = true
End Function


'====================  OpenRoutNo  ======================================
' Function Name : OpenRoutNo
' Function Desc : OpenRoutNo Reference Popup
'==========================================================================
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") 	
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ROUT_NO"							' Field명(0)
    arrField(1) = "DESCRIPTION"						' Field명(1)
    arrField(2) = "BOM_NO"							' Field명(1)
    arrField(3) = "MAJOR_FLG"						' Field명(1)
   
    arrHeader(0) = "라우팅"						' Header명(0)
    arrHeader(1) = "라우팅명"					' Header명(1)
    arrHeader(2) = "BOM Type"					' Header명(1)
    arrHeader(3) = "주라우팅"					' Header명(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value		= arrRet(0)		
		frm1.txtRoutNoDesc.Value		= arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function



'**************************** Function OpenOprNo() ***********************************8
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "품목","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	
	
	If frm1.txtRoutNo.value= "" Then
		Call DisplayMsgBox("971012","X", "라우팅","X")
		frm1.txtRoutNo.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
				  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "공정"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "공정"		
	arrHeader(1) = "작업장"	
	arrHeader(2) = "작업장명"
	arrHeader(3) = "공정작업명"
	arrHeader(4) = "사내구분"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "검사여부"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtOprNo.focus
		Exit Function
	Else
		frm1.txtOprNo.Value = arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(3)
	End If	
	
End Function

'=============================================  2.5.2 LoadInspStand()  ======================================
'=	Event Name : LoadInspStand
'=	Event Desc :
'========================================================================================================
Function LoadInspStand()
	Dim intRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		'공장코드/명/품목코드/명/검사분류코드 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtItemCd", Trim(.txtItemCd.value)
		WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
			WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		End if
		
	End With
	PgmJump(BIZ_PGM_JUMP_ID)

End Function


'============================================= EnableField()  ======================================
'=	Event Name : EnableField
'=	Event Desc :
'========================================================================================================
Sub EnableField(Byval strInspClass)
	If	strInspClass = "P" Then
		Process.style.display	= ""
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "N")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "N")
	Else	
		Process.style.display	= "none"
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
	End if
End Sub



'============================================= cboInspClassCd_onchange()  ======================================
'=	Event Name : cboInspClassCd_onchange()
'=	Event Desc :
'========================================================================================================
Sub cboInspClassCd_onchange()
	Call EnableField(frm1.cboInspClassCd.value)
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	
	Call AppendNumberPlace("6", "2", "2")
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call InitVariables                                                      				'⊙: Initializes local global variables
	Call InitComboBox
	Call SetToolBar("11101000000011")
	Call SetDefaultVal
	Call EnableField(frm1.cboInspClassCd.value)
	
	frm1.cboLotQualityIndex.value = "A"
	Call LockAOQLLTPD("A")
	frm1.cboAOQL.value = ""
	frm1.cboLTPD.value = ""
	
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
	
	lgBlnFlgChgValue = False
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : cboLotQualityIndex_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboLotQualityIndex_onchange()
	lgBlnFlgChgValue = True
	Call LockAOQLLTPD(frm1.cboLotQualityIndex.value)
	Call ClearDataAsLotQualityIndex(frm1.cboLotQualityIndex.value)
End Sub

'=======================================================================================================
'   Event Name : cboLTPD_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboLTPD_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboAOQL_onchange()
'   Event Desc : change flag setting
'=======================================================================================================
Sub cboAOQL_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPBar_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPBar_Change()	
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtPBar_KeyPress()
'   Event Desc : change flag setting
'=======================================================================================================
Function  txtPBar_KeyPress(KeyAscii)
	txtPBar_KeyPress = false
	If KeyAscii = 13 Then
		Call MainSave()
	End If
	txtPBar_KeyPress = true
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    
    	Dim IntRetCD 
    
    	FncQuery = False                                                        						'⊙: Processing is NG
    
    	Err.Clear                                                               						'☜: Protect system from crashing

    	'-----------------------
    	'Check previous data area
    	'-----------------------
    	If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
    	End If
    
    	'-----------------------
    	'Erase contents area
    	'-----------------------
    	Call ggoOper.ClearField(Document, "2")							'⊙: Clear Contents  Field
    	Call InitVariables										'⊙: Initializes local global variables
    	
    	frm1.cboLotQualityIndex.value = "A"
		Call LockAOQLLTPD("A")
		Call ClearDataAsLotQualityIndex("A")
    
    	'-----------------------
    	'Check condition area
    	'-----------------------
    	If Not chkField(Document, "1") Then							'⊙: This function check indispensable field
       		Exit Function
    	End If
    
    	'-----------------------
    	'Query function call area
	   	'-----------------------
		If DbQuery = False then
			Exit Function
		End If										'☜: Query db data
       
    	FncQuery = True										'⊙: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
	FncNew = False                                                          '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	Call SetDefaultVal
	frm1.cboLotQualityIndex.value = "A"
	Call LockAOQLLTPD("A")
	Call ClearDataAsLotQualityIndex("A")
	Call InitVariables                                                      '⊙: Initializes local global variables
	Call SetToolBar("11101000000011")
	Call EnableField(frm1.cboInspClassCd.value)
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If
	FncNew = True                                                            						'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    
    Dim IntRetCD 
    
    FncDelete = False                                                      						'⊙: Processing is NG
    
    Err.Clear                                                               						'☜: Protect system from crashing
    'On Error Resume Next                                                	
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    	'-----------------------
    	'Delete function call area
    	'-----------------------
    	If DbDelete = False Then   
    		Exit Function                                                        					'☜:
    	End If
    	
    	'-----------------------
    	'Erase condition area
    	'Erase contents area
    	'-----------------------
    	Call ggoOper.ClearField(Document, "A")                                         				'⊙: Clear Contents  Field
		
		frm1.cboLotQualityIndex.value = "A"
		Call LockAOQLLTPD("A")
		Call ClearDataAsLotQualityIndex("A")
    	
    	FncDelete = True                                                        						'⊙: Processing is OK
    	    	
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    		
    	Dim IntRetCD 
    
    	FncSave = False                                                         						'⊙: Processing is NG
    
    	Err.Clear                                                               						'☜: Protect system from crashing
    	On Error Resume Next                                                    
		
    	'-----------------------
    	'Precheck area
    	'-----------------------
    	If lgBlnFlgChgValue = False Then
        	IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        	Exit Function
    	End If
		
    	'-----------------------
    	'Check content area
    	'-----------------------
    	If Not chkField(Document, "A") Then   			'⊙: Check contents area
			Exit Function
		End If
    	
    	'-----------------------
    	'Save function call area
    	'-----------------------
		If DbSave = False then	
			Exit Function
		End If			                                                  			'☜: Save db data
    
    	FncSave = True                                                          						'⊙: Processing is OK
    	
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE														'⊙: Indicates that current mode is Crate mode
	lgBlnFlgChgValue = True
	Call ggoOper.ClearField(Document, "1")                                      					'⊙: Clear Condition Field
	Call ggoOper.LockField(Document, "N")												'⊙: This function lock the suitable field
	Call cboLotQualityIndex_onchange
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
	FncPaste = false
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

	FncCancel = false
	'On Error Resume Next                                     					            		'☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	FncInsertRow = false
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	FncDeleteRow = false

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	FncPrev = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgPrevNo = "" Then
	 	Call DisplayMsgBox("900011", "X", "X", "X")  '☜ 바뀐부분 
	 	Exit Function
	End If
	
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd.value)	'☆: 조회 조건 데이타 
	strVal = strVal & "&txtYr=" & frm1.txtYr1.Text
	strVal = strVal & "&cboInspClassCd=" & lgPrevNo									'☆: 조회 조건 데이타 
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncPrev = true
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false
	Dim strVal
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")  '☜ 바뀐부분 
		Exit Function
	End If

	
	strVal = BIZ_PGM_QRY_ID & "?txtPlantCd=" & Trim(frm1.txtPlantCd1.value)	'☆: 조회 조건 데이타 
	strVal = strVal & "&txtYr=" & frm1.txtYr1.Text
	strVal = strVal & "&cboInspClassCd=" & lgNextNo	
	
	Call RunMyBizASP(MyBizASP, strVal)
	FncNext = true
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = false
    	'On Error Resume Next                                                    						'☜: Protect system from crashing												'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	Call parent.FncFind(Parent.C_SINGLE, False)     
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
	FncScreenSave = false
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
	FncScreenRestore = false
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	DbQuery = False
	
	Err.Clear                                                               						'☜: Protect system from crashing
	Call LayerShowHide(1)
	
	With frm1	
		strVal = BIZ_PGM_QRY_ID		& "?txtMode="			& Parent.UID_M0001 _
									& "&txtPlantCd="		& Trim(.txtPlantCd.Value) _
									& "&txtItemCd="			& Trim(.txtItemCd.value) _
									& "&txtInspItemCd="		& Trim(.txtInspItemCd.value) _
									& "&cboInspClassCd="	& Trim(.cboInspClassCd.Value) _
									& "&txtRoutNo="			& Trim(.txtRoutNo.value) _
									& "&txtOprNo="			& Trim(.txtOprNo.value)
		
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
		
		DbQuery = True                                                          				'⊙: Processing is NG
	End With
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	DbQueryOk = false
	'-----------------------
	'Reset variables area
	'-----------------------
	Call SetToolBar("11111000001111")
	Call ggoOper.LockField(Document, "Q")		'⊙: This function lock the suitable field
	Call EnableField(frm1.cboInspClassCd.value)
	Call LockAOQLLTPD(frm1.cboLotQualityIndex.value)
	
	lgIntFlgMode = Parent.OPMD_UMODE			'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	
	Call LayerShowHide(1)
	
	DbSave = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value			= Parent.UID_M0002
		.txtUpdtUserId.value	= Parent.gUsrID
		.txtInsrtUserId.value	= Parent.gUsrID
		.txtFlgMode.value		= lgIntFlgMode
				
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
	DbSave = True
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()									'☆: 저장 성공후 실행 로직 
	DbSaveOk = false
	Call InitVariables
	Call MainQuery()
	DbSaveOk = true
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Err.Clear                                                               '☜: Protect system from crashing
    	
   	Call LayerShowHide(1)
    	
	DbDelete = False														'⊙: Processing is NG
	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID		& "?txtMode="			& Parent.UID_M0003 _
								& "&txtPlantCd="		& Trim(frm1.txtPlantCd.value) _
								& "&txtItemCd="			& Trim(frm1.txtItemCd.value) _
								& "&cboInspClassCd="	& Trim(frm1.cboInspClassCd.value) _
								& "&txtInspItemCd="		& Trim(frm1.txtInspItemCd.value)
	If Trim(frm1.cboInspClassCd.value) = "P" Then
		strVal = strVal & "&txtRoutNo="	& Trim(.txtRoutNo.value) _
						& "&txtOprNo="		& Trim(.txtOprNo.value)
	End if	
							
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()									'☆: 삭제 성공후 실행 로직 
	DbDeleteOk = false
	lgBlnFlgChgValue = False
	Call MainNew()
	DbDeleteOk = true
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>선별형 검사조건</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
		        							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        									<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="12"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
	        						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
									<INPUT TYPE=TEXT NAME="txtItemNm" SIZE="20" MAXLENGTH="20" tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR ID="Process">
					      			<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="12XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사항목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="검사항목" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
										<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
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
					<TD HEIGHT=30 WIDTH=100%>
						<TABLE CLASS="TB2" WIDTH="100%" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td CLASS="TD5" NOWPAP HEIGHT=5></td>
								<td CLASS="TD656" NOWPAP HEIGHT=5></td>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>검사방식</TD>
								<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtInspMthdCd" SIZE="10" MAXLENGTH="4" ALT="검사방식" tag="14">
								<INPUT TYPE=TEXT NAME="txtInspMthdNm" SIZE="40" MAXLENGTH="40" tag="14" ></TD>
							</TR>
							<TR>
								<td CLASS="TD5" NOWPAP HEIGHT=5></td>
								<td CLASS="TD656" NOWPAP HEIGHT=5></td>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=* WIDTH=100% valign=top>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>로트품질지표</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLotQualityIndex" ALT="로트품질지표" STYLE="WIDTH: 150px" tag="22"></SELECT></TD>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>AOQL</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboAOQL"  ALT="AOQL" STYLE="WIDTH: 80px" tag="22"></SELECT>&nbsp;%</TD>
									<TD CLASS="TD5" NOWRAP>LTPD</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLTPD"  ALT="LTPD" STYLE="WIDTH: 80px" tag="22"></SELECT>&nbsp;%</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공정평균불량률</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1215ma1_fpDoubleSingle1_txtPBar.js'></script>&nbsp;%
									</TD>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
								</TR>
								<TR>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
									<td CLASS="TD5" NOWPAP HEIGHT=5></td>
									<td CLASS="TD6" NOWPAP HEIGHT=5></td>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
								
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
      	<TD WIDTH="100%">
      		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
        		<TR>
        			<TD WIDTH=10>&nbsp;</td>
        			<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspStand">검사기준</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
       			</TR>
      		</TABLE>
      	</TD>
    </TR>
    	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspItemCd" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

