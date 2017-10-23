<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211ma1.asp																*
'*  4. Program Name         : 통관등록																	*
'*  5. Program Desc         : 통관등록																	*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
%>
<%%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<Script Language="VBS">
Option Explicit				
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)


	Const BIZ_PGM_ID				= "s4211mb1.asp"			
	Const BIZ_PGM_SOQRY_ID			= "s4211mb2.asp"
	Const BIZ_PGM_LCQRY_ID			= "s4211mb3.asp"	
	Const EXCC_DETAIL_ENTRY_ID		= "s4212ma1_ko441"		
	Const EXPORT_CHARGE_ENTRY_ID	= "s6111ma1_ko441"	
	Const EXCC_ASSIGN_ENTRY_ID ="s4214ma1"		'☆: 이동할 ASP명 : container 배정
	Const EXCC_ADDITEM_ENTRY_ID = "s4214ma1_ko441"
	Const TAB1 = 1
	Const TAB2 = 2
	Const TAB3 = 3

	Const gstrEDTypeMajor = "S9012"
	Const gstrPayTermsMajor = "B9004"
	Const gstrIncoTermsMajor = "B9006"
	Const gstrReturnOfficeMajor = "S9013"
	Const gstrTransFormMajor = "S9010"
	Const gstrPackingTypeMajor = "B9007"
	Const gstrCustomsMajor = "S9013"
	Const gstrExportTypeMajor = "S9009"
	Const gstrOriginMajor =	"B9094"	
	Const gstrEpTypesMajor = "S9008"
	Const gstrTransMethMajor = "S9011"
	Dim gSelframeFlg					
	Dim gblnWinEvent					

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE							
	lgBlnFlgChgValue = False							
	lgIntGrpCount = 0									
		
	gblnWinEvent = False
End Function
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtIVDt.text = EndDate
	frm1.txtGrossWeight.text = Parent.UNIFormatNumber(0, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
	frm1.txtDocAmt.text = Parent.UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtXchRate.text = Parent.UNIFormatNumber(0, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	frm1.txtLocAmt.text = Parent.UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtFOBDocAmt.text = Parent.UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtUSDXchRate.text = Parent.UNIFormatNumber(0, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	frm1.txtFOBLocAmt.text = Parent.UNIFormatNumber(0, ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	frm1.txtLocCCCurrency.value = Trim(Parent.gCurrency)
	frm1.txtLocFOBCurrency.value = Trim(Parent.gCurrency)
	
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGroup, "Q") 
        	frm1.txtSalesGroup.value = lgSGCd
	End If
	
	lgBlnFlgChgValue = False
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
		
	gSelframeFlg = TAB2
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenExCCNoPop()
	Dim iCalledAspName
	Dim IntRetCD
	Dim strRet
		
		
	If gblnWinEvent = True Or UCase(frm1.txtCCNo.className) = "PROTECTED" Then Exit Function
		
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("S4211PA1_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4211PA1_ko441", "X")
		'lblnWinEvent = False
		gblnWinEvent = False
		Exit Function
	End If
		
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetExCCNo(strRet)
	End If	
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSORef()
	Dim iCalledAspName
	Dim IntRetCD
	Dim strRet
				
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call Parent.DisplayMsgBox("200005", "x", "x", "x")
		Exit Function
	End If
		
	If gblnWinEvent = True Then Exit Function			
	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("S3111RA6_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3111RA6_ko441", "X")			
		gblnWinEvent = False
		Exit Function
	End If
				
				
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
				
	If strRet = "" Then
		Exit Function
	Else
		Call SetSORef(strRet)
	End If	
End Function	
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenLCRef()
	Dim iCalledAspName
	Dim IntRetCD
	Dim strRet

	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call Parent.DisplayMsgBox("200005", "x", "x", "x")
		Exit Function
	End If
		
	If gblnWinEvent = True Then Exit Function			
'		gblnWinEvent = True
				
	iCalledAspName = AskPRAspName("S3211RA6_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3211RA6_ko441", "X")			
'			gblnWinEvent = False
		Exit Function
	End If
				
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	If strRet(0) <> "" Then
		Call SetLCRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenDNRef()
	Dim iCalledAspName
	Dim IntRetCD
	Dim strRet
		
	If lgIntFlgMode = Parent.OPMD_UMODE Then 
		Call Parent.DisplayMsgBox("200005", "x", "x", "x")
		Exit Function
	End If
		
	If gblnWinEvent = True Then Exit Function			
'		gblnWinEvent = True
				
	iCalledAspName = AskPRAspName("M4111RA6_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4111RA6_ko441", "X")			
'			gblnWinEvent = False
		Exit Function
	End If
				
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetDNRef(strRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenPort(strMinorCD, strMinorNM, strPopNm, iwhere)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
		arrParam(0) = strPopNm							
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(strMinorCD)					
		arrParam(3) = ""								
		arrParam(4) = "MAJOR_CD = " & FilterVar("B9092", "''", "S") & ""				
		arrParam(5) = strPopNm							
		
		arrField(0) = "Minor_CD"						
		arrField(1) = "Minor_NM"						
	    
		arrHeader(0) = strPopNm							
		arrHeader(1) = strPopNm & "_명"				

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOpenPort(iwhere, arrRet)
	End If	
			
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCountry(strCntryCD, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "국가"			
	arrParam(1) = "B_COUNTRY"				
	arrParam(2) = Trim(strCntryCD)			
	arrParam(3) = ""						
	arrParam(4) = ""						
	arrParam(5) = "국가"				

	arrField(0) = "COUNTRY_CD"				
	arrField(1) = "COUNTRY_NM"				

	arrHeader(0) = "국가"				
	arrHeader(1) = "국가명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCountry(strPopPos, arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If frm1.txtSalesGroup.className = "protected" Then Exit Function
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "영업그룹"					
	arrParam(1) = "B_SALES_GRP"						
	arrParam(2) = Trim(frm1.txtSalesGroup.value)	
	arrParam(3) = ""								
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "영업그룹"					

	arrField(0) = "SALES_GRP"						
	arrField(1) = "SALES_GRP_NM"					

	arrHeader(0) = "영업그룹"					
	arrHeader(1) = "영업그룹명"			
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenBizPartner(strBizPartnerCD, strBizPartnerNM, strPopPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos							
	arrParam(1) = "B_BIZ_PARTNER"					
	arrParam(2) = Trim(strBizPartnerCD)				
	arrParam(3) = ""										
	arrParam(4) = "bp_type IN ( " & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ", " & FilterVar("S", "''", "S") & " ) AND usage_flag = " & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = strPopPos							

	arrField(0) = "BP_CD"							
	arrField(1) = "BP_NM"							

	arrHeader(0) = strPopPos					
	arrHeader(1) = strPopPos & "_명"			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(strPopPos, arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "중량단위"				
	arrParam(1) = "B_UNIT_OF_MEASURE"			
	arrParam(2) = Trim(frm1.txtWeightUnit.value)
	arrParam(3) = ""							
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & ""				
	arrParam(5) = "중량단위"				

	arrField(0) = "UNIT"						
	arrField(1) = "UNIT_NM"						

	arrHeader(0) = "중량단위"				
	arrHeader(1) = "중량단위명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUnit(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos						
	arrParam(1) = "B_Minor"						
	arrParam(2) = Trim(strMinorCD)				
	arrParam(3) = ""							
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""	
	arrParam(5) = strPopPos							

	arrField(0) = "Minor_CD"						
	arrField(1) = "Minor_NM"						

	arrHeader(0) = strPopPos						
	arrHeader(1) = strPopPos & "_명"			

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(strMajorCd, arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCustoms(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = strPopPos							
	arrParam(1) = "B_Minor"							
	arrParam(2) = Trim(strMinorCD)					
	arrParam(3) = ""								
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""	
	arrParam(5) = strPopPos							

	arrField(0) = "Minor_CD"						
	arrField(1) = "Minor_NM"						

	arrHeader(0) = strPopPos						
	arrHeader(1) = strPopPos & "_명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCustoms(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "화페"						
	arrParam(1) = "B_CURRENCY"						
	arrParam(2) = Trim(frm1.txtCurrency.value)		
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "화폐"						

	arrField(0) = "CURRENCY"						
	arrField(1) = "CURRENCY_DESC"					

	arrHeader(0) = "화폐"						
	arrHeader(1) = "화폐명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "영업그룹"						
	arrParam(1) = "B_SALES_GRP"							
	arrParam(2) = Trim(frm1.txtSalesGroup.value)		
	arrParam(3) = ""									
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						
	arrParam(5) = "영업그룹"						

	arrField(0) = "SALES_GRP"							
	arrField(1) = "SALES_GRP_NM"						

	arrHeader(0) = "영업그룹"						
	arrHeader(1) = "영업그룹명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetExCCNo(strRet)
	frm1.txtCCNo.value = strRet(0)
	frm1.txtCCNo.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetSORef(strRet)
		
	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	Call ggoOper.ClearField(Document, "2")	         						
	Call InitVariables
	Call SetDefaultVal

	frm1.txtSONo.value = Trim(strRet)
	strVal = BIZ_PGM_SOQRY_ID & "?txtSONo=" & Trim(frm1.txtSONo.value)	

	Call RunMyBizASP(MyBizASP, strVal)									

	Call ProtectSORelTag	
		
	frm1.txtRefFlg.value = "S"
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetLCRef(strRet)
	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	Call ggoOper.ClearField(Document, "2")	         						
	Call InitVariables
	Call SetDefaultVal

	strVal = BIZ_PGM_LCQRY_ID & "?txtMode=" & Parent.UID_M0001							
	strVal = strVal & "&txtLCNo=" & strRet(0)
	strVal = strVal & "&txtLCKind=" & strRet(1)

	Call RunMyBizASP(MyBizASP, strVal)											
		
	Call ProtectSORelTag
		
	frm1.txtRefFlg.value = "L"
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetDNRef(strRet)

	Call ggoOper.ClearField(Document, "2")	         						
	Call InitVariables
	Call SetDefaultVal

	frm1.txtApplicant.value = strRet(1)
	frm1.txtApplicantNm.value = strRet(2)
	frm1.txtCurrency.value = strRet(4)

	frm1.txtBeneficiary.value = Parent.gCompany
	frm1.txtBeneficiaryNm.value = Parent.gCompanyNm
		
	Call ReleaseSORelTag
	Call ReferenceQueryOk()
		
	frm1.txtRefFlg.value = "M"
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetCountry(strPopPos, arrRet)
	Select Case UCase(strPopPos)
		Case "LOADING"
			frm1.txtLoadingCntry.Value = arrRet(0)
			frm1.txtLoadingCntryNm.Value = arrRet(1)
			frm1.txtLoadingCntry.focus
		Case "DISCHARGE"
			frm1.txtDischgeCntry.Value = arrRet(0)
			frm1.txtDischgeCntryNm.Value = arrRet(1)
			frm1.txtDischgeCntry.focus
		Case "ORIGIN"
			frm1.txtOriginCntry.Value = arrRet(0)
			frm1.txtOriginCntryNm.Value = arrRet(1)
			frm1.txtOriginCntry.focus
		Case Else
	End Select

	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetSalesGroup(arrRet)
	frm1.txtSalesGroup.value = arrRet(0)
	frm1.txtSalesGroupNm.value = arrRet(1)
	frm1.txtSalesGroup.focus
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetBizPartner(strPopPos, arrRet)
	Select Case strPopPos
		Case "수입자"
			frm1.txtApplicant.Value = arrRet(0)
			frm1.txtApplicantNm.Value = arrRet(1)
			frm1.txtApplicant.focus	
		Case "수출자"
			frm1.txtBeneficiary.Value = arrRet(0)
			frm1.txtBeneficiaryNm.Value = arrRet(1)
			frm1.txtBeneficiary.focus
		Case "대행자"
			frm1.txtAgent.Value = arrRet(0)
			frm1.txtAgentNm.Value = arrRet(1)
			frm1.txtAgent.focus
		Case "제조자"
			frm1.txtManufacturer.Value = arrRet(0)
			frm1.txtManufacturerNm.Value = arrRet(1)
			frm1.txtManufacturer.focus
		Case "신고자"
			frm1.txtReporter.Value = arrRet(0)
			frm1.txtReporterNm.Value = arrRet(1)
			frm1.txtReporter.focus
		Case "환급신청인"
			frm1.txtReturnAppl.Value = arrRet(0)
			frm1.txtReturnApplNm.Value = arrRet(1)
			frm1.txtReturnAppl.focus	
		Case "운송신고인"
			frm1.txtTransRepCd.Value = arrRet(0)
			frm1.txtTransRepNm.Value = arrRet(1)
			frm1.txtTransRepCd.focus	
		Case Else
	End Select

	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetUnit(arrRet)
	frm1.txtWeightUnit.Value = arrRet(0)
	frm1.txtWeightUnit.focus

	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetMinorCd(strMajorCd, arrRet)
	Select Case strMajorCd
		Case gstrEDTypeMajor
			frm1.txtEDType.Value = arrRet(0)
			frm1.txtEDTypeNm.Value = arrRet(1)
			frm1.txtEDType.focus
		Case gstrPayTermsMajor
			frm1.txtPayTerms.Value = arrRet(0)
			frm1.txtPayTermsNm.Value = arrRet(1)
			frm1.txtPayTerms.focus
		Case gstrIncoTermsMajor
			frm1.txtIncoTerms.Value = arrRet(0)
			frm1.txtIncoTermsNm.Value = arrRet(1)
			frm1.txtIncoTerms.focus
		Case gstrOriginMajor
			frm1.txtOrigin.Value = arrRet(0)
			frm1.txtOriginNm.Value = arrRet(1)
			frm1.txtOrigin.focus
		Case gstrReturnOfficeMajor
			frm1.txtReturnOffice.Value = arrRet(0)
			frm1.txtReturnOfficeNm.Value = arrRet(1)
			frm1.txtReturnOffice.focus
		Case gstrTransFormMajor
			frm1.txtTransForm.Value = arrRet(0)
			frm1.txtTransFormNm.Value = arrRet(1)
			frm1.txtTransForm.focus
		Case gstrPackingTypeMajor
			frm1.txtPackingType.Value = arrRet(0)
			frm1.txtPackingTypeNm.Value = arrRet(1)
			frm1.txtPackingType.focus
		Case gstrCustomsMajor
			frm1.txtCustoms.Value = arrRet(0)
			frm1.txtCustomsNm.Value = arrRet(1)
			frm1.txtCustoms.focus
		Case gstrTransMethMajor
			frm1.txtTransMeth.value = arrRet(0)
			frm1.txtTransMethNm.value = arrRet(1)
			frm1.txtTransMeth.focus						
		Case Else

	End Select

	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetOpenPort(iwhere, arrRet)
	
	Select Case iwhere				
		Case 0
			frm1.txtLoadingPort.Value = arrRet(0)
			frm1.txtLoadingPortNm.Value = arrRet(1)	
			frm1.txtLoadingPort.focus
		Case 1
			frm1.txtDischgePort.Value = arrRet(0)
			frm1.txtDischgePortNm.Value = arrRet(1)	
			frm1.txtDischgePort.focus
	End Select			
					
	lgBlnFlgChgValue = True
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetCurrency(strRet)


	frm1.txtCurrency.value = strRet(0)
	Call CurFormatNumericOCX()
	frm1.txtCurrency.focus

End Function	
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetSalesGroup(strRet)
	frm1.txtSalesGroup.value = strRet(0)
	frm1.txtSalesGroupNm.value = strRet(1)
	frm1.txtSalesGroup.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetCustoms(strRet)
	frm1.txtCustoms.value = strRet(0)
	frm1.txtCustomsNm.value = strRet(1)
	frm1.txtCustoms.focus
End Function
'========================================================================================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877					
	Dim strTemp, arrVal

	Select Case Kubun
		
	Case 1,3
		Parent.WriteCookie CookieSplit , frm1.txtCCNo.value

	Case 0
		strTemp = Parent.ReadCookie(CookieSplit)
				
		If strTemp = "" then Exit Function
				
		frm1.txtCCNo.value =  strTemp
			
		If Err.number <> 0 Then
			Err.Clear
			Parent.WriteCookie CookieSplit , ""
			Exit Function 
		End If
			
		Call MainQuery()
						
		Parent.WriteCookie CookieSplit , ""
		
	Case 2	
		Parent.WriteCookie CookieSplit , "ED" & Parent.gRowSep & frm1.txtSalesGroup.value & Parent.gRowSep & frm1.txtSalesGroupNm.value & Parent.gRowSep & frm1.txtCCNo.value
		
	End Select
			 		
End Function
'========================================================================================================
Sub ProtectSORelTag()
	With frm1
		Call ggoOper.SetReqAttr(.txtCurrency, "Q")
		Call ggoOper.SetReqAttr(.txtBeneficiary, "Q")
		Call ggoOper.SetReqAttr(.txtAgent, "Q")
		Call ggoOper.SetReqAttr(.txtManufacturer, "Q")
		Call ggoOper.SetReqAttr(.txtPayTerms, "Q")
		Call ggoOper.SetReqAttr(.txtPayDur, "Q")
		Call ggoOper.SetReqAttr(.txtIncoTerms, "Q")
		Call ggoOper.SetReqAttr(.txtSalesGroup, "Q")
	End With
End Sub	
'========================================================================================================
Sub ReleaseSORelTag()
	With frm1
		Call ggoOper.SetReqAttr(.txtCurrency, "N")
		Call ggoOper.SetReqAttr(.txtBeneficiary, "N")
		Call ggoOper.SetReqAttr(.txtAgent, "D")
		Call ggoOper.SetReqAttr(.txtManufacturer, "D")
		Call ggoOper.SetReqAttr(.txtPayTerms, "D")
		Call ggoOper.SetReqAttr(.txtPayDur, "D")
		Call ggoOper.SetReqAttr(.txtIncoTerms, "N")
		Call ggoOper.SetReqAttr(.txtSalesGroup, "N")
	End With			
End Sub
'=========================================================================== 
Function JumpChgCheck(Byval IWhere)

	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then Exit Function
	End If

	Select Case IWhere
	Case 0
		Call CookiePage(1)
		Call PgmJump(EXCC_DETAIL_ENTRY_ID)
	Case 1
		Call CookiePage(2)
		Call PgmJump(EXPORT_CHARGE_ENTRY_ID)
	Case 2
		Call CookiePage(1)
		Call PgmJump(EXCC_ASSIGN_ENTRY_ID)
	Case 3
		Call CookiePage(3)
		Call PgmJump(EXCC_ADDITEM_ENTRY_ID)
	End Select
		
End Function
'============================================================================================================
Function ProtectXchRate()
	If frm1.txtCurrency.value = Parent.gCurrency Then
		Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
		frm1.txtXchRate.text = 1
		Call ggoOper.SetReqAttr(frm1.txtUsdXchRate, "Q")
	End If	
End Function
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'통관금액
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCCCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'FOB금액
		
		ggoOper.FormatFieldByObjectOfCur .txtFobDocAmt, .txtFobCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000,Parent.gComNumDec
		'환율
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtCurrency.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'USD환율
		ggoOper.FormatFieldByObjectOfCur .txtUsdXchRate, "USD", parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
				

	End With
End Sub
'========================================================================================================
Sub Form_Load()
	Call Parent.GetGlobalVar														
	Call LoadInfTB19029															
	Call Parent.AppendNumberPlace("6", "10", "0")
	Call Parent.AppendNumberPlace("7", "3", "0")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call CurFormatNumericOCX()
    Call GetValue_ko441()
	Call ggoOper.LockField(Document, "N")					
	Call SetDefaultVal
	Call CookiePage(0)	
	Call InitVariables
	Call changeTabs(TAB1)
		
	Call SetToolBar("11100000000011")									
		
	If UCase(Trim(frm1.txtCCNo.value)) <> "" Then
		Call MainQuery
	End If

	gSelframeFlg = TAB1
	frm1.txtCCNo.focus
    gIsTab     = "Y"
    gTabMaxCnt = 2  
		
End Sub

'========================================================================================================
Sub btnCCNoOnClick()
	Call OpenExCCNoPop()
End Sub
'========================================================================================================
Sub btnEDTypeOnClick()
	If frm1.txtEDType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtEDType.value, frm1.txtEDTypeNm.value, "신고구분", gstrEDTypeMajor)
	End If
End Sub
'========================================================================================================
Sub btnWeightUnitOnClick()
	If frm1.txtWeightUnit.readOnly <> True Then
		Call OpenUnit()
	End If
End Sub
'========================================================================================================
Sub btnCurrencyOnClick()
	If frm1.txtCurrency.readOnly <> True Then
		Call OpenCurrency()
	End If
End Sub
'========================================================================================================
Sub btnApplicantOnClick()
	If frm1.txtApplicant.readOnly <> True Then
		Call OpenBizPartner(frm1.txtApplicant.value, frm1.txtApplicantNm.value, "수입자")
	End If
End Sub
'======================================  3.2.5 btnBeneficiary_OnClick()  ================================
Sub btnBeneficiaryOnClick()
	If frm1.txtBeneficiary.readOnly <> True Then
		Call OpenBizPartner(frm1.txtBeneficiary.value, frm1.txtBeneficiaryNm.value, "수출자")
	End If
End Sub
'========================================================================================================
Sub btnAgentOnClick()
	If frm1.txtAgent.readOnly <> True Then
		Call OpenBizPartner(frm1.txtAgent.value, frm1.txtAgentNm.value, "대행자")
	End If
End Sub
'========================================================================================================
Sub btnManufacturerOnClick()
	If frm1.txtManufacturer.readOnly <> True Then
		Call OpenBizPartner(frm1.txtManufacturer.value, frm1.txtManufacturerNm.value, "제조자")
	End If
End Sub
'========================================================================================================
Sub btnPayTermsOnClick()
	If frm1.txtPayTerms.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPayTerms.value, frm1.txtPayTermsNm.value, "결제방법", gstrPayTermsMajor)
	End If
End Sub
'========================================================================================================
Sub btnIncoTermsOnClick()
	If frm1.txtIncoTerms.readOnly <> True Then
		Call OpenMinorCd(frm1.txtIncoTerms.value, frm1.txtIncoTermsNm.value, "가격조건", gstrIncoTermsMajor)
	End If
End Sub
'========================================================================================================
Sub btnLoadingPortOnClick()
	If frm1.txtLoadingPort.readOnly <> True Then
		Call OpenPort(frm1.txtLoadingPort.value, frm1.txtLoadingPortNm.value, "선적항", 0)
	End If
End Sub
'========================================================================================================
Sub btnDischgePortOnClick()
	If frm1.txtDischgePort.readOnly <> True Then
		Call OpenPort(frm1.txtDischgePort.value, frm1.txtDischgePortNm.value, "도착항", 1)
	End If
End Sub
'========================================================================================================
Sub btnOriginOnClick()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "원산지", gstrOriginMajor)
	End If
End Sub
'========================================================================================================
Sub btnLoadingCntryOnClick()
	If frm1.txtLoadingCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtLoadingCntry.value, "LOADING")
	End If
End Sub
'========================================================================================================
Sub btnDischgeCntryOnClick()
	If frm1.txtDischgeCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtDischgeCntry.value, "DISCHARGE")
	End If
End Sub
'========================================================================================================
Sub btnOriginOnClick()
	If frm1.txtOrigin.readOnly <> True Then
		Call OpenMinorCd(frm1.txtOrigin.value, frm1.txtOriginNm.value, "원산지", gstrOriginMajor)
	End If
End Sub
'========================================================================================================
Sub btnOriginCntryOnClick()
	If frm1.txtOriginCntry.readOnly <> True Then
		Call OpenCountry(frm1.txtOriginCntry.value, "ORIGIN")
	End If
End Sub
'========================================================================================================
Sub btnReporterOnClick()
	If frm1.txtReporter.readOnly <> True Then
		Call OpenBizPartner(frm1.txtReporter.value, frm1.txtReporterNm.value, "신고자")
	End If
End Sub
'========================================================================================================
Sub btnSalesGroupOnClick()
	If frm1.txtSalesGroup.readOnly <> True Then
		Call OpenSalesGroup()
	End If
End Sub
'========================================================================================================
Sub btnReturnApplOnClick()
	If frm1.txtReturnAppl.readOnly <> True Then
		Call OpenBizPartner(frm1.txtReturnAppl.value, frm1.txtReturnApplNm.value, "환급신청인")
	End If
End Sub
'========================================================================================================
Sub btnReturnOfficeOnClick()
	If frm1.txtReturnOffice.readOnly <> True Then
		Call OpenMinorCd(frm1.txtReturnOffice.value, frm1.txtReturnOfficeNm.value, "환급기관", gstrReturnOfficeMajor)
	End If
End Sub
'========================================================================================================
Sub btnTransFormOnClick()
	If frm1.txtTransForm.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransForm.value, frm1.txtTransFormNm.value, "컨테이너운송방법", gstrTransFormMajor)
	End If
End Sub
'========================================================================================================
Sub btnPackingTypeOnClick()
	If frm1.txtPackingType.readOnly <> True Then
		Call OpenMinorCd(frm1.txtPackingType.value, frm1.txtPackingTypeNm.value, "포장방법", gstrPackingTypeMajor)
	End If
End Sub
'========================================================================================================
Sub btnTransRepCdOnClick()
	If frm1.txtTransRepCd.readOnly <> True Then
		Call OpenBizPartner(frm1.txtTransRepCd.value, frm1.txtTransRepNm.value, "운송신고인")
	End If
End Sub
'========================================================================================================
Sub btnTransMethOnClick()
	If frm1.txtTransMeth.readOnly <> True Then
		Call OpenMinorCd(frm1.txtTransMeth.value, frm1.txtTransMethNm.value, "보세운송방법", gstrTransMethMajor)
	End If
End Sub
'========================================================================================================
Sub btnCustomsOnClick()
	If frm1.txtCustoms.readOnly <> True Then
		Call OpenCustoms(frm1.txtCustoms.value, frm1.txtCustomsNm.value, "세관", gstrCustomsMajor)
	End If
End Sub
'========================================================================================================
Sub chkSONoFlg_OnClick()
	frm1.txtSONoFlg.value = frm1.chkSONoFlg1.value
End Sub		
'=======================================================================================================
Sub txtIVDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIVDt.Action = 7 
        Call SetFocusToDocument("M")
        frm1.txtIVDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtEDDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEDDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtEDDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtShipFinDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtShipFinDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtShipFinDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtEPDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEPDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtEPDt.Focus
    End If
End Sub
'======================================================================================================
Sub txtTransFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTransFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtTransFromDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtTransToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTransToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtTransToDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtInspCertDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspCertDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtInspCertDt.Focus
    End If
End Sub
'=======================================================================================================
Sub txtQuarCertDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtQuarCertDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtQuarCertDt.Focus
    End If
End Sub
'========================================================================================================
Sub txtEDDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtIVDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtShipFinDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtEPDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtTransFromDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtTransToDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtInspCertDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtQuarCertDt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtGrossWeight_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtTotPackingCnt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtDocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtXchRate_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtFOBDocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtUSDXchRate_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtFOBLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================
Sub txtPayDur_Change()
	lgBlnFlgChgValue = True
End Sub	

'========================================================================================================


Sub txtCurrency_Onchange
    Call CurFormatNumericOCX()
End Sub

'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False												

	Err.Clear														

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")	         						

	Call InitVariables												
	Call SetDefaultVal

	If Not chkField(Document, "1") Then					
		Exit Function
	End If

	Call DbQuery()										

	FncQuery = True										
End Function
	
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False                                      

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")					
	Call ggoOper.LockField(Document, "N")					
	Call SetDefaultVal
	Call SetToolBar("11100000000011")						
	Call InitVariables										
	Call ReleaseSORelTag
	Call changeTabs(TAB1)
	FncNew = True												
End Function
'========================================================================================================
Function FncDelete()
	Dim IntRetCD

	FncDelete = False											
		
	If lgIntFlgMode <> Parent.OPMD_UMODE Then							
		Call Parent.DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
		
	IntRetCD = Parent.DisplayMsgBox("900003", Parent.VB_YES_NO, "x", "x")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	Call DbDelete												

	FncDelete = True											
End Function
'========================================================================================================
Function FncSave()
	Dim IntRetCD
		
	FncSave = False										
		
	Err.Clear											
		
	If lgBlnFlgChgValue = False Then					
	    IntRetCD = Parent.DisplayMsgBox("900001", "x", "x", "x")		
	    Exit Function
	End If

	If Not chkField(Document, "2") Then								
	    If gPageNo > 0 Then
	        gSelframeFlg = gPageNo
	    End If
	    Exit Function
	End If 

	
	If Len(Trim(frm1.txtEDDt.Text)) And Len(Trim(frm1.txtIVDt.Text)) Then
		If Parent.UniConvDateToYYYYMMDD(frm1.txtIVDt.Text, Parent.gDateFormat, "-") > Parent.UniConvDateToYYYYMMDD(frm1.txtEDDt.Text, Parent.gDateFormat, "-") Then
			Call Parent.DisplayMsgBox("970023", "x", frm1.txtEDDt.Alt, frm1.txtIVDt.Alt)			
			Call ClickTab1()
			frm1.txtEDDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtEPDt.Text)) And Len(Trim(frm1.txtEDDt.Text)) Then
		If Parent.UniConvDateToYYYYMMDD(frm1.txtEDDt.Text, Parent.gDateFormat, "-") > Parent.UniConvDateToYYYYMMDD(frm1.txtEPDt.Text, Parent.gDateFormat, "-") Then
			Call Parent.DisplayMsgBox("970023", "x", frm1.txtEPDt.Alt, frm1.txtEDDt.Alt)			
			Call ClickTab1()
			frm1.txtEPDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If Len(Trim(frm1.txtShipFinDt.Text)) And Len(Trim(frm1.txtEPDt.Text)) Then 
		If Parent.UniConvDateToYYYYMMDD(frm1.txtEPDt.Text, Parent.gDateFormat, "-") > Parent.UniConvDateToYYYYMMDD(frm1.txtShipFinDt.Text, Parent.gDateFormat, "-") Then
			Call Parent.DisplayMsgBox("970023", "x", frm1.txtShipFinDt.Alt, frm1.txtEPDt.Alt)			
			Call ClickTab1()
			frm1.txtShipFinDt.Focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If

	If frm1.txtXchRate.text <= 0 Then
		Call Parent.DisplayMsgBox("970023", "x", "환율","0")
		Call ClickTab1()
		frm1.txtXchRate.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	Call DbSave													
		
	FncSave = True												
End Function
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE												

	Call ggoOper.ClearField(Document, "1")									
	Call ggoOper.LockField(Document, "N")									
	frm1.txtCCNo1.value = "" 
	lgBlnFlgChgValue = True
End Function
'========================================================================================================
Function FncCancel() 
	On Error Resume Next												
End Function
'========================================================================================================
Function FncInsertRow()
	On Error Resume Next												
End Function
'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next												
End Function
'========================================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function
'========================================================================================
Function FncPrev() 
    Dim strVal
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")			
	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call Parent.DisplayMsgBox("900002", "x", "x", "x")  
        
        Exit Function
    End If

				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	frm1.txtPrevNext.value = "PREV"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001					
    strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo1.value)		
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)
	         
	Call RunMyBizASP(MyBizASP, strVal)
End Function
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")		
	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call Parent.DisplayMsgBox("900002", "x", "x", "x")  
        
        Exit Function
    End If
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	frm1.txtPrevNext.value = "NEXT"

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						
    strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo1.value)			
    strVal = strVal & "&txtPrevNext=" & Trim(frm1.txtPrevNext.value)	
	         
	Call RunMyBizASP(MyBizASP, strVal)
End Function
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLE)
End Function
'========================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE, True)
End Function
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = Parent.DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
'========================================================================================================
Function DbQuery()
	Err.Clear														

	DbQuery = False													

	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
    
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001					
	strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)		
	strVal = strVal & "&empty=empty"
		
	Call RunMyBizASP(MyBizASP, strVal)								
	
	DbQuery = True													
End Function
'========================================================================================================
Function DbSave()
	Err.Clear														
		
	DbSave = False													
		
	If frm1.chkSONoFlg.checked = True Then
		frm1.txtSoNoFlg.value = "Y"
	Else
		frm1.txtSoNoFlg.value = "N"
	End If	
		
	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
    
	With frm1
		.txtMode.value = Parent.UID_M0002									
		.txtFlgMode.value = lgIntFlgMode
		.txtInsrtUserId.value = Parent.gUsrID

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

	DbSave = True								
End Function

'========================================================================================================
Function DbDelete()
	Err.Clear									

	DbDelete = False							

	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003		
	strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo1.value)	

	Call RunMyBizASP(MyBizASP, strVal)							

	DbDelete = True												
End Function
'========================================================================================================
Function DbQueryOk()											

	lgIntFlgMode = Parent.OPMD_UMODE									

	Call ggoOper.LockField(Document, "Q")						
	Call SetToolBar("111110001101111")
		
	frm1.txtPrevNext.value = ""
		
	If frm1.txtRefFlg.value = "M" Then
		Call ReleaseSORelTag
	Else 
		Call ProtectSORelTag	
	End If

	If frm1.txtNetWeight.value > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtWeightUnit, "Q")
	End If

	If frm1.txtBillCount.value <> "0" Then
		Call ggoOper.SetReqAttr(frm1.txtXchRate, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtXchRate, "N")
	End If
		
	IF Len(Trim(frm1.txtSONo.value)) Then frm1.chkSONoFlg1.checked = True
		
	If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	End If
	lgBlnFlgChgValue = False
		
	frm1.txtCCNo.focus
		
End Function
'========================================================================================================
Function ReferenceQueryOk()												
	Call ProtectXchRate()
	Call SetToolBar("111010000000111")

End Function
'========================================================================================================
Function DbSaveOk()														
	Call InitVariables
	Call MainQuery()
End Function
'========================================================================================================
Function DbDeleteOk()												
	Call MainNew()
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2kcm.inc" --> 
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수출신고1</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수출신고2</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenSORef">수주참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCRef">L/C참조</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNRef">외주출고참조</A></TD>
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
											<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="통관관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" ALIGN=top TYPE="BUTTON" ONCLICK ="vbscript:btnCCNoOnClick()"></TD>
											<TD CLASS=TD6 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
						</TR>
						<TR>
							<TD WIDTH=100% VALIGN=TOP>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
												<TABLE <%=LR_SPACE_TYPE_60%>>
													<TR>
														<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo1" SIZE=20 MAXLENGTH=18 TAG="25XXXU" ALT="통관관리번호"></TD>
														<TD CLASS=TD5 NOWRAP>수주번호</TD>
														<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE=TEXT NAME="txtSONo" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="수주번호">&nbsp;&nbsp;&nbsp;
															<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkSONoFlg" ID="chkSONoFlg1">
															<LABEL FOR="chkSONoFlg">수주번호지정</LABEL>
														</TD>
													</TR>				
													<TR>	
														<TD CLASS=TD5 NOWRAP>송장번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIVNo" ALT="송장번호" MAXLENGTH=35 TYPE=TEXT SIZE=35 TAG="25XXXU">
														<TD CLASS=TD5 NOWRAP>작성일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtIVDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="작성일"></OBJECT>');</SCRIPT></TD>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>신고번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEDNo" ALT="신고번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="21XXXU">
														<TD CLASS=TD5 NOWRAP>신고일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtEDDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="신고일"></OBJECT>');</SCRIPT></TD>
													</TR>												
													<TR>
													    <TD CLASS=TD5 NOWRAP>면허번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEPNo" ALT="면허번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="21XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>면허일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 NAME="txtEPDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="면허일"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>Vessel명</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="Vessel명" TYPE=TEXT MAXLENGTH=50 SIZE=35 TAG="21XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>선적완료일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtShipFinDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="선적완료일"></OBJECT>');</SCRIPT></TD>																												
													</TR>
																										
													<TR>														
														<TD CLASS=TD5 NOWRAP>중량단위</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="중량단위" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnWeightUnitOnClick()"></TD>														
														<TD CLASS=TD5 NOWRAP>L/C번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C번호" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>Carton수</TD>
														<TD CLASS=TD6 NOWRAP>														
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtCarton" style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" ALT="Carton수" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														<TD CLASS=TD5 NOWRAP>총용적</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtMeasurement" TABINDEX = "-1" style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" ALT="총용적량" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>총중량</TD>
														<TD CLASS=TD6 NOWRAP>
														<!-- 2003/01/25 필수입력
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtGrossWeight" style="HEIGHT: 20px; WIDTH: 150px" tag="22X2Z" ALT="총중량" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														-->
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtGrossWeight" style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" ALT="총중량" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														<TD CLASS=TD5 NOWRAP>총순중량</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtNetWeight" TABINDEX = "-1" style="HEIGHT: 20px; WIDTH: 150px" tag="24X3Z" ALT="총순중량" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>화폐</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="23XXXU" ALT="화폐"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCurrencyOnClick()"></TD>																												
														<TD CLASS=TD5 NOWRAP>총포장개수</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtTotPackingCnt" TABINDEX = "-1" style="HEIGHT: 20px; WIDTH: 150px" tag="24X6" ALT="총포장개수" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
													</TR>
													
													<TR>
														<TD CLASS=TD5 NOWRAP>환율</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="22X5Z" ALT="환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														<TD CLASS=TD5 NOWRAP>USD환율</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtUsdXchRate" style="HEIGHT: 20px; WIDTH: 150px" tag="21X5Z" ALT="USD환율" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>통관금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" ALT="통관금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCCCurrency" ALT="통관금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU">
																	</TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>통관자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="통관자국금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCCCurrency" ALT="통관자국금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>FOB금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtFobDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="FOB금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
																	</TD>																
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtFobCurrency"  ALT="화폐" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>FOB자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtFobLocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2" ALT="FOB자국금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>																
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocFobCurrency" ALT="자국화폐" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>가격조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="가격조건"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoTerms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnIncoTermsOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="24"></TD>										
														<TD CLASS=TD5 NOWRAP>영업그룹</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=5 TAG="22XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>결제방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="21XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPayTermsOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>결제기간</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtPayDur" style="HEIGHT: 20px; WIDTH: 50px" TAG="21X7" ALT="결제기간" Title="FPDOUBLESINGLE"><PARAM NAME="MaxValue" VALUE="999"><PARAM NAME="MinValue" VALUE="0"></OBJECT>');</SCRIPT>&nbsp;일.</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수입자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24" ALT="수입자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>수출자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="22XXXU" ALT="수출자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnBeneficiaryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>대행자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnAgentOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>제조자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnManufacturerOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
													</TR>
												<%Call SubFillRemBodyTD5656(2)%>
												</TABLE>
								</DIV>
								<DIV ID="TabDiv" SCROLL=no>
												<TABLE <%=LR_SPACE_TYPE_60%>>												
													<TR>
														<TD CLASS=TD5 NOWRAP>선적항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLoadingPortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>선적항국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingCntry" ALT="선적항국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLoadingCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>도착항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDischgePortOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>도착항국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgeCntry" ALT="도착항국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgeCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnDischgeCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgeCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													
													
													<TR>
														<TD CLASS=TD5 NOWRAP>원산지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>원산지국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnOriginCntryOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
													</TR>														
													<TR>
														<TD CLASS=TD5 NOWRAP>최종목적지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="최종목적지" TYPE=TEXT MAXLENGTH=120 SIZE=35 TAG="21XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>신고자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReporter" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="신고자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReporter" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnReporterOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtReporterNm" SIZE=20 TAG="24"></TD>
													</TR>	
													<TR>
														<TD CLASS=TD5 NOWRAP>환급신청인</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReturnAppl" SIZE=10 MAXLENGTH=10 TAG="21XXXU" ALT="환급신청인"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReturnAppl" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnReturnApplOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtReturnApplNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>환급기관</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReturnOffice" SIZE=10 MAXLENGTH=30 TAG="21XXXU" ALT="환급기관"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReturnOffice" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnReturnOfficeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtReturnOfficeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>신고구분</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtEDType" SIZE=10 MAXLENGTH=5 TAG="21XXXU" ALT="신고구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEDType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnEDTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtEDTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>세관</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCustoms" ALT="세관" SIZE=10 MAXLENGTH=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCustoms" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnCustomsOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtCustomsNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>컨테이너 운송방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransForm" ALT="컨테이너 운송방법" SIZE=10 MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransForm" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTransFormOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransFormNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>포장조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" ALT="포장조건" SIZE=10 MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnPackingTypeOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운송신고인</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransRepCd" SIZE=10 MAXLENGTH=10 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransRepCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTransRepCdOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransRepNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>보세운송방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransMeth" SIZE=10 MAXLENGTH=5 TAG="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransMeth" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnTransMethOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtTransMethNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운송시작일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime5 NAME="txtTransFromDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="운송시작일"></OBJECT>');</SCRIPT></TD>
														<TD CLASS=TD5 NOWRAP>운송종료일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime6 NAME="txtTransToDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="운송종료일"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>검사증번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInspCertNo" ALT="검사증번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="21XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>검사증발급일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime7 NAME="txtInspCertDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="검사증발급일"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>검역증번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtQuarCertNo" ALT="검역증번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="21XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>검역증발급일</TD>
														<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime8 NAME="txtQuarCertDt" CLASS=FPDTYYYYMMDD tag="21X1" Title="FPDATETIME" ALT="검역증발급일"></OBJECT>');</SCRIPT></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>장치장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDevicePlce" ALT="장치장소" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>참고사항 1</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark1" ALT="참고사항1" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>2</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark2" ALT="참고사항2" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>3</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark3" ALT="참고사항3" TYPE=TEXT MAXLENGTH=120 SIZE=86 TAG="21XXXU"></TD>
													</TR>
												<%Call SubFillRemBodyTD5656(3)%>
												</TABLE>
								</DIV>
							</TD>
						</TR>
					</TABLE>		
				</TD>
			</TR>
			<TR HEIGHT=20>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_30%>>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck(2)">Container배정</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(0)">통관내역등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(1)">판매경비등록</A>&nbsp;|&nbsp;<A HREF = "VBSCRIPT:JumpChgCheck(3)">통관INVOICE추가내역등록</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHPayTerms" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHIncoterms" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHPayDur" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHCCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtLCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtRefFlg" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtSONoFlg" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtBillCount" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
