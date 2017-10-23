<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사Posting Batch작업 
'*  3. Program ID           : i21511Post phy inv Svr
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*			       i21511Post Phy Inv Svr
'*			       I21119Lookup Phy inv Svr
'*  7. Modified date(First) : 2000/04/13
'*  8. Modified date(Last)  : 2006/08/29
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                         

Const BIZ_PGM_ID = "i2151mb1.asp"
Const BIZ_PGM_DEL_ID = "i2151mb2.asp"									

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
                                       	
    lgBlnFlgChgValue = False                              	                
    IsOpenPop = False							
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim strCanFlag
    
    If 	CommonQueryRs(" REFERENCE "," B_CONFIGURATION ", " MAJOR_CD = " & FilterVar("I0017", "''", "S") & " AND MINOR_CD = " & FilterVar("01","''","S") _
						& " AND SEQ_NO = 1 ", _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
		lgF0 = Split(lgF0, Chr(11))
		strCanFlag = lgF0(0)
	Else
		strCanFlag = "N"	
	End If
	
	If strCanFlag = "Y" Then
		txtCancelFlag.style.display = ""
	Else
		txtCancelFlag.style.display = "NONE"
	End If
	
End Sub

'------------------------------------------  OpenPhyInvNo()  --------------------------------------------
Function OpenPhyInvNo()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam1,arrParam2,arrParam3,arrParam4
        
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")  
		frm1.txtPlantCd.Focus
		Exit Function
	Else 
		'-----------------------
		'Check Plant CODE	
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.Focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	iCalledAspName = AskPRAspName("i2111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "i2111pa1", "X")
		IsOpenPop = False
		Exit Function
    End If
    
	IsOpenPop = True

	arrParam1 = frm1.txtPhyInvNo.value
	'************* PDY ADD LSW 2006-10-17**************
	If frm1.RadioOutputType.rdoCase1.Checked Then
		arrParam2 = "PDN"
	Else
		arrParam2 = "PDY"
	End If
	
	If Trim(frm1.txtPlantCd.value) <> "" then	
		arrParam3 = frm1.txtPlantCd.value
		arrParam4 = frm1.txtSLCd.value
	End If
	
   	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam1,arrParam2,arrParam3,arrParam4), _
 		 "dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPhyInvNo.focus
		Exit Function
	Else
    	Call SetPhyInvNo(arrRet)
	End If	
End Function

'------------------------------------------ OpenPlant()  --------------------------------------------------
Function OpenPlant()
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
	
	arrHeader(0) = "공장"		
	arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
End Function

'------------------------------------------  OpenSL()  --------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.Focus
		Exit Function
	Else 
		'-----------------------
		'Check Plant CODE	
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.Focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	If Trim(frm1.txtPlantCd.value) <> "" then
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")	
	Else
	arrParam(4) = ""
	End If
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
End Function

'------------------------------------------  OpenCostCd()  -------------------------------------------------
  Function OpenCostCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtCostCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd.Focus
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "Cost Center 팝업"			
	arrParam(1) = "B_COST_CENTER A,B_PLANT B"
	arrParam(2) = Trim(frm1.txtCostCd.Value)		
	arrParam(3) = ""								
	arrParam(4) = "A.BIZ_AREA_CD = B.BIZ_AREA_CD AND B.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "Cost Center"					
	
	arrField(0) = "COST_CD"							
	arrField(1) = "COST_NM"							
    
	arrHeader(0) = "Cost Center"			    	
	arrHeader(1) = "Cost Center 명"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		Call SetCostCd(arrRet)
	End If	
    
End Function

'==========================================  2.4.3 Set???()  =============================================
'------------------------------------------  SetPhyInvNo()  --------------------------------------------------
Function SetPhyInvNo(byRef arrRet)
	frm1.txtPhyInvNo.Value 	= arrRet(0)
	frm1.txtInspDt.value   	= arrRet(1)	
	frm1.txtSLCd.Value 		= arrRet(2)
	frm1.txtSLNm.Value 		= arrRet(3)	
	frm1.txtPlantCd.Value 	= arrRet(5)
	frm1.txtPlantNm.Value 	= arrRet(6)
	frm1.hPosSts.value 		= arrRet(4)		
	frm1.txtPhyInvNo.focus
	lgBlnFlgChgValue		= True	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	frm1.txtPlantCd.focus
	lgBlnFlgChgValue	  	 = True	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
Function SetSL(byRef arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue	  = True
End Function

'------------------------------------------  SetCostCd()  --------------------------------------------------
Function SetCostCd(byRef arrRet)
	frm1.txtCostCd.value	= arrRet(0)
	frm1.txtCostNm.value	= arrRet(1)
	frm1.txtCostCd.focus
	lgBlnFlgChgValue		= True
End Function

'------------------------------------------  OpenMvmtListRef()  -------------------------------------------------
' Name : OpenMvmtListRef()
' Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtListRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	Dim Param4
	 
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtPlantNm.value)
	Param3 = Trim(frm1.txtPhyInvNo.value)
	Param4 = Trim(frm1.txtInspDt.value)
	  
	If Param1 = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If Param3 = "" then
		Call DisplayMsgBox("169971","X", "X", "X")    
		frm1.txtPhyInvNo.focus
		Exit Function
	End If
	 
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2141RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2141RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")      
	     
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	End If 
End Function


'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then
		txtCostTitle.style.display = ""
	Else
		frm1.txtCostCd.tag = "25"
		ggoOper.SetReqAttr frm1.txtCostCd, "Q"
	End if

	Call InitVariables							
	Call LoadInfTB19029		
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")							
	
	Call SetToolbar("10000000000011")
	Call SetDefaultVal
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_DblClick(Button)
'=======================================================================================================
Sub txtInspDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtInspDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                              
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
         IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")         
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")                    
    Call ggoOper.LockField(Document, "N")  
                       
    Call InitVariables										
    Call SetToolbar("10000000000011")
    Call SetDefaultVal
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtPhyInvNo.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
    FncNew = True																

End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    Call BtnDisabled(1)
    FncSave = False                                       
    
    Err.Clear                                             
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                          
		Call BtnDisabled(0)
       Exit Function
    End If
        
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X", "X", "X")                 
        Call BtnDisabled(0)
        Exit Function
    End If
    
    If Plant_SLCd_PhyInvNo_Check = False Then 
		Call BtnDisabled(0)
		Call SetToolbar("10100000000011")
		Exit Function
	End If
    
    if Trim(frm1.txtCostCd.Value) <> "" then
		If 	CommonQueryRs(" COST_NM "," B_COST_CENTER ", " COST_CD = " & FilterVar(frm1.txtCostCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("124400","X","X","X")
				frm1.txtCostNm.Value = ""
				frm1.txtCostCd.focus
				Call BtnDisabled(0)
				Call SetToolbar("10100000000011")
				Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtCostNm.Value = lgF0(0)
	End If
	
    '-----------------------
    'Save function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Call SetToolbar("10100000000011")
		Exit Function
	End If
	
	'************* Delete ADD LSW 2006-10-17**************
	If frm1.RadioOutputType.rdoCase1.Checked Then	
		If DBSave() = False Then
			Call BtnDisabled(0)
			Exit Function
		End If
	Else
		If DBSave2() = False Then
			Call BtnDisabled(0)
			Exit Function
		End If
	End If
	
    FncSave = True                                                
    lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)                               
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , True)                               
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DBSave
'========================================================================================
Function DbSave() 

     Err.Clear	
     DbSave = False														
 
	frm1.txtMode.value			 = Parent.UID_M0002						
	frm1.txtFlgMode.value 		 = lgIntFlgMode
	Call LayerShowHide(1)	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)											

    DbSave = True                                                       
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()												
    Dim ItemDocNo
	
	ItemDocNo = frm1.hItemDocumentNo.value
	
    If  Trim(ItemDocNo) <> "" Then 
        Call DisplayMsgBox("169910","X",ItemDocNo, "X") 	
    Else
		Call DisplayMsgBox("800154","X", "X", "X") 
    End if

	Call SetToolbar("10100000000011")

End Function

'************* Delete ADD LSW 2006-10-17**************
'========================================================================================
' Function Name : DBSave2
'========================================================================================
Function DbSave2() 

     Err.Clear	
     DbSave2 = False														
 
	frm1.txtMode.value			 = Parent.UID_M0002						
	frm1.txtFlgMode.value 		 = lgIntFlgMode
	Call LayerShowHide(1)	
	Call ExecMyBizASP(frm1, BIZ_PGM_DEL_ID)											

    DbSave2 = True                                                       
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSaveOk2
'========================================================================================
Function DbSaveOk2()												
	Call DisplayMsgBox("800154","X", "X", "X") 
	Call SetToolbar("10100000000011")

End Function

'========================================================================================
' Function Name : Plant_SLCd_PhyInvNo_Check
'========================================================================================
Function Plant_SLCd_PhyInvNo_Check()

	Plant_SLCd_PhyInvNo_Check = False
	
	'-----------------------
	'Check PL/SLCd CODE	
	'-----------------------
	If 	CommonQueryRs(" A.SL_NM, B.PLANT_NM "," B_STORAGE_LOCATION A, B_PLANT B ", " A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.PLANT_CD = B.PLANT_CD AND A.SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus 
			Exit function
		Else
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)

			If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							
				Call DisplayMsgBox("125700","X","X","X")
				frm1.txtSLNm.Value = ""
				frm1.txtSLCd.focus
				Exit function
			End If
		
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtSLNm.Value = lgF0(0)

		End If
					
		Call DisplayMsgBox("169922","X","X","X")

		frm1.txtSLNm.Value = lgF0(0)
		frm1.txtPlantCd.focus
		Exit function
	End If
				
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1,Chr(11))
	frm1.txtSLNm.Value = lgF0(0)
	frm1.txtPlantNm.Value = lgF1(0)

	'-----------------------
	'Check PhyInvNo CODE	
	'-----------------------
	If 	CommonQueryRs("  A.POS_BLK_INDCTR, A.DOC_STS_INDCTR, CONVERT(CHAR(10), A.REAL_INSP_DT, 21), A.SL_CD, B.PLANT_CD "," I_PHYSICAL_INVENTORY_HEADER	A, I_PHYSICAL_INVENTORY_DETAIL B ", _
	    " A.PHY_INV_NO =" & FilterVar(frm1.txtPhyInvNo.Value, "''", "S") & " AND  A.PHY_INV_NO = B.PHY_INV_NO ", _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
		Call DisplayMsgBox("160301","X","X","X")
		frm1.txtInspDt.value = ""
		frm1.txtPhyInvNo.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	lgF1 = Split(lgF1,Chr(11))
	lgF2 = Split(lgF2,Chr(11))
	lgF3 = Split(lgF3,Chr(11))
	lgF4 = Split(lgF4,Chr(11))
	
	If UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(lgF4(0))) Then
		Call DisplayMsgBox("169943","X","X","X")
		frm1.txtPlantCd.focus
		Exit function
	End If

	If UCase(Trim(frm1.txtSLCd.Value)) <> UCase(Trim(lgF3(0))) Then
		Call DisplayMsgBox("160418","X","X","X")
		frm1.txtSLCd.focus
		Exit function
	End If
	
	If UCase(Trim(lgF1(0))) <> "PD" Then
		Call DisplayMsgBox("169908","X","X","X")
		frm1.txtPhyInvNo.focus
		Exit function
	ElseIf Trim(lgF0(0)) = "Y" Then
		If frm1.RadioOutputType.rdoCase1.Checked Then
			Call DisplayMsgBox("169907","X","X","X")
			frm1.txtPhyInvNo.focus
			Exit function
		End If
	End If 
	
	Set gActiveElement = document.activeElement
	frm1.txtInspDt.value = UniConvDateAToB(lgF2(0),Parent.gServerDateFormat,Parent.gDateFormat)	
	
    Plant_SLCd_PhyInvNo_Check = True
    
End Function

'============================================= rdoCase2_onclick()  ======================================
'=	Event Name : rdoCase2_onclick()
'=	Event Desc :
'========================================================================================================
Sub rdoCase2_onclick()
	frm1.txtCostCd.tag = "25"
	ggoOper.SetReqAttr frm1.txtCostCd, "Q"
End Sub

'============================================= rdoCase1_onclick()  ======================================
'=	Event Name : rdoCase1_onclick()
'=	Event Desc :
'========================================================================================================
Sub rdoCase1_onclick()
	frm1.txtCostCd.tag = "22"
	ggoOper.SetReqAttr frm1.txtCostCd, "N"
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사조정(Batch)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						  </TR>
						</TABLE>
					</TD>
					<!--<TD WIDTH=* align=right><A href="vbscript:OpenMvmtListRef()">실사선별후 수불참조</A></TD>-->
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>	
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%> >
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="22XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=20 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="22XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=25 MAXLENGTH=20 tag="24">
								</TD>
							</TR>
							<TR ID="txtCancelFlag" STYLE="DISPLAY: none">
								<TD CLASS="TD5" NOWRAP>구분</TD>
								<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" checked><LABEL FOR="rdoCase1">조정등록</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X"><LABEL FOR="rdoCase2">조정취소</LABEL>
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5">실사번호</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtPhyInvNo" SIZE=20 MAXLENGTH=16 tag="22XXXU" ALT="실사번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPhyInvNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPhyInvNo()">
								</TD>
							</TR>
							<TR ID="txtCostTitle" STYLE="DISPLAY: none">
								<TD CLASS="TD5" NOWRAP>Cost Center</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=23 MAXLENGTH=23 tag="24">
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>실사일자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspDt" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: center" tag="24X1" ALT="실사일자"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
 			     <TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Fncsave()" Flag=1>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPosSts" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemDocumentNo" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

