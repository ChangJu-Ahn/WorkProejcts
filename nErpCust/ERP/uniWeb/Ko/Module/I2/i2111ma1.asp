<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Physical Inventory Docunemt Creation
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat

'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/4/06
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Lee Seung Wook
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBSCRIPT">
Option Explicit                                                           

Const BIZ_PGM_ID = "i2111mb1.asp"										

<!-- #Include file="../../inc/lgvariables.inc" -->

 Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                  
    lgBlnFlgChgValue = False                          
    lgIntGrpCount = 0                            	                           	
    
    IsOpenPop = False								
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtInspDt.text  = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	lgBlnFlgChgValue = False
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End if
	If frm1.txtSLCd.value = "" Then
		frm1.txtSLNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	Set gActiveElement = document.activeElement  
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
Sub InitComboBox()
	'Call SetCombo(frm1.cboABCFlag, "A", "A")
	'Call SetCombo(frm1.cboABCFlag, "B", "B")
	'Call SetCombo(frm1.cboABCFlag, "C", "C")
	
	'ABC FLAG SEARCH B_MINOR 2005-03-18 LSW
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I1001", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboABCFlag ,lgF0  ,lgF0  ,Chr(11))
	
	Call SetCombo(frm1.cboCntPerd, "Y", "Y")
	Call SetCombo(frm1.cboCntPerd, "N", "N")	

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0004", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboInvMgr ,lgF0  ,lgF1  ,Chr(11))
	
End Sub

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
	
    if frm1.txtPlantCd.value  = "" then
    Call DisplayMsgBox("169901","X", "X", "X")
    frm1.txtPlantCd.focus    
    exit function
    End if
   
    If Plant_SLCd_Check(0) = False Then 
		Exit Function
	End If
   
    If IsOpenPop = True Then Exit Function
   
	IsOpenPop = True
	
	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")		' Where Condition
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

'------------------------------------------  OpenItemOrigin()  --------------------------------------------------
Function OpenItemOrigin()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If
	
    If Plant_SLCd_Check(0) = False Then 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "b1b11pa3", "X")			
		IsOpenPop = False
		Exit Function
    End If
 
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	       
	arrParam(1) = Trim(frm1.txtItemOriginCd.Value)	 
	arrParam(2) = ""				                 
	arrParam(3) = ""				                 
	
	arrField(0) = 1 
    arrField(1) = 2 
    arrField(2) = 9 
    arrField(3) = 6 
    arrField(4) = 45
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemOriginCd.focus
		Exit Function
	Else
		Call SetItemOrigin(arrRet)
	End If	
End Function

'------------------------------------------  OpenItemDest()  --------------------------------------------------
Function OpenItemDest()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If
	
    If Plant_SLCd_Check(0) = False Then 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
    End If
    
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	     
	arrParam(1) = Trim(frm1.txtItemDestCd.Value)	 
	arrParam(2) = ""				                 
	arrParam(3) = ""				                 
	
	arrField(0) = 1 
    arrField(1) = 2 
    arrField(2) = 9 
    arrField(3) = 6 
    arrField(4) = 45
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemDestCd.focus
		Exit Function
	Else
		Call SetItemDest(arrRet)
	End If	
End Function

'------------------------------------------  OpenItemGroupOrigin()  --------------------------------------------------
Function OpenItemGroupOrigin()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목그룹 팝업"					               
	arrParam(1) = "B_ITEM_GROUP"						               
	arrParam(2) = Trim(frm1.txtItemGroupOriginCd.Value)			  
	arrParam(3) = ""			 				                  
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "					              
	arrParam(5) = "품목그룹"						          
	
	arrField(0) = "ITEM_GROUP_CD"	
	arrField(1) = "ITEM_GROUP_NM"	
	
	arrHeader(0) = "품목그룹"		
	arrHeader(1) = "품목그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemGroupOriginCd.focus
		Exit Function
	Else
		Call SetItemGroupOrigin(arrRet)
	End If	
End Function

'------------------------------------------  OpenItemGroupDest()  --------------------------------------------------
Function OpenItemGroupDest()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹 팝업"				
	arrParam(1) = "B_ITEM_GROUP"					
	arrParam(2) = Trim(frm1.txtItemGroupDestCd.Value)		
	arrParam(3) = ""			 				            
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "					        
	arrParam(5) = "품목그룹"						    
	
	arrField(0) = "ITEM_GROUP_CD"	
	arrField(1) = "ITEM_GROUP_NM"	
	
	arrHeader(0) = "품목그룹"		
	arrHeader(1) = "품목그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemGroupDestCd.focus
		Exit Function
	Else
		Call SetItemGroupDest(arrRet)
	End If	
End Function

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "Tracking No"					  
	arrParam(1) = "s_so_tracking"					  
	arrParam(2) = frm1.txtTrackingNo.value 		
	
	arrParam(3) = ""							
	arrParam(4) = ""							
	arrParam(5) = "Tracking No"			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking No"		
    arrHeader(1) = "품목"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus		
	lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
Function SetSL(byRef arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetItemOrigin()  --------------------------------------------------
Function SetItemOrigin(byRef arrRet)
	frm1.txtItemOriginCd.Value	=arrRet(0)
	frm1.txtItemOriginNm.Value	=arrRet(1)
	frm1.txtItemOriginCd.focus
    lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetItemDest()  --------------------------------------------------
Function SetItemDest(byRef arrRet)
	frm1.txtItemDestCd.Value	=arrRet(0)
	frm1.txtItemDestNm.Value	=arrRet(1)
	frm1.txtItemDestCd.focus
	lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetItemGroupOrigin()  --------------------------------------------------
Function SetItemGroupOrigin(byRef arrRet)
	frm1.txtItemGroupOriginCd.Value	=arrRet(0)
	frm1.txtItemGroupOriginNm.Value	=arrRet(1)
	frm1.txtItemGroupOriginCd.focus
	lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetItemGroupDest()  --------------------------------------------------
Function SetItemGroupDest(byRef arrRet)
	frm1.txtItemGroupDestCd.Value	=arrRet(0)
	frm1.txtItemGroupDestNm.Value	=arrRet(1)
	frm1.txtItemGroupDestCd.focus
	lgBlnFlgChgValue = True  
End Function

'------------------------------------------  SetTrackingNo()  --------------------------------------------------
Function SetTrackingNo(byRef arrRet)
	frm1.txtTrackingNo.Value	=arrRet(0)
	frm1.txtTrackingNo.focus
	lgBlnFlgChgValue = True  
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

	Call InitVariables								
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)			
	Call ggoOper.LockField(Document, "N")							
	
	Call InitComboBox
	Call SetToolbar("10100000000011")
	Call SetDefaultVal
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	Set gActiveElement = document.activeElement
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

'=======================================================================================================
'   Event Name : txtInspDt_Change()
'=======================================================================================================
Sub txtInspDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboCntPerd_onchange()
'=======================================================================================================
Sub cboCntPerd_onchange()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : cboABCFlag_onchange()
'=======================================================================================================
Sub cboABCFlag_onchange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                            
    
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")                                
    Call ggoOper.LockField(Document, "N")                                 
    Call InitVariables															        
    Call SetToolbar("10100000000000")
    Call SetDefaultVal
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
    FncNew = True														
    Set gActiveElement = document.activeElement

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

    If Plant_SLCd_Check(1) = False Then 
       Call BtnDisabled(0)
		Exit Function
	End If

    '-----------------------
    'Save function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
       Call BtnDisabled(0)
       Exit Function
	End If
    
    If DbSave = False Then
       Call BtnDisabled(0)
       Exit Function
    End If
    
    FncSave = True                                                       
    lgBlnFlgChgValue = False
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
'=======================================================================================
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)
    Set gActiveElement = document.activeElement                                    
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , True)
    Set gActiveElement = document.activeElement                                       
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
    Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
'========================================================================================
Function DbSave() 

 Call LayerShowHide(1)
 
 Err.Clear
 DbSave = False														
 Dim strVal

 With frm1
      .txtMode.value = Parent.UID_M0002									
      .txtFlgMode.value = lgIntFlgMode
      Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
 End With	
 DbSave = True                                                      
   
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()														

	Dim PhyInvNo

	PhyInvNo = UCase(frm1.txtPhyInvNo.value)
 	
    Call DisplayMsgBox("169916","X" ,PhyInvNo, "X") 	
    Call InitVariables
    frm1.txtPhyInvNo.focus 

End Function

'========================================================================================
' Function Name : Plant_SLCd_Check
'========================================================================================
Function Plant_SLCd_Check(ByVal ChkIndex)

	Dim strInspDt
	Dim strClosedDt
	
	Plant_SLCd_Check = False

	'-----------------------
	'Check Plant CODE		
	'-----------------------
	If 	CommonQueryRs(" PLANT_NM, CONVERT(CHAR(10), INV_CLS_DT, 21)"," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	lgF1 = Split(lgF1,Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	strClosedDt = UniConvDateAToB(lgF1(0),Parent.gServerDateFormat,Parent.gDateFormat)
	
	If ChkIndex >= 1 Then

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
	
	Plant_SLCd_Check = True
   
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB5" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>
		</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사선별(Batch)</font></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%> >
							<TR>	
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD656" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="22XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>실사일자</TD>
								<TD CLASS="TD656">
								<script language =javascript src='./js/i2111ma1_fpDateTime1_txtInspDt.js'></script></TD>					
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD656">
								<INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="22XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=40 MAXLENGTH=40 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목그룹</TD>
								<TD CLASS="TD656">
								<INPUT TYPE=TEXT NAME="txtItemGroupOriginCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="품목그룹1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupOrigin" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemGroupOrigin()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupOriginNm" SIZE=25 MAXLENGTH=25 tag="24">&nbsp;~&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupDestCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="품목그룹2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupDest" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemGroupDest()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupDestNm" SIZE=25 MAXLENGTH=25 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD656">
								<INPUT TYPE=TEXT NAME="txtItemOriginCd" SIZE=15 MAXLENGTH=18 tag="21XXXU" ALT="품목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemOrigin" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemOrigin()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemOriginNm" SIZE=25 MAXLENGTH=25 tag="24">&nbsp;~&nbsp;<INPUT TYPE=TEXT NAME="txtItemDestCd" SIZE=15 MAXLENGTH=18 tag="21XXXU" ALT="품목2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemDest" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemDest()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemDestNm" SIZE=25 MAXLENGTH=25 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>실사번호</TD>
								<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtPhyInvNo" SIZE=16 MAXLENGTH=16 tag="25XXXU" ALT="실사번호"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>Tracking No</TD>
								<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=15 MAXLENGTH=25 tag="21XXXU" ALT="Tracking No"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" onclick=vbscript:OpenTrackingNo></TD>
							</TR>				
							<TR>
								<TD CLASS="TD5" NOWRAP>실사주기</TD>
								<TD CLASS="TD656"><SELECT Name="cboCntPerd" ALT="실사주기" STYLE="WIDTH: 60px" tag="21"><OPTION Value=""></OPTION></SELECT></TD>					
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ABC구분</TD>
								<TD CLASS="TD656"><SELECT Name="cboABCFlag" ALT="ABC구분" STYLE="WIDTH: 60px" tag="21"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>재고담당자</TD>
								<TD CLASS="TD656"><SELECT Name="cboInvMgr" ALT="재고담당자"  STYLE="WIDTH: 150px" tag="21"><OPTION Value=""></OPTION></SELECT></TD>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncSave()" Flag=1>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>	
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
