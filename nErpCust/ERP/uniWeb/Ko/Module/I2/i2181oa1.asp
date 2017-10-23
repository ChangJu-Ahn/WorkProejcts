<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사과부족표출력 
'*  3. Program ID           : i21511Post phy inv Svr
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*			       i21511Post Phy Inv Svr
'*			       I21119Lookup Phy inv Svr
'*  7. Modified date(First) : 2000/04/13
'*  8. Modified date(Last)  : 2000/04/13
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Mrs Kim
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
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   ******************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->						
<!--'==========================================  1.1.1 Style Sheet  =========================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--'==========================================  1.1.2 공통 Include   =======================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                           

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim hPosSts
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                     				
    lgBlnFlgChgValue = False                              	                  			
    lgIntGrpCount = 0                            	                           			
    
    IsOpenPop = False							
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "OA") %>
End Sub

'------------------------------------------  OpenPhyInvNo()  --------------------------------------------
'	Name : OpenPhyInvNo()
'	Description : Pnysical Inventory No PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPhyInvNo()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam1,arrParam2,arrParam3,arrParam4
        
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus  
		Exit Function
	End If

	If Trim(frm1.txtSLCd.value)  = "" then 
		Call DisplayMsgBox("169902","X", "X", "X")
		frm1.txtSLCd.focus    
		Exit Function
	End If
	
    'If Plant_SLCd_PhyInvNo_Check(1) = False Then 
	'	Exit Function
	'End If

	iCalledAspName = AskPRAspName("i2111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "i2111pa1", "X")
		IsOpenPop = False
		Exit Function
    End If

	IsOpenPop = True

	arrParam1 = frm1.txtPhyInvNo.value
	arrParam2 = "PD"
	if frm1.txtPlantCd.value <> "" then	
	arrParam3 = frm1.txtPlantCd.value
	arrParam4 = frm1.txtSLCd.value
	end if
	
        	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam1,arrParam2,arrParam3,arrParam4), _
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
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
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
'	Name : OpenSL()
'	Description : SL Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus  
		Exit Function
	End If

    'If Plant_SLCd_PhyInvNo_Check(0) = False Then 
	'	Exit Function
	'End If


	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	if frm1.txtPlantCd.value <> "" then
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")		
	else
	arrParam(4) = ""
	end if
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

'------------------------------------------  SetPhyInvNo()  --------------------------------------------------
'	Name : SetPhyInvNo()
'	Description : OpenPhyInvNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPhyInvNo(byRef arrRet)

	frm1.txtPhyInvNo.Value    	= arrRet(0)
	frm1.txtSLCd.Value 			= arrRet(2)
	frm1.txtSLNm.Value 			= arrRet(3)	
	frm1.txtPlantCd.Value 		= arrRet(5)
	frm1.txtPlantNm.Value 		= arrRet(6)
	frm1.hPosSts.value 			= arrRet(4)		
	frm1.txtInspDt.text	      	= arrRet(1)
	frm1.txtPhyInvNo.focus	
	lgBlnFlgChgValue			= True	
	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus	
	lgBlnFlgChgValue	  	 = True	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
'	Name : SetSL()
'	Description : SL Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSL(byRef arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue	  = True
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call InitVariables									
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)		
	Call ggoOper.LockField(Document, "N")							
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("10000000000011")
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInspDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtInspDt.Focus
    End If
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Call BtnPreview()
    FncQuery = True    
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 

	If Not chkField(Document, "1") Then				
		Exit Function
	End If
    
    If Plant_SLCd_PhyInvNo_Check(2) = False Then 
		Exit Function
	End If
		
	dim var1, var2
    dim condvar
    Dim ObjName
    
    var1 = Trim(frm1.txtPhyInvNo.value)
	var2 = UCase(Trim(frm1.txtPlantCd.value))
    
	condvar = condvar & "PHYINVNO|" & var1
	condvar = condvar & "|PLANTCD|" & var2

	ObjName = AskEBDocumentName("i2181oa1", "ebr")
	Call FncEBRprint(EBAction, ObjName, condvar)
	
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                         
       Exit Function
    End If
    
    If Plant_SLCd_PhyInvNo_Check(2) = False Then 
		Exit Function
	End If

    Dim var1, var2
    Dim condvar
    Dim ObjName
    
    var1 = Trim(frm1.txtPhyInvNo.value)
	var2 = UCase(Trim(frm1.txtPlantCd.value))
	
	condvar = condvar & "PHYINVNO|" & var1
	condvar = condvar & "|PLANTCD|" & var2 

	ObjName = AskEBDocumentName("i2181oa1", "ebr")
	Call FncEBRPreview(ObjName, condvar)
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)                                  
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , True)                                         
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : Plant_SLCd_PhyInvNo_Check
' Function Desc : 
'========================================================================================
Function Plant_SLCd_PhyInvNo_Check(ByVal ChkIndex)
	'-----------------------
	'Check Plant CODE	
	'-----------------------
    If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Plant_SLCd_PhyInvNo_Check = False
		Exit function
    End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)

	If ChkIndex >= 1 Then
	'-----------------------
	'Check SLCd CODE	 
	'-----------------------
		If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125700","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Plant_SLCd_PhyInvNo_Check = False
			Exit function
		End If

		lgF0 = Split(lgF0, Chr(11))
		frm1.txtSLNm.Value = lgF0(0)
			
		If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("169922","X","X","X")
			frm1.txtSLCd.focus
			Plant_SLCd_PhyInvNo_Check = False
			Exit function
		End If
	End If

	If ChkIndex = 2  Then
	'-----------------------
	'Check PhyInvNo CODE
	'-----------------------
		If 	CommonQueryRs(" DOC_STS_INDCTR, CONVERT(CHAR(10), REAL_INSP_DT, 21) "," I_PHYSICAL_INVENTORY_HEADER ", _
		    " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S") & " AND PHY_INV_NO = " & FilterVar(frm1.txtPhyInvNo.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("160301","X","X","X")
			frm1.txtInspDt.Text = ""
			frm1.txtPhyInvNo.focus
			Plant_SLCd_PhyInvNo_Check = False
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		lgF1 = Split(lgF1,Chr(11))
		frm1.txtInspDt.text = UNIDateClientFormat(lgF1(0))
		If lgF0(0) <> "PD" Then
			Call DisplayMsgBox("169908","X","X","X")
			frm1.txtPhyInvNo.focus
			Plant_SLCd_PhyInvNo_Check = False
			Exit function
		End If
	End If

    Plant_SLCd_PhyInvNo_Check = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사과부족표출력</font></td>
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
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=20 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="12XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=25 MAXLENGTH=20 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">실사번호</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtPhyInvNo" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="실사번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPhyInvNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPhyInvNo()">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>실사일자</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i2181oa1_fpDateTime1_txtInspDt.js'></script></TD>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreView()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>                  
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
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>                  
</HTML>
