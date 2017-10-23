<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i1511ma1.asp
'*  4. Program Name         : VMI Storage Location
'*  5. Program Desc         :
'*  6. Comproxy List        : PI5G010
'*							  PI5G120
'*  7. Modified date(First) : 2003/01/02
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : VB Conversion
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

Const BIZ_PGM_QRY_ID  = "i1511mb1.asp"									
Const BIZ_PGM_SAVE_ID = "i1511mb2.asp"									
Const BIZ_PGM_DEL_ID  = "i1511mb2.asp"									

Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->

'#########################################################################################################
'												2. Function부 
'######################################################################################################### 
'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                
    lgBlnFlgChgValue = False                                        
    lgIntGrpCount = 0                                               
    IsOpenPop = False												
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
End Sub


Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0003", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboSLGroup,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0004", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboInvMgr ,lgF0  ,lgF1  ,Chr(11))
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'========================================== 2.4.2 Open???()  =============================================
'------------------------------------------  OpenConPlant()  -------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

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
		Call SetConPlant(arrRet)
	End If	
End Function

'------------------------------------------  OpenConSLCd()  -------------------------------------------------
Function OpenConSLCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtslcd.className) = "PROTECTED"  Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("169901","X","X","X")    
		frm1.txtPlantCd.focus 
		Exit Function
	Else
		If Plant_SLCd_Check(0) = False Then Exit Function   
	End If

	IsOpenPop = True

	arrParam(0) = "VMI 창고팝업"									
	arrParam(1) = "I_VMI_STORAGE_LOCATION"								
	arrParam(2) = Trim(frm1.txtSLCd.Value)						        
	arrParam(3) = ""													
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "VMI 창고"												
	
    arrField(0) = "SL_CD"												
    arrField(1) = "SL_NM"												
    
    arrHeader(0) = "VMI 창고"										
    arrHeader(1) = "VMI 창고명"										
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetConSLCd(arrRet)
	End If	
End Function

'==========================================  2.4.3 Set???()  =============================================
'------------------------------------------  SetConPlant()  --------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
End Function

'------------------------------------------  SetConSLCd()  --------------------------------------------------
Function SetConSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
	frm1.txtSLCd.focus
End Function

'#########################################################################################################
'												3. Event부 
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
Sub cboSLGroup_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub cboInvMgr_Onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    
    Call InitVariables															
    Call LoadInfTB19029															
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")									
    
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal
    
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'#########################################################################################################
'												5. Interface부 
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                               
    
    Err.Clear                                                            

    If Not chkField(Document, "1") Then Exit Function						
    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X","X")			
		If IntRetCD = vbNo Then Exit Function
    End If
    
    If Plant_SLCd_Check(1) = False Then Exit Function                                     
    '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables														
    Call SetDefaultVal
    '-----------------------
    'Query function call area
    '----------------------- 
    If DBQuery = False Then Exit Function 
       
    FncQuery = True														
        
End Function

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
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")        
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                     
    Call ggoOper.LockField(Document, "N")                                      
    Call InitVariables															
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
    Set gActiveElement = document.activeElement
    FncNew = True															

End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False														
    
    If lgIntFlgMode = Parent.OPMD_CMODE Or _         
		UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(frm1.txthPlantCd.Value)) Or _       
		UCase(Trim(frm1.txtSLCd.Value)) <> UCase(Trim(frm1.txtSLCd1.Value)) Then             
       
        Call DisplayMsgBox("900002", "X","X","X" )      
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")		          
	If IntRetCD = vbNo Then Exit Function

    If DbDelete = False Then Exit Function
    
    FncDelete = True                                                    
    
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                      
    
    Err.Clear                                                              
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then Exit Function                            
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                       
        Exit Function
    End If
    
	If Trim(frm1.txtPlantCd.value) = "" then
	   Call DisplayMsgBox("169901", "X", "X", "X")
	   Exit Function
	End If
   
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		If UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(frm1.txthPlantCd.Value)) Or _   
			UCase(Trim(frm1.txtSLCd.Value)) <> UCase(Trim(frm1.txtSLCd1.Value)) Then        
			Call DisplayMsgBox("900002", "X","X","X" )       
			Exit Function
		End If
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
    FncSave = True                                                         
    
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 

	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE											
    frm1.txtSLCd.Value = ""
    frm1.txtSLNm.Value = ""
    Call ggoOper.LockField(Document, "N")								
    Call SetToolbar("11101000000011")
  
    frm1.txtSLCd1.value = ""
    frm1.txtSLCd1.focus
     
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
    Call parent.FncFind(Parent.C_SINGLE, False)                               
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")					
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
'========================================================================================
Function DbDelete() 
    Err.Clear                                                       
    
    DbDelete = False														
    
    Dim strVal
    
    Call LayerShowHide(1) 
    
    strVal = BIZ_PGM_DEL_ID &	"?txtMode="    & Parent.UID_M0003				& _						
								"&txtFlgMode=" & lgIntFlgMode					& _
								"&txtCommand=" & "DELETE"						& _
								"&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	& _
								"&txtSLCd1="   & Trim(frm1.txtSLCd1.value)				
	
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbDelete = True                                                       

End Function

'========================================================================================
' Function Name : DbDeleteOk
'========================================================================================
Function DbDeleteOk()													
    Call InitVariables
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                              
    
    DbQuery = False                                                        
    
    Dim strVal
    
    Call LayerShowHide(1) 
    
    strVal = BIZ_PGM_QRY_ID &	"?txtMode="    & Parent.UID_M0001				& _						
								"&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	& _
								"&txtSLCd="    & Trim(frm1.txtSLCd.value)				
	    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbQuery = True                                                        

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()													
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE											
    
    Call ggoOper.LockField(Document, "Q")								
	Call SetToolbar("11111000001111")
	frm1.txtSLCd.focus
End Function

'========================================================================================
' Function Name : DBSave
'========================================================================================
Function DbSave() 

    Err.Clear															

	DbSave = False														

    Dim strVal

	Call LayerShowHide(1) 
	
	frm1.txtMode.value       = Parent.UID_M0002								
	frm1.txtFlgMode.value    = lgIntFlgMode
		
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
    DbSave = True                                                        
    
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()														

    frm1.txtSLCd.value = frm1.txtSLCd1.value 
    
    Call InitVariables
    Call MainQuery()
	frm1.txtPlantCd.focus
End Function

'========================================================================================
' Function Name : Plant_SLCd_Check
'========================================================================================
Function Plant_SLCd_Check(ByVal ChkIndex)
	'-----------------------
	'Check Plant CODE		
	'-----------------------
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Parent.FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Plant_SLCd_Check = False
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
			
	If ChkIndex	>= 1 Then       

		'-----------------------
		'Check SLCd CODE	 
		'-----------------------
		If 	CommonQueryRs(" SL_NM "," I_VMI_STORAGE_LOCATION ", " SL_CD = " & Parent.FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("162001","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Plant_SLCd_Check = False
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtSLNm.Value = lgF0(0)
		
	End If
	
	Plant_SLCd_Check = True
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>VMI 창고등록</font></TD>
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
			<TABLE <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD <%=HEIGHT_TYPE_02%> >
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>VMI 창고</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="12XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSLCd()"> <INPUT TYPE=TEXT NAME="txtSLNm" SIZE = 40 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% >
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>VMI 창고</TD>
								<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd1" SIZE=10  MAXLENGTH=7 tag="23XXXU" ALT="창고">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm1" SIZE=25 MAXLENGTH = 40 tag="22" ALT="창고명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>창고그룹</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboSLGroup" ALT="창고그룹" STYLE="Width: 98px;" tag="21"><OPTION VALUE = ""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>재고담당자</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboInvMgr" ALT="재고담당자" STYLE="Width: 98px;" tag="21"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
						<% SubFillRemBodyTD656 (11)%>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	    <TD <%=HEIGHT_TYPE_01%> >
	    </TD>
	</TR>
	<TR HEIGHT=20 >
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%> >
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCommand" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
