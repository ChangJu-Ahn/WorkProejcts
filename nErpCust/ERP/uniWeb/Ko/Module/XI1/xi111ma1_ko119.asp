
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : INTERFACE
'*  2. Function Name        : 
'*  3. Program ID           : xi111ma1_ko119.asp
'*  4. Program Name         : INTERFACE SETING MANAGEMENT.
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2006/04/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../Inc/IncSvrCcm.inc" -->
<!-- #Include file="../../Inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../Inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../Inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../Inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../Inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../Inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "xi111mb1_ko119.asp"           
Const BIZ_PGM_SAVE_ID = "xi111mb2_ko119.asp"           
   

<!-- #Include file="../../Inc/lgVariables.inc" -->	

Dim lgBlnFlgConChg     '☜: Condition 변경 Flag
Dim IsOpenPop             
Dim lgRdoOldVal1
Dim lgRdoOldVal2

Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                                               
    lgBlnFlgChgValue = False                                                
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False              
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
'  frm1.txtValidFromDt.text = StartDate
 ' frm1.txtValidToDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")
  frm1.rdoFlg1.checked = True
  lgRdoOldVal1 = 1
End Sub

'------------------------------------------  OpenSystemId()  --------------------------------------------
' Name : OpenSystemId()
' Description : SystemId PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSystemId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSystemId1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "시스템ID팝업"
	arrParam(1) = "T_IF_SYSTEM_CONFIG_KO119"
	arrParam(2) = Trim(UCase(frm1.txtSystemId1.Value))
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "시스템ID"
	 
	arrField(0) = "ED15" & parent.gColSep & "SYSTEM_ID"
	arrField(1) = "ED30" & parent.gColSep &  "SYSTEM_NM"
	    
	arrHeader(0) = "시스템ID"
	arrHeader(1) = "시스템ID명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetSystemId(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtSystemId1.focus
 
End Function


'------------------------------------------  SetSystemId()  -----------------------------------------
' Name : SetSystemId()
' Description : SystemId Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetSystemId(byval arrRet)
	frm1.txtSystemId1.Value    = arrRet(0)  
	frm1.txtSystemIdNm1.Value    = arrRet(1)  
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function


'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
End Function
'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")           '⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11101000000011")

    Call SetDefaultVal
    Call InitVariables
	frm1.txtSystemId1.focus
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=======================================================================================================
'   Event Name : rdoFlg1_OnClick()
'   Event Desc : change flag setting
'=======================================================================================================
Sub rdoFlg1_OnClick()
	If lgRdoOldVal1 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 1
End Sub

'=======================================================================================================
'   Event Name : rdoFlg2_OnClick()
'   Event Desc : change flag setting
'=======================================================================================================
Sub rdoFlg2_OnClick()
	If lgRdoOldVal1 = 2 Then Exit Sub
 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 2    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear                 '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")     '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '-----------------------
	If frm1.txtSystemId1.value = "" Then
		frm1.txtSystemIdNm1.value = ""
	End If
 
    Call ggoOper.ClearField(Document, "2")         
    
    Call SetDefaultVal
    Call InitVariables               
    
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Exit Function          
    End If 
    
    FncQuery = True                
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	Dim IntRetCD 
    
    FncNew = False                
    
	'-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")                                       
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    Call InitVariables               
    
    frm1.txtSystemId2.focus
    Set gActiveElement = document.activeElement  
    
    FncNew = True                

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False              
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
	'-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")              
	If IntRetCD = vbNo Then
		Exit Function
	End If
    lgIntFlgMode="1003"
    If DbDelete = False Then   
		Exit Function           
    End If 
        
    FncDelete = True                                                        
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                       
    
    Err.Clear                                                             
    
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
    If Not chkField(Document, "2") Then                             
       Exit Function
    End If
    '-----------------------
    'check valid value
    '-----------------------
    If chkValidValue=False Then Exit Function 
    
    
	'-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If 
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")   
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE            
    
    ' 조건부 필드를 삭제한다.
    Call ggoOper.ClearField(Document, "1")                                  
    Call ggoOper.LockField(Document, "N")         
    Call SetToolbar("11101000000011")
    
    '---------------------------------------------------
    ' Default Value Setting
    '---------------------------------------------------
    frm1.txtPlantCd.value="" : frm1.txtPlantNm.value=""
    frm1.txtSystemId2.value = "" :frm1.txtSystemIdNM2.value = ""              
    frm1.txtSystemId2.focus
    Set gActiveElement = document.activeElement
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next                                                    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
     On Error Resume Next                                                   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                                    
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                              
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")          
    
    Call SetDefaultVal
    Call InitVariables               

    Err.Clear                                                              
    
    '------------------------------------
    ' Query Logic 수행 
    '------------------------------------    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001       
    strVal = strVal & "&txtSystemId1=" & Trim(UCase(frm1.txtSystemId1.value))
    strVal = strVal & "&PrevNextFlg=" & "P"  
    
	Call RunMyBizASP(MyBizASP, strVal)   

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD
 
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	   
    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")         
    
    Call SetDefaultVal
    Call InitVariables              
    
    Err.Clear                                                               
    
    '------------------------------------
    ' Query Logic 수행 
    '------------------------------------  
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001      
    strVal = strVal & "&txtSystemId1=" & Trim(UCase(frm1.txtSystemId1.value))
    strVal = strVal & "&PrevNextFlg=" & "N"  
    
	Call RunMyBizASP(MyBizASP, strVal)   

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
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False              
    
    LayerShowHide(1)
  
    Dim strVal
    
		If frm1.rdoFlg1.checked Then
		strVal =frm1.rdoFlg1.value
	Else
		strVal =frm1.rdoFlg2.value
	End If
	 
	With frm1
		.txtMode.value = parent.UID_M0003           
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtRdoFlg.value = strVal
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)          
	End With
    DbDelete = True                                                         
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()              
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                              
    
    DbQuery = False                                                        
 
	LayerShowHide(1)
     
    Dim strVal
      
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001      
    strVal = strVal & "&txtSystemId1=" & frm1.txtSystemId1.value
    strVal = strVal & "&PrevNextFlg=" & ""

    Call RunMyBizASP(MyBizASP, strVal)          
 
    DbQuery = True                                                          
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()              
 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11111000001111")
    lgIntFlgMode = parent.OPMD_UMODE            
    lgBlnFlgChgValue = false
    
    frm1.txtSystemIdNm2.focus 
    Set gActiveElement = document.activeElement  
     
    Call ggoOper.LockField(Document, "Q")         
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
 
	Err.Clear                

	DbSave = False               

	Dim strVal
	 
	If frm1.rdoFlg1.checked Then
		strVal =frm1.rdoFlg1.value
	Else
		strVal =frm1.rdoFlg2.value
	End If
	 
	With frm1
		.txtMode.value = parent.UID_M0002           
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtRdoFlg.value = strVal
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)          
	End With
	 
	DbSave = True                                                           
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()               

    frm1.txtSystemId1.value = frm1.txtSystemId2.value 
    frm1.txtSystemIdNm1.value = frm1.txtSystemIdNm2.value     

    Call InitVariables
    
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function chkValidValue() 
	Dim strPlant
	Dim strWhere
	Dim strDataNm
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

              
	strPlant=trim(frm1.txtPlantCd.value)	
	chkValidValue=True
	
	If strPlant=""  Then 
		frm1.txtPlantCd.value="*"
		frm1.txtPlantNM.value="*"
		Exit Function 
	ElseIf strPlant="*" Then 
		frm1.txtPlantNm.value="*"
		Exit Function 
	End If
	
	strWhere = " plant_cd = " & FilterVar(strPlant, "''", "S") & "  "

	Call CommonQueryRs(" plant_nm ","	 b_plant  ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("970000","X",frm1.txtPlantCd.alt,"X")			
		frm1.txtPlantNm.value = ""
		chkValidValue = False
		frm1.txtPlantCd.focus 
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	frm1.txtPlantNm.value = strDataNm(0)
		
	chkValidValue=True
    
End Function


</SCRIPT>
<!-- #Include file="../../Inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5" NOWRAP>시스템ID</TD>
									<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtSystemId1" SIZE=15 MAXLENGTH=10 tag="12XXXU"  ALT="시스템ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSystemId()"> <INPUT TYPE=TEXT NAME="txtSystemIdNm1" size=50 tag="14"></TD>
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
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100% valign=top>
									<FIELDSET>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>시스템ID</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSystemId2" SIZE=15 MAXLENGTH=10 tag="23XXXU"  ALT="시스템ID">&nbsp;<INPUT TYPE=TEXT NAME="txtSystemIdNm2" SIZE=50 MAXLENGTH=50 tag="23" ALT="시스템ID명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>공장</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=12 MAXLENGTH=4 tag="25XXXU"  ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnHighSytemId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=50 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효구분</TD>
												<TD CLASS=TD656 NOWRAP>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlg" tag="2X" ID="rdoFlg1" VALUE="Y" CHECKED><LABEL FOR="rdoFlg1">사용</LABEL>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlg" tag="2X"  ID="rdoFlg2" VALUE="N"><LABEL FOR="rdoFlg2">미사용</LABEL>
												</TD>													
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Alias Name</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtAliasNm" SIZE=50 MAXLENGTH=30 tag="22" ALT="Alias Name"></TD>
											</TR>											
											<TR>
												<TD CLASS=TD5 NOWRAP>IP Address</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtIPAdd" SIZE=50 MAXLENGTH=30 tag="22" ALT="IP Address"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>PORT 번호</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtPortNo" SIZE=50 MAXLENGTH=10 tag="22" ALT="PORT 번호"  ></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>환경설정파일</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtConfigFNm" SIZE=50 MAXLENGTH=50 tag="22" ALT="환경설정파일"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>환경설정STEP</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtConfigSNm" SIZE=50 MAXLENGTH=50 tag="22" ALT="환경설정파일"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>접속 URL</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtUrl" SIZE=50 MAXLENGTH=100 tag="21" ALT="접속 URL"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>접속ID</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtLoginId" SIZE=50 MAXLENGTH=30 tag="22" ALT="접속ID"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>접속비밀번호</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtLoginPwd" SIZE=50 MAXLENGTH=30 tag="22" ALT="접속비밀번호"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>담당자 E-Mail ID</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtEmail" SIZE=50 MAXLENGTH=100 tag="21xxx" ALT="담당자 E-Mail"></TD>
											</TR>											
											<TR>
												<TD CLASS=TD5 NOWRAP>비고</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtRemark" SIZE=120 MAXLENGTH=200 tag="21" ALT="비고"></TD>
											</TR>
											</TABLE>
									</FIELDSET>
								</TD>	
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hSytemId" tag="14"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtRdoFlg" tag="24">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../Inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
