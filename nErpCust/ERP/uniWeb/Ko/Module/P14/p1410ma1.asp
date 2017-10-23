<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1410ma1.asp
'*  4. Program Name         : ECN Management
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/03/05
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1410mb1.asp"
Const BIZ_PGM_SAVE_ID = "p1410mb2.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim lgRdoOldVal1
Dim lgRdoOldVal2
Dim lgRdoOldVal3					

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo

'Dim blnFlgSetValue1
'Dim blnFlgSetValue2
Dim blnBomFlg
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
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

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtValidFromDt.text= UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtValidToDt.text	= UniConvDateAToB("2999-12-31", parent.gServerDateFormat, parent.gDateFormat)
	
	frm1.cboStatus.value = "2"
	frm1.txtEBomFlg.value = "N"
	frm1.txtMBomFlg.value = "N"
End Sub

Sub InitComboBox()
    On Error Resume Next
End Sub

'------------------------------------------  OpenECNInfo()  ----------------------------------------------
'	Name : OpenECNInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenECNInfo()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtECNNo.value)	' ECNNo
	arrParam(1) = ""						' ReasonCd
	arrParam(2) = ""						' Status
	arrParam(3) = ""						' EBomFlg
	arrParam(4) = ""						' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetECNInfo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtECNNo.Focus
	
End Function

'------------------------------------------  OpenReasonPopup()  ------------------------------------------
'	Name : OpenReasonPopup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtReasonCd.className) = UCase(parent.UCN_PROTECTED) Then
		Exit Function
	End If

	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "설계변경근거팝업"					' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) = UCase(Trim(frm1.txtReasonCd.value))	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""
	
	arrParam(5) = "설계변경근거"						' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
        
    arrHeader(0) = "설계변경근거"					' Header명(0)
    arrHeader(1) = "설계변경근거명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtReasonCd.Focus
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetECNInfo(byval arrRet)
	frm1.txtEcnNo.Value    = arrRet(0)		
	frm1.txtEcnDesc.Value  = arrRet(1)
	
	frm1.txtEcnNo.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetReasonInfo()  --------------------------------------------------
'	Name : SetReasonInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonInfo(byval arrRet)
	frm1.txtReasonCd.Value			= arrRet(0)	
	frm1.txtIssuedBy.Value			= arrRet(1)
	
	frm1.txtReasonCd.focus
	Set gActiveElement = document.activeElement
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************

'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029

	Call AppendNumberPlace("7","3","2")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
    'Call SetCookieVal
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
	Call InitVariables				

	frm1.txtECNNo.focus
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : cboClassMgr_onChange()
'   Event Desc :
'==========================================================================================
Sub cboStatus_onChange()
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                                        

	'-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
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
    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
    Call InitVariables															'⊙: Initializes local global variables
    frm1.txtECNDesc1.focus
    frm1.cboStatus.value = "2"
    Set gActiveElement = document.activeElement

    FncNew = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False                                                         

	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidTODt) = False Then Exit Function
	
	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001", "X", "X", "X") 
        Exit Function
    End If

	'-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             
       Exit Function
    End If

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
    
    frm1.txtECNNo1.value = ""
    frm1.cboStatus.value = "2"
    frm1.txtEBomFlg.value = ""
    frm1.txtEBomDt.text = ""
    frm1.txtMBomFlg.value = ""
    frm1.txtMBomDt.text = ""
    
    frm1.txtECNDesc1.focus

    Set gActiveElement = document.activeElement  

    lgBlnFlgChgValue = True
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

    LayerShowHide(1)

    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtECNNo=" & Trim(frm1.txtECNNo.value)
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
    Err.Clear         
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtECNNo=" & Trim(frm1.txtECNNo.value)			
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
    Call parent.FncFind(parent.C_SINGLE, False)                                         
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
'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         
    
    LayerShowHide(1)							
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtECNNo=" & Trim(frm1.txtECNNo.value)			
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
    Dim LayerN1
	frm1.hECNNo.value = frm1.txtECNNo1.value		'CHECK - MB1에서 할것인지 고려 
    
	Set LayerN1 = window.document.all("MousePT").style
	
    lgIntFlgMode = parent.OPMD_UMODE											
    lgBlnFlgChgValue = false
	frm1.txtEcnNo.focus 
	
	Set gActiveElement = document.activeElement 
    Call ggoOper.LockField(Document, "Q")
	Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtReasonCd, "Q")
	Call ggoOper.SetReqAttr(frm1.txtValidFromDt, "Q")
	Call ggoOper.SetReqAttr(frm1.txtValidToDt, "Q")
	Call ggoOper.SetReqAttr(frm1.cboStatus, "N")		'Q,N,D
	Call ggoOper.SetReqAttr(frm1.txtECNNo1,"Q")
	Call SetToolbar("11101000111111")
	

End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	LayerShowHide(1)
		
	With frm1
		.txtMode.value = parent.UID_M0002										
		.txtFlgMode.value = lgIntFlgMode
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
    
    DbSave = True    
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															

    dim LayerN1
   
	Set LayerN1 = window.document.all("MousePT").style
	
    Call InitVariables
    
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설계변경정보등록</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE CLASS="BasicTB" CELLSPACING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo" SIZE=20 MAXLENGTH=18 tag="12XXXU"  ALT="설계변경번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNNoPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtECNDesc" SIZE=60 tag="14"></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=2 WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					<!--<TABLE <%=LR_SPACE_TYPE_60%>>-->
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100%  valign=top>
									<FIELDSET>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR> 
												<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo1" SIZE=20 MAXLENGTH=18 tag="21xxxU"  ALT="설계변경번호"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경내용</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtECNDesc1" SIZE=80 MAXLENGTH=100 tag="22XXXX" ALT="설계변경내용"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경근거</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=10 MAXLENGTH=2 tag="22XXXU" ALT="설계변경근거"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReasonPopup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonPopup()">&nbsp;<INPUT TYPE=TEXT NAME="txtIssuedBy" SIZE=40 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경상태</TD>
												<TD CLASS=TD656 NOWRAP>
													<SELECT NAME="cboStatus" ALT="설계변경상태" STYLE="Width: 96px;" tag="22">
														<OPTION VALUE="1">Active</OPTION>
														<OPTION VALUE="2" SELECTED>Inactive</OPTION>
													</SELECT>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계BOM반영여부</TD>
												<TD CLASS=TD656 NOWRAP>
													<TABLE CELLSPACING=0>
														<TR>
															<TD><INPUT TYPE=TEXT NAME="txtEBomFlg" SIZE=10 MAXLENGTH=1 tag="24" ALT="설계BOM반영여부">&nbsp;</TD>
															<TD><script language =javascript src='./js/p1410ma1_I818333822_txtEBomDt.js'></script></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>생산BOM반영여부</TD>
												<TD CLASS=TD656 NOWRAP>
													<TABLE CELLSPACING=0>
														<TR>
															<TD><INPUT TYPE=TEXT NAME="txtMBomFlg" SIZE=10 MAXLENGTH=1 tag="24" ALT="생산BOM반영여부">&nbsp;</TD>
															<TD><script language =javascript src='./js/p1410ma1_I211735677_txtMBomDt.js'></script></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>시작일</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1410ma1_I619072836_txtValidFromDt.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>종료일</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1410ma1_I878183186_txtValidToDt.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>비고</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtRemark" SIZE=50 MAXLENGTH=40 tag="21" ALT="비고"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD656 NOWRAP>&nbsp;</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hECNNo" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
