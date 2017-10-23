
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b03ma1.asp
'*  4. Program Name         : Item Group Management
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : Hong Chang Ho
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

Const BIZ_PGM_QRY_ID = "b1b03mb1.asp"           
Const BIZ_PGM_SAVE_ID = "b1b03mb2.asp"           
Const BIZ_PGM_DEL_ID = "b1b03mb3.asp"          

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
  frm1.txtValidFromDt.text = StartDate
  frm1.txtValidToDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")
  frm1.rdoLowItemGroupFlg2.checked = True
  lgRdoOldVal1 = 2  
End Sub

'------------------------------------------  OpenItemGroup()  --------------------------------------------
' Name : OpenItemGroup()
' Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd1.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd1.focus
 
End Function

'------------------------------------------  OpenHighItemGroup()  ----------------------------------------
' Name : OpenHighItemGroup()
' Description : HighItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenHighItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If IsOpenPop = True Or UCase(frm1.txtHighItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업" 
	arrParam(1) = "B_ITEM_GROUP"    
	arrParam(2) = Trim(frm1.txtHighItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "LEAF_FLG = " & FilterVar("N", "''", "S") & "  AND DEL_FLG = " & FilterVar("N", "''", "S") & "  AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & ""   
	arrParam(5) = "품목그룹"
 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
		   
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	   
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetHighItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtHighItemGroupCd.focus
	
End Function

'------------------------------------------  SetItemGroup()  -----------------------------------------
' Name : SetItemGroup()
' Description : ItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd1.Value    = arrRet(0)  
	frm1.txtItemGroupNm1.Value    = arrRet(1)  
End Function

'------------------------------------------  SetHighItemGroup()  -----------------------------------------
' Name : SetHighItemGroup()
' Description : HighItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetHighItemGroup(byval arrRet)
	lgBlnFlgChgValue = True 
 
	If Not ChkHighItemGroup(frm1.txtItemGroupCd2.value,arrRet(0)) Then Exit Function
 
	frm1.txtHighItemGroupCd.Value    = arrRet(0)   
	frm1.txtHighItemGroupNm.Value    = arrRet(1) 
End Function

'------------------------------------------  ChkHighItemGroup()  -----------------------------------------
' Name : ChkHighItemGroup(strData1, strData2)
' Description : 상위품목그룹과 품목그룹이 동일한 지 체크 
'---------------------------------------------------------------------------------------------------------
Function ChkHighItemGroup(strData1, strData2)
	ChkHighItemGroup = False
 
	If UCase(Trim(strData1)) = UCase(Trim(strData2)) Then
		Call DisplayMsgBox("127421", "X", "상위품목그룹", "품목그룹")
		frm1.txtHighItemGroupCd.value = ""
		frm1.txtHighItemGroupNm.value = "" 
		frm1.txthighItemGroupCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
 
	ChkHighItemGroup = True
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
	frm1.txtItemGroupCd1.focus
	Set gActiveElement = document.activeElement 
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
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
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
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : rdoLowItemGroupFlg1_OnClick()
'   Event Desc : change flag setting
'=======================================================================================================
Sub rdoLowItemGroupFlg1_OnClick()
	If lgRdoOldVal1 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 1
End Sub

'=======================================================================================================
'   Event Name : rdoLowItemGroupFlg2_OnClick()
'   Event Desc : change flag setting
'=======================================================================================================
Sub rdoLowItemGroupFlg2_OnClick()
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
	'    If frm1.txtItemGroupCd1.value = "" Then
	frm1.txtItemGroupNm1.value = ""
	' End If
 
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
    
    frm1.txtItemGroupCd2.focus
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
    frm1.txtItemGroupCd2.value = ""
 
	frm1.txtValidFromDt.Text  = StartDate
	frm1.txtValidToDt.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
    
    frm1.txtItemGroupCd2.focus
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
    strVal = strVal & "&txtItemGroupCd1=" & Trim(UCase(frm1.txtItemGroupCd1.value))
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
    strVal = strVal & "&txtItemGroupCd1=" & Trim(UCase(frm1.txtItemGroupCd1.value))
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
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003      
    strVal = strVal & "&txtItemGroupCd2=" & Trim(frm1.txtItemGroupCd2.value)
    
	Call RunMyBizASP(MyBizASP, strVal)          
 
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
    strVal = strVal & "&txtItemGroupCd1=" & frm1.txtItemGroupCd1.value
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
    Call SetToolbar("11111000111111")
    lgIntFlgMode = parent.OPMD_UMODE            
    lgBlnFlgChgValue = false
    
    frm1.txtItemGroupNm2.focus 
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
	 
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function     
	 
	With frm1
		.txtMode.value = parent.UID_M0002           
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)          
	End With
	 
	DbSave = True                                                           
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()               

    frm1.txtItemGroupCd1.value = frm1.txtItemGroupCd2.value 
    frm1.txtItemGroupNm1.value = frm1.txtItemGroupNm2.value     

    Call InitVariables
    
    Call MainQuery()

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목그룹등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>품목그룹</TD>
									<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtItemGroupCd1" SIZE=15 MAXLENGTH=10 tag="12XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()"> <INPUT TYPE=TEXT NAME="txtItemGroupNm1" size=50 tag="14"></TD>
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
												<TD CLASS=TD5 NOWRAP>품목그룹</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd2" SIZE=15 MAXLENGTH=10 tag="23XXXU"  ALT="품목그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm2" SIZE=50 MAXLENGTH=40 tag="22" ALT="품목그룹명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>상위품목그룹</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=15 MAXLENGTH=10 tag="21XXXU"  ALT="상위품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnHighItemGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenHighItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=50 tag="24"></TD>
											</TR>											
											<TR>
												<TD CLASS=TD5 NOWRAP>레벨</TD>
												<TD CLASS=TD656 NOWRAP><INPUT NAME="txtlevel1" SIZE=5 tag="24" ALT="레벨"></TD>												
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>최하위품목그룹여부</TD>
												<TD CLASS=TD656 NOWRAP>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLowItemGroupFlg" tag="2X" ID="rdoLowItemGroupFlg1" VALUE="Y"><LABEL FOR="rdoLowItemGroupFlg1">예</LABEL>
													<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLowItemGroupFlg" tag="2X" CHECKED ID="rdoLowItemGroupFlg2" VALUE="N"><LABEL FOR="rdoLowItemGroupFlg2">아니오</LABEL>
												</TD>													
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/b1b03ma1_I109819245_txtValidFromDt.js'></script>&nbsp;~&nbsp;
													<script language =javascript src='./js/b1b03ma1_I649504750_txtValidToDt.js'></script>										
												</TD>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="14"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../Inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
