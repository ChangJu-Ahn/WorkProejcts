<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Move Type..... 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Mr  Koh
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

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                         

'==========================================  1.2.1 Global 상수 선언  ======================================
Const BIZ_PGM_QRY_ID  = "i1411mb1.asp"										
Const BIZ_PGM_SAVE_ID = "i1411mb2.asp"										
Const BIZ_PGM_DEL_ID  = "i1411mb3.asp"										
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgNextNo				
Dim lgPrevNo					
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                       	             
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
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.optDebitCreditFlag1.checked	= true
	frm1.optStckTypeCtrlFlag1.checked	= true
	frm1.optStckTypeCtrlFlag2.disabled	= true
	frm1.optPriceCtrlFlag1.checked		= true
	frm1.optPostCtrlFlag1.checked		= true
	frm1.optSLMovFlag1.checked			= true
	frm1.optPlantMovFlag1.checked		= true
	frm1.optItemMovFlag2.checked		= true	
	frm1.optTrackingNoMovFlag1.checked	= true
	frm1.txtMovType1.focus
	Set gActiveElement = document.activeElement
End Sub

'==========================================  2.2.2 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	'재고유형1
	call SetCombo(frm1.cboStckTypeFlagOrigin, "G", "양품재고")
	call SetCombo(frm1.cboStckTypeFlagOrigin, "B", "불량재고")
	call SetCombo(frm1.cboStckTypeFlagOrigin, "T", "이동중재고")
	call SetCombo(frm1.cboStckTypeFlagOrigin, "Q", "검사중재고")
	
	'재고유형2
	call SetCombo(frm1.cboStckTypeFlagDest, "G", "양품재고")
	call SetCombo(frm1.cboStckTypeFlagDest, "B", "불량재고")
	call SetCombo(frm1.cboStckTypeFlagDest, "T", "이동중재고")
	call SetCombo(frm1.cboStckTypeFlagDest, "Q", "검사중재고")
	
	'수불구분 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("I0002", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	call SetCombo2(frm1.cboTrnsType,lgF0,lgF1,Chr(11))


End Sub

'========================================== cboTrnsType_Onchange ========================================
'	Name : cboTrnsType_Onchange()
'	Description : Combo Display
'========================================================================================================= 
Sub cboTrnsType_Onchange()
	With frm1
		If .cboTrnsType.value = "PR" or _
		   .cboTrnsType.value = "MR" or _
		   .cboTrnsType.value = "OR" Then
			.optPriceCtrlFlag1.Checked = True
		Else 
			.optPriceCtrlFlag2.Checked = True
		End If	
	End With
End Sub


'------------------------------------------  OpenMovType1()  -------------------------------------------------
'	Name : OpenMovType1()
'	Description : Move Type PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMovType1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtMovType1.ClassName)= UCase(Parent.UCN_PROTECTED) Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "수불유형 팝업"					  
	arrParam(1) = "B_MINOR"						          
	arrParam(2) = Trim(frm1.txtMovType1.Value)			
	arrParam(3) = ""                 				    
	arrParam(4) = "MAJOR_CD = " & FilterVar("I0001", "''", "S") & ""					
	arrParam(5) = "수불유형"						
	
	arrField(0) = "MINOR_CD"	
	arrField(1) = "MINOR_NM"	
	
	arrHeader(0) = "수불유형"		
	arrHeader(1) = "수불유형명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMovType1.focus
		Exit Function
	Else
		Call SetMovType1(arrRet)
	End If	
End Function

'------------------------------------------  OpenMovType2()  -------------------------------------------------
'	Name : OpenMovType2()
'	Description : Move Type PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMovType2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtMovType2.ClassName)= UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "수불유형 팝업"				            
	arrParam(1) = "B_MINOR"					                    
	arrParam(2) = Trim(frm1.txtMovType2.Value)			        
	arrParam(3) = ""                             				
	arrParam(4) = "MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND MINOR_CD NOT IN (SELECT MOV_TYPE FROM I_MOVETYPE_CONFIGURATION)"						    ' Where Condition	
	arrParam(5) = "수불유형"					            
	
	arrField(0) = "MINOR_CD"	
	arrField(1) = "MINOR_NM"	
	
	arrHeader(0) = "수불유형"		
	arrHeader(1) = "수불유형명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMovType2.focus
		Exit Function
	Else
		Call SetMovType2(arrRet)
	End If	
End Function

'------------------------------------------  SetMovType1()  --------------------------------------------------
'	Name : SetMovType1()
'	Description : Move Type Conf. Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetMovType1(byval arrRet)
	frm1.txtMovType1.Value      = arrRet(0)
	frm1.txtMovTypeNm1.Value    = arrRet(1)
	frm1.txtMovType1.focus
End Function

'------------------------------------------  SetMovType2()  --------------------------------------------------
'	Name : SetMovType2()
'	Description : Move Type Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetMovType2(byval arrRet)
	frm1.txtMovType2.Value    	    = arrRet(0)
	frm1.txtMovTypeNm2.Value    	= arrRet(1)
	lgBlnFlgChgValue = True
End Function

 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	txtStckTypeCtrlFlag.style.display = "none"
	
	Call InitVariables																
	Call LoadInfTB19029																
	Call ggoOper.FormatField(Document, "2", CInt(ggAmtOfMoney.DecPoint), CInt(ggQty.DecPoint), _ 
                        CInt(ggUnitCost.DecPoint), CInt(ggExchRate.DecPoint), Parent.gDateFormat)
	
	Call ggoOper.LockField(Document, "N")											
	Call SetDefaultVal		

	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox	
	Call SetToolbar("11101000000011")										    
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	Dim IntRetCD 
	
	FncQuery = False                                                     
	
	Err.Clear                                                            
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then Exit Function
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then Exit Function
	End If
	
	'-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")				
	Call ggoOper.LockField(Document, "N")               
	Call InitVariables									

	If 	CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND MINOR_CD = " & FilterVar(frm1.txtMovType1.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("169950","X","X","X")
		frm1.txtMovTypeNm1.Value = ""
		frm1.txtMovType1.focus 
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtMovTypeNm1.Value = lgF0(0)
	
  '-----------------------
	'Query function call area
	'----------------------- 
	If DBQuery = False Then	Exit Function
	
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
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")           		
		If IntRetCD = vbNo Then Exit Function
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")                    				  
	Call ggoOper.LockField(Document, "N")                                     
	Call InitVariables														
	Call SetToolbar("1110100000011")
	Call SetDefaultVal

	
	FncNew = True								

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
 Function FncDelete() 
	Dim IntRetCD
	FncDelete = False								
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then		
		Call DisplayMsgBox("900002", "X", "X", "X")                                		
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then Exit Function
	'-----------------------
	'Delete function call area
	'-----------------------
	If DBDelete = False Then Exit Function								
 
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
	
	With frm1
		If .optPriceCtrlFlag1.Checked = True Then
			If .cboTrnsType.value = "PR" or _
			   .cboTrnsType.value = "MR" or _
			   .cboTrnsType.value = "OR" Then
			Else 
		        Call DisplayMsgBox("161005","X","X","X")                             
				Exit Function
			End If	
		End If
	End With

	'-----------------------
	'Save function call area
	'-----------------------	
	If DBSave = False Then Exit Function								

	FncSave = True                                                        				
    
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
	 Call parent.FncExport(Parent.C_SINGLE)											
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , True)                                                 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
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
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
 Function FncPrev() 
	Dim IntRetCD 
    
    FncPrev = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                
        Call DisplayMsgBox("900002","X","X","X")                             
        Exit Function
    End If
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then Exit Function								
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbPrev = False Then Exit Function  
   
	FncPrev = True

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
 Function FncNext() 
	
	Dim IntRetCD 
    
    FncNext = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                
        Call DisplayMsgBox("900002","X","X","X")                           
        Exit Function
    End If
	
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")					
		If IntRetCD = vbNo Then	Exit Function
    End If
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then	Exit Function							
    
    '-----------------------
    'Query function call area
    '----------------------- 
    If DbNext = False Then Exit Function  
    
	FncNext = False
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	Err.Clear                                                               				

   	Call LayerShowHide(1)

	DbDelete = False							
	
	Dim strVal
	
	strVal = BIZ_PGM_DEL_ID &	"?txtMode="     & Parent.UID_M0003				& _					
								"&txtMovType1=" & Trim(frm1.txtMovType1.value)			
		
	Call RunMyBizASP(MyBizASP, strVal)				
	
	
	DbDelete = True			                                                   		

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()
	Call InitVariables								
	Call MainNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	
	Err.Clear                                                               					
   	Call LayerShowHide(1) 
	
	DbQuery = False                                                        					 
	
	Dim strVal
	
	strVal = BIZ_PGM_QRY_ID &	"?txtMode="     & Parent.UID_M0001				& _
								"&txtMovType1=" & Trim(frm1.txtMovType1.value)	& _
								"&PrevNextFlg=" & ""
	
		
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
	lgIntFlgMode = Parent.OPMD_UMODE					

	Call ggoOper.LockField(Document, "Q")				
	
	Call SetToolbar("11111000110111")					
	frm1.txtMovType1.focus
	 
End Function

'========================================================================================
' Function Name : DbPrev
' Function Desc : This function is previous data query and display
'========================================================================================
Function DbPrev()

    Dim strVal

    DbPrev = False                                                       
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID &	"?txtMode="     & Parent.UID_M0001				& _
								"&txtMovType1=" & Trim(frm1.txtMovType1.value)	& _
								"&PrevNextFlg=" & "P"
    
	Call RunMyBizASP(MyBizASP, strVal)									
	
	DbPrev = True

End Function

'========================================================================================
' Function Name : DbNext
' Function Desc : This function is next data query and display
'========================================================================================
Function DbNext()

    DbNext = False                                                       
    
    Dim strVal
    
	LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID &	"?txtMode="     & Parent.UID_M0001				& _
								"&txtMovType1=" & Trim(frm1.txtMovType1.value)	& _
								"&PrevNextFlg=" & "N"
    
	Call RunMyBizASP(MyBizASP, strVal)									
	
	DbNext = True
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Err.Clear															
	
   	Call LayerShowHide(1) 
	
	DbSave = False														
	
	Dim strVal
	
	With frm1
		.txtMode.value		= Parent.UID_M0002							
		.txtFlgMode.value	= lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With
	
	DbSave = True                                     					         
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															
	
	frm1.txtMovType1.value = frm1.txtMovType2.value 
	
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수불유형 Conf. 등록</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						  </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=RIGHT>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=* >
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100% >
						<FIELDSET CLASS="CLSFLD" ALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>수불유형</TD>
									<TD CLASS="TD656">
									<INPUT TYPE=TEXT NAME="txtMovType1" SIZE=5 MAXLENGTH=3 tag="12XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMovType1()">&nbsp;<INPUT TYPE=TEXT NAME="txtMovTypeNm1" SIZE=40 tag="14">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>수불유형</TD>
								<TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtMovType2" SIZE=5 MAXLENGTH=3 tag="23XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMovType2()">&nbsp;<INPUT TYPE=TEXT NAME="txtMovTypeNm2" SIZE=40 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>재고증감구분</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 130px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optDebitCreditFlag" CHECKED ID="optDebitCreditFlag1" VALUE="D" tag="25"><LABEL FOR="optDebitCreditFlag1">증가(DEBIT)</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 130px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optDebitCreditFlag" ID="optDebitCreditFlag2" VALUE="C" tag="25"><LABEL FOR="optDebitCreditFlag2">감소(CREDIT)</LABEL></SPAN>
								</TD>
							</TR>
							<TR ID="txtStckTypeCtrlFlag" STYLE="DISPLAY: none" >
								<TD CLASS="TD5" NOWRAP>재고유형관리방법</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 130px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optStckTypeCtrlFlag" CHECKED ID="optStckTypeCtrlFlag1" VALUE="A" tag="25"><LABEL FOR="optStckTypeCtrlFlag1">자동지정(A)</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 130px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optStckTypeCtrlFlag" ID="optStckTypeCtrlFlag2" VALUE="U" tag="25"><LABEL FOR="optStckTypeCtrlFlag2">사용자지정(U)</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>재고유형1</TD>
								<TD CLASS="TD656"><SELECT Name="cboStckTypeFlagOrigin" ALT="재고유형1" STYLE="WIDTH: 150px" tag="23"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>재고유형2</TD>
								<TD CLASS="TD656"><SELECT Name="cboStckTypeFlagDest" ALT="재고유형2" STYLE="WIDTH: 150px" tag="23"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>재고단가반영구분</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPriceCtrlFlag" CHECKED ID="optPriceCtrlFlag1" VALUE="Y" tag="25"><LABEL FOR="optPriceCtrlFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPriceCtrlFlag" ID="optPriceCtrlFlag2" VALUE="N" tag="25"><LABEL FOR="optPriceCtrlFlag2">NO</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>수불구분</TD>
								<TD CLASS="TD656"><SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 150px" tag="23"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>회계Posting구분</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPostCtrlFlag" CHECKED ID="optPostCtrlFlag1" VALUE="Y" tag="25"><LABEL FOR="optPostCtrlFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPostCtrlFlag" ID="optPostCtrlFlag2" VALUE="N" tag="25"><LABEL FOR="optPostCtrlFlag2">NO</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고간 이동 여부</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optSLMovFlag" CHECKED ID="optSLMovFlag1" VALUE="Y" tag="25"><LABEL FOR="optSLMovFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optSLMovFlag" ID="optSLMovFlag2" VALUE="N" tag="25"><LABEL FOR="optSLMovFlag2">NO</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장간 이동 여부</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPlantMovFlag" CHECKED ID="optPlantMovFlag1" VALUE="Y" tag="25"><LABEL FOR="optPlantMovFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optPlantMovFlag" ID="optPlantMovFlag2" VALUE="N" tag="25"><LABEL FOR="optPlantMovFlag2">NO</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목간 이동 여부</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optItemMovFlag" ID="optItemMovFlag1" VALUE="Y" tag="25"><LABEL FOR="optItemMovFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optItemMovFlag" CHECKED ID="optItemMovFlag2" VALUE="N" tag="25"><LABEL FOR="optItemMovFlag2">NO</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>제번간 이동 여부</TD>
								<TD CLASS="TD656">
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optTrackingNoMovFlag" CHECKED ID="optTrackingNoMovFlag1" VALUE="Y" tag="25"><LABEL FOR="optTrackingNoMovFlag1">YES</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="optTrackingNoMovFlag" ID="optTrackingNoMovFlag2" VALUE="N" tag="25"><LABEL FOR="optTrackingNoMovFlag2">NO</LABEL></SPAN></TD>
							</TR>
							<% SubFillRemBodyTD656 (7)%>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

