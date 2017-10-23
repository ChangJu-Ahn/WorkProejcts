<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2211MA1
'*  4. Program Name         : 판매계획환경설정정보 
'*  5. Program Desc         : 판매계획환경설정정보 
'*  6. Comproxy List        : PS2G211.dll
'*  7. Modified date(First) : 2002/12/13
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sonbumyeol
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<%'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID = "s2211mb1.asp"												'☆: Head Query 비지니스 로직 ASP명 
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim  EndDate,  StartDate 
EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

'==========================================  1.2.3 Global Variable값 정의  ===============================
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
 Dim IsOpenPop          
 

'#########################################################################################################
'												2. Function부 
'######################################################################################################### 
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                       	              '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                	              	'⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       		'⊙: Initializes Group View Size
	frm1.chkProcessBySg1.checked = True    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'☆: 사용자 변수 초기화 
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboConSpType.focus
End Sub

'========================================================================================================= 
Sub InitComboBox()
	
	' 판매계획유형 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConSpType,lgF0,lgF1,Chr(11))
	Call SetCombo2(frm1.cboSpType,lgF0,lgF1,Chr(11))
	
	'배분방법 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0019", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDistrMethodCfm,lgF0,lgF1,Chr(11))
	Call SetCombo2(frm1.cboDistrMethodFc,lgF0,lgF1,Chr(11))

	'잔량처리방법 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0020", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboPmRmnQty,lgF0,lgF1,Chr(11))
	
	'단가적용규칙 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0022", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboPriceRule,lgF0,lgF1,Chr(11))
	
	'환율구분 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("A1004", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboXchgRateFg,lgF0,lgF1,Chr(11))

	'환율처리 
	Call SetCombo(frm1.cboPmNonXchgRate, "C", "이전환율적용")
	Call SetCombo(frm1.cboPmNonXchgRate, "S", "Error")
	
End Sub


'==========================================  GetMethodofCreatePeriod()  ========================================
'	Name : GetMethodofCreatePeriod()
'	Description : 기간생성방법 Fetch
'========================================================================================================= 
Sub GetMethodofCreatePeriod()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	On Error Resume Next
	
	Err.Clear
	
	With frm1
		iStrSelectList	= " MN.MINOR_CD, MN.MINOR_NM "
		iStrFromList	= " dbo.B_MINOR MN INNER JOIN dbo.S_SP_PERIOD_HISTORY SP ON (SP.CREATE_METHOD = MN.MINOR_CD AND MN.MAJOR_CD = " & FilterVar("S0018", "''", "S") & ") "
		iStrWhereList	= " SP.FROM_DT <= GETDATE() " & _
						  " AND	SP.TO_DT >= GETDATE() " & _
						  " AND	SP.SP_TYPE =  " & FilterVar(.cboSptype.value , "''", "S") & " "
	
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrRs = Split(iStrRs, parent.gColSep)
			.txtHCreateMethod.value = Trim(iArrRs(1))
			.txtMethodofCrPeriod.value = Trim(iArrRs(2))
			.txtMethodofCrPeriod2.value = Trim(iArrRs(2))
		Else
			If Err.number = 0 Then
				iStrFromList	= " dbo.B_MINOR MN INNER JOIN dbo.B_CONFIGURATION CF ON (CF.MAJOR_CD = MN.MAJOR_CD AND CF.MINOR_CD = MN.MINOR_CD) "
				iStrWhereList	= " CF.MAJOR_CD = " & FilterVar("S0018", "''", "S") & " " & _
								  " AND CF.SEQ_NO = (SELECT CAST(REFERENCE AS SMALLINT) " & _
								  " FROM B_CONFIGURATION " & _
								  " WHERE MAJOR_CD = " & FilterVar("S0023", "''", "S") & " " & _
								  " AND	SEQ_NO = 1 " & _	
								  " AND	MINOR_CD =  " & FilterVar(.cboSptype.value , "''", "S") & ")"
	
				If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
					iArrRs = Split(iStrRs, parent.gColSep)
					.txtHCreateMethod.value = Trim(iArrRs(1))
					.txtMethodofCrPeriod.value = Trim(iArrRs(2))
					.txtMethodofCrPeriod2.value = Trim(iArrRs(2))
				Else
					If Err.number = 0 Then
						.txtHCreateMethod.value = ""
						.txtMethodofCrPeriod.value = ""
						.txtMethodofCrPeriod2.value = ""
					Else 
						MsgBox Err.description, vbInformation,Parent.gLogoName
						Err.Clear
						Exit Sub
					End If
				End If
			Else
				MsgBox Err.description, vbInformation,Parent.gLogoName
				Err.Clear
				Exit Sub
			End If
		End If
	
		' 일배분인 경우 Default 값 처리 
		If lgIntFlgMode = Parent.OPMD_CMODE And .txtHCreateMethod.value = "50" Then
			.cboDistrMethodCfm.value = "20"
			.cboDistrMethodfc.value = "20"
			.cboPmRmnQty.value = "10"
		End if
	End With

End Sub

'==========================================  LockFiled()  ========================================
'	Name : LockFiled()
'	Description : 판매계획유형, 기간생성방법에 따른 필드 Locking 처리  
'========================================================================================================= 
Sub LockField()

	With frm1
		' 배분방법이 일배분인 경우 
		If .txtHCreateMethod.value = "50" Then
			Call ggoOper.SetReqAttr(.cboDistrMethodCfm ,"Q")
			Call ggoOper.SetReqAttr(.cboDistrMethodfc ,"Q")
			Call ggoOper.SetReqAttr(.cboPmRmnQty ,"Q")
		Else
			Call ggoOper.SetReqAttr(.cboDistrMethodCfm ,"N")
			Call ggoOper.SetReqAttr(.cboDistrMethodfc ,"N")
			Call ggoOper.SetReqAttr(.cboPmRmnQty ,"N")
		End If

		' 판매계획유형에 따른 필드 Locking 처리 
		If .cboSpType.value = "E" Then
			Call ggoOper.SetReqAttr(.chkUseStep1 ,"N")
			Call ggoOper.SetReqAttr(.chkUseStep2 ,"N")
			
			If lgIntFlgMode = Parent.OPMD_UMODE Then
				Call chkUseStep1_onclick()
				Call chkUseStep2_onclick()
			End If
		Else
			Call ggoOper.SetReqAttr(.chkProcessByPlant1 ,"Q")
			Call ggoOper.SetReqAttr(.chkProcessByPlant2 ,"Q")
			Call ggoOper.SetReqAttr(.chkProcessBySg1 ,"Q")
			Call ggoOper.SetReqAttr(.chkSameQtyFlag1 ,"Q")
			Call ggoOper.SetReqAttr(.chkUseStep1 ,"Q")
			Call ggoOper.SetReqAttr(.chkUseStep2 ,"Q")
				
			If lgIntFlgMode = Parent.OPMD_CMODE Then
				.chkProcessByPlant1.checked = False
				.chkProcessByPlant2.checked = False
				.chkProcessBySg1.checked = True
				.chkSameQtyFlag1.checked = False
				.chkUseStep1.checked = False
				.chkUseStep2.checked = False
			End If
		End If
	End With

End Sub

'#########################################################################################################
'					3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029              '⊙: Load table , B_numeric_format
		                                             '⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","0")
	
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	
	Call SetToolBar("1110100000011111")         '⊙: 버튼 툴바 제어 
	Call InitComboBox
	Call InitVariables              '⊙: Initializes local global variables
	Call SetDefaultVal
	
End Sub


'==========================================================================================
'   Event Name : chkUseStep1_onclick()
'   Event Desc : 프로그램사용->공장별품목판매계획조정 onclick 이벤트 처리 
'==========================================================================================
Sub chkUseStep1_onclick()

		if frm1.chkUseStep1.checked then
			Call ggoOper.SetReqAttr(frm1.chkSameQtyFlag1,"N")
			Call ggoOper.SetReqAttr(frm1.chkProcessByPlant1,"N")
			Call ggoOper.SetReqAttr(frm1.chkProcessBySg1,"N")
		else
			frm1.chkSameQtyFlag1.checked = false
			Call ggoOper.SetReqAttr(frm1.chkSameQtyFlag1,"Q")
			frm1.chkProcessByPlant1.checked = false
			Call ggoOper.SetReqAttr(frm1.chkProcessByPlant1,"Q") 
			frm1.chkProcessBySg1.checked = true
			Call ggoOper.SetReqAttr(frm1.chkProcessBySg1,"Q")
		end if
		Call chkProcessByPlant1_onclick()
		lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : chkUseStep2_onclick()
'   Event Desc : 프로그램사용->공장별일별품목판매계획조정 onclick 이벤트 처리 
'==========================================================================================
Sub chkUseStep2_onclick()
		if frm1.chkUseStep2.checked then
			If frm1.chkProcessByPlant1.checked Then
				frm1.chkProcessByPlant2.checked = True
				Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"Q")
			Else
				Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"N")
			End If
		else
			If frm1.chkProcessByPlant1.checked Then
				frm1.chkProcessByPlant2.checked = True
			Else
				frm1.chkProcessByPlant2.checked = False
			End If
			Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"Q") 
		end if
		lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : chkSameQtyFlag1_onclick()
'   Event Desc : 전단계범위내수정->공장별일별품목판매계획조정 onclick 이벤트 처리 
'==========================================================================================
Sub chkSameQtyFlag1_onclick()
		if frm1.chkSameQtyFlag1.checked then
			'frm1.chkSameQtyFlag2.checked = true
			'Call ggoOper.SetReqAttr(frm1.chkSameQtyFlag2,"Q")
		else
			'Call ggoOper.SetReqAttr(frm1.chkSameQtyFlag2,"N")
		end if
		lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : chkProcessByPlant1_onclick()
'   Event Desc : 공장별진행->공장별품목판매계획확정 onclick 이벤트 처리 
'==========================================================================================

Sub chkProcessByPlant1_onclick()
		if frm1.chkProcessByPlant1.checked then
			frm1.chkProcessByPlant2.checked = true
			Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"Q")
		else
			If frm1.chkUseStep2.checked Then
				Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"N")
			Else
				frm1.chkProcessByPlant2.checked = False
				Call ggoOper.SetReqAttr(frm1.chkProcessByPlant2,"Q")
			End If
		end if
		lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : 
'   Event Desc : 데이타 변경을 체크하기 위한 이벤트처리함수 들 
'==========================================================================================

Sub cboSpType_onChange()
	lgBlnFlgChgValue = true
	Call GetMethodofCreatePeriod
	Call LockField
End Sub

'Sub chkSameQtyFlag2_onClick()
'	lgBlnFlgChgValue = true	
'End Sub

Sub chkProcessByPlant2_onClick()
	lgBlnFlgChgValue = true	
End Sub

Sub chkProcessBySg1_onClick()
	lgBlnFlgChgValue = true	
End Sub

'Sub chkProcessBySg2_onClick()
'	lgBlnFlgChgValue = true	
'End Sub

Sub txtFixedInterval_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub txtFcInterval_Change()
	lgBlnFlgChgValue = true	
End Sub

Sub cboDistrMethodCfm_onChange()
	lgBlnFlgChgValue = true	
End Sub

Sub cboDistrMethodFc_onChange()
	lgBlnFlgChgValue = true	
End Sub

Sub cboPmRmnQty_onChange()
	lgBlnFlgChgValue = true	
End Sub

Sub cboPriceRule_onChange()
	lgBlnFlgChgValue = true	
End Sub

Sub cboXchgRateFg_onChange()
	lgBlnFlgChgValue = true	
End Sub

Sub cboPmNonXchgRate_onChange()
	lgBlnFlgChgValue = true	
End Sub



'#########################################################################################################
'												5. Interface부 
'######################################################################################################### 
'========================================================================================
Function FncQuery() 

	Dim IntRetCD 
	
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear                                                               '☜: Protect system from crashing
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")		'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase contents area
	'----------------------- 
	Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
	Call InitVariables									'⊙: Initializes local global variables
	
  '-----------------------
	'Query function call area
	'----------------------- 
	If DBQuery = False Then									
		Exit Function
	End if
	
	FncQuery = True									'⊙: Processing is OK
        
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          <%'⊙: Processing is NG%>
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")                                      <%'⊙: Clear Contents  Field%>
    Call ggoOper.LockField(Document, "N")                                       <%'⊙: Lock  Suitable  Field%>
    Call SetToolbar("11101000000011")
	Call InitVariables              '⊙: Initializes local global variables
	Call SetDefaultVal	

	Set gActiveElement = document.activeElement 
    
    FncNew = True

End Function


'========================================================================================
 Function FncDelete() 
	Dim IntRetCD
	FncDelete = False									'⊙: Processing is NG
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then				                             	'Check if there is retrived data
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
	If DBDelete = False Then									
		Exit Function
	End if									'☜: Delete db data
	
	FncDelete = True                                                        					'⊙: Processing is OK

End Function


'========================================================================================
 Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False                                                         					'⊙: Processing is NG
	
	Err.Clear						                                                        '☜: Protect system from crashing
	
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                            				 '⊙: Check contents area
		Exit Function
	End If
	
	'-----------------------
	'Precheck area
	'-----------------------	
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                     		     	'⊙: No data changed!!		
		Exit Function
	End If
		
	'-----------------------
	'Save function call area
	'-----------------------	
	
	If DBSave = False Then									
		Exit Function
	End if	                              		                '☜: Save db data	
	FncSave = True                                                        					  '⊙: Processing is OK
    
End Function


'========================================================================================
Function FncCopy() 
     
End Function


'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	 Call parent.FncExport(Parent.C_SINGLE)											 '☜: 화면 유형 
End Function

'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , True)                                                    '☜: Protect system from crashing
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************

'========================================================================================
Function DbDelete() 
	Err.Clear                                                               					'☜: Protect system from crashing

   	Call LayerShowHide(1)

	DbDelete = False									'⊙: Processing is NG
	
	Dim iStrVal
	
	With frm1
		.txtMode.value		= Parent.UID_M0003							'☜: 비지니스 처리 ASP 의 상태 
		iStrVal = .cboSpType.value & parent.gColSep
		
		.txtSpreadDel.value = iStrVal
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With
	
	DbDelete = True			                                                   			'⊙: Processing is NG

End Function


'========================================================================================
Function DbDeleteOk()
	Call MainNew()
End Function


'========================================================================================
Function DbQuery() 
	
	Err.Clear                                                               					'☜: Protect system from crashing
   	Call LayerShowHide(1) 
	
	DbQuery = False                                                        					 '⊙: Processing is NG
	
	Dim iStrVal
	
	iStrVal = BIZ_PGM_ID & "?txtMode="     & Parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 
	iStrVal = iStrVal & "&txtSpType=" & frm1.cboConSpType.value
			
	Call RunMyBizASP(MyBizASP, iStrVal)							'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True                                                          					'⊙: Processing is NG

End Function

'========================================================================================
Function DbQueryOk()									'☆: 조회 성공후 실행로직 
	'-----------------------
	'Reset variables area
	'-----------------------
	
	lgIntFlgMode = Parent.OPMD_UMODE							'⊙: Indicates that current mode is Update mode

	Call ggoOper.LockField(Document, "Q")						'⊙: This function lock the suitable field
	
	Call GetMethodofCreatePeriod
	Call LockField
	
	Call SetToolbar("1111100000011111")										'버튼 툴바 제어 

	lgBlnFlgChgValue = False
End Function


'========================================================================================
Function DbSave() 
	Err.Clear																'☜: Protect system from crashing
	
   	Call LayerShowHide(1) 
	
	DbSave = False															'⊙: Processing is NG
	
	Dim iStrVal
	Dim IntTemp1, IntTemp2
	
	With frm1
		.txtMode.value		= Parent.UID_M0002							'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value	= lgIntFlgMode

		iStrVal = .cboSpType.value & parent.gColSep
		iStrVal = iStrVal & .txtFixedInterval.text & parent.gColSep
		iStrVal = iStrVal & .txtFcInterval.text & parent.gColSep
		iStrVal = iStrVal & .cboDistrMethodCfm.value & parent.gColSep
		iStrVal = iStrVal & .cboDistrMethodFc.value & parent.gColSep
		iStrVal = iStrVal & .cboPmRmnQty.value & parent.gColSep
		iStrVal = iStrVal & .cboPriceRule.value & parent.gColSep
		iStrVal = iStrVal & .cboXchgRateFg.value & parent.gColSep
		iStrVal = iStrVal & .cboPmNonXchgRate.value & parent.gColSep
		
		if .chkUseStep1.checked then IntTemp1 = 512 else IntTemp1 = 0 end if
		if .chkUseStep2.checked then IntTemp2 = 4096 else	IntTemp2 = 0 end if
		iStrVal = iStrVal & Cstr(IntTemp1 OR IntTemp2) & parent.gColSep
		
		if .chkSameQtyFlag1.checked then IntTemp1 = 512 else IntTemp1 = 0 end if
		' 실행계획인 경우 공장별 일별 품목 판매계획은 항상 공장별 품목판매계획수량과 일치해야 한다. - 2003.09.18
		if UCase(.cboSpType.value) = "E" then IntTemp2 = 4096	else IntTemp2 = 0 end if
		iStrVal = iStrVal & Cstr(IntTemp1 OR IntTemp2) & parent.gColSep
		
		if .chkProcessBySg1.checked then IntTemp1 = 256 else IntTemp1 = 0 end if
		'if .chkProcessBySg2.checked then IntTemp2 = 2048	else IntTemp2 = 0 end if
		iStrVal = iStrVal & IntTemp1 & parent.gColSep
		
		if .chkProcessByPlant1.checked then IntTemp1 = 1024 else IntTemp1 = 0 end if
		if .chkProcessByPlant2.checked then IntTemp2 = 8192	else IntTemp2 = 0 end if
		iStrVal = iStrVal & Cstr(IntTemp1 OR IntTemp2)

		if lgIntFlgMode = Parent.OPMD_UMODE then
			.txtSpreadUpd.value = iStrVal
		elseif lgIntFlgMode = Parent.OPMD_CMODE then
			.txtSpreadIns.value = iStrVal
		end if
				 
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With
	
	DbSave = True                                     					               '⊙: Processing is NG
    
End Function

'========================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
    Call InitVariables
	frm1.cboConSpType.value = frm1.cboSpType.value
	Call MainQuery
End Function

</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>판매계획환경설정정보</font></td>
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
									<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
									<TD CLASS="TD6"><SELECT Name="cboConSpType" ALT="판매계획유형" STYLE="WIDTH: 150px" tag="12XXXU"></SELECT></TD>
									<TD CLASS=TDT NOWRAP></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
								<TD CLASS="TD6"><SELECT Name="cboSpType" ALT="판매계획유형" STYLE="WIDTH: 150px" tag="23XXXU"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>확정구간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s2211ma1_fpDoubleSingle7_txtFixedInterval.js'></script>
									<INPUT NAME="txtMethodofCrPeriod" ALT="기간생성방법" TYPE="Text" MAXLENGTH="10" SIZE=13 tag="24XXXU" style="position:relative;top:-4;left:6">
								</TD>
								<TD CLASS=TD5 NOWRAP>예시구간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s2211ma1_fpDoubleSingle7_txtFcInterval.js'></script>
									<INPUT NAME="txtMethodofCrPeriod2" ALT="기간생성방법" TYPE="Text" MAXLENGTH="10" SIZE=13 tag="24XXXU" style="position:relative;top:-4;left:6">
								</TD>							
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>배분방법(확정)</TD>
								<TD CLASS="TD6"><SELECT Name="cboDistrMethodCfm" ALT="배분방법(확정)" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>배분방법(예시)</TD>
								<TD CLASS="TD6"><SELECT Name="cboDistrMethodFc" ALT="배분방법(예시)" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>잔량처리방법</TD>
								<TD CLASS="TD6"><SELECT Name="cboPmRmnQty" ALT="잔량처리방법" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>단가적용규칙</TD>
								<TD CLASS="TD6"><SELECT Name="cboPriceRule" ALT="단가적용규칙" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>환율구분</TD>
								<TD CLASS="TD6"><SELECT Name="cboXchgRateFg" ALT="환율구분" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>환율처리</TD>
								<TD CLASS="TD6"><SELECT Name="cboPmNonXchgRate" ALT="환율처리" STYLE="WIDTH: 150px" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>프로그램사용</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkUseStep1" ID="chkUseStep1" tag="21" Class="Check"><LABEL FOR="chkUseStep1">공장별품목판매계획조정</LABEL>&nbsp;&nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP>전단계범위내수정</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkSameQtyFlag1" ID="chkSameQtyFlag1" tag="24" Class="Check"><LABEL FOR="chkSameQtyFlag1">공장별품목판매계획조정</LABEL>&nbsp;&nbsp;
								</TD>																													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkUseStep2" ID="chkUseStep2" tag="21" Class="Check"><LABEL FOR="chkUseStep2">공장별일별품목판매계획조정</LABEL>&nbsp;&nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<!--INPUT TYPE=CHECKBOX NAME="chkSameQtyFlag2" ID="chkSameQtyFlag2" tag="21" Class="Check"><LABEL FOR="chkSameQtyFlag2">공장별일별품목판매계획조정</LABEL-->&nbsp;&nbsp;
								</TD>																													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>공장별진행</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkProcessByPlant1" ID="chkProcessByPlant1" tag="24" Class="Check"><LABEL FOR="chkProcessByPlant1">공장별품목판매계획확정</LABEL>&nbsp;&nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP>영업그룹별진행</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkProcessBySg1" ID="chkProcessBySg1" tag="24" Class="Check"><LABEL FOR="chkProcessBySg1">품목판매계획공장별배분</LABEL>&nbsp;&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=CHECKBOX NAME="chkProcessByPlant2" ID="chkProcessByPlant2" tag="24" Class="Check"><LABEL FOR="chkProcessByPlant2">공장별일별품목판매계획확정</LABEL>&nbsp;&nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
									<!--INPUT TYPE=CHECKBOX NAME="chkProcessBySg2" ID="chkProcessBySg2" tag="21" Class="Check"><LABEL FOR="chkProcessBySg2">품목판매계획일별배분</LABEL-->&nbsp;&nbsp;
								</TD>
							</TR>
							
							<%Call SubFillRemBodyTD5656(8)%>					
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpreadIns" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpreadUpd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpreadDel" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCreateMethod" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
