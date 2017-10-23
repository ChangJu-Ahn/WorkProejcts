<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<%Response.Expires = -1%>
<!--
'======================================================================================================
'*  1. Module Name          :  
'*  2. Function Name        : 실제원가관리
'*  3. Program ID           : c4112ba1_ko441
'*  4. Program Name         : 재공수량 정리작업
'*  5. Program Desc         : 공장별 표준원가 계산을 실행한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/07
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Ig Sung, Cho
'* 10. Modifier (Last)      : Lee Tae Soo 
'* 11. Comment              :
'=======================================================================================================  -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================  -->
<!--'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'=======================================================================================================  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit																	<%'☜: indicates that All variables must be declared in advance%>

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "c4112bb1_ko441.asp"												'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
<!-- #Include file="../../inc/lgvariables.inc" -->		
'==========================================================================================================
'Dim lgBlnFlgChgValue
'Dim lgIntGrpCount
'Dim lgIntFlgMode
Dim IsOpenPop          

Dim lgIndPrevKey
Dim lgDirPrevKey
'Dim lgLngCurRows
'Dim lgSortKey
'Dim lgKeyStream

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
 '==========================================  5.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
        
    lgIndPrevKey = ""
    lgDirPrevKey = ""
    lgLngCurRows = 0
	lgSortKey = 1
	    
End Sub

 '******************************************  5.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim StartDate
	Dim EndDate

	Call FormatDATEField(frm1.txtYyyymm)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
	frm1.txtYyyymm.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
End Sub


'Sub SetDefaultVal()
'	Dim StartDate
'	Dim EndDate
	
'	Call FormatDATEField(frm1.txtYyyymm)
'	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 2)
'	StartDate	= "<%=GetSvrDate%>"
'	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
'	frm1.txtYyyymm.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    	
	
'End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "BA") %>
End Sub

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
    
    arrHeader(0) = "공장코드"
    arrHeader(1) = "공장명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	
	lgBlnFlgChgValue = True
End Function

Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("125000","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "15"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/B1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	

End Function

Function SetItemCd(Byval arrRet)
	With frm1
		 frm1.txtItemCd.focus
		.TxtItemCd.Value = arrRet(0)
		.TxtItemNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
	End With
End Function

Function ExeStdCost()
	Dim IntRetCD,iRow  
    Dim strVal
    Dim lGrpCnt
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    Err.Clear
    ExeStdCost = False	
    	
  '  if frm1.txtPlantCd.value = "" then
	'	frm1.txtPlantNm.value = ""
  '  end if
    
  ' if frm1.txtPlantCd.value <> "" then  
   ' IntRetCd = CommonQueryRs(" PLANT_CD "," B_PLANT "," plant_cd = " & FilterVar(Trim(frm1.txtPlantCd.Value)," ","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
		
	'If IntRetCD=False  Then
	'   Call DisplayMsgBox("125000","X","X","X")                         '☜ : 공장이 존재하지 않습니다 
	'   Exit Function
	'END IF
   'End if
   	
	'IF Trim(frm1.txtItemCd.Value) <> "" Then	
	'	IntRetCd = CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT "," plant_cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "  and item_cd = " & FilterVar(frm1.txtItemCd.Value," ","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
			
	'	If IntRetCD=False  Then
	'	   Call DisplayMsgBox("122700","X","X","X")                         '☜ : 공장이 존재하지 않습니다 
	'	   Exit Function
	'	END IF
	'END IF	
		 
    'If Not chkField(Document, "1") Then
    '   Exit Function
    'End If

	'If frm1.chkMaterial.checked = False And frm1.chkProcess.checked = False And frm1.chkIndirect.checked = False And frm1.chkRollUp.checked = False Then
	'	IntRetCD =  DisplayMsgBox("232520","x","x","x")
	'	Exit Function
	'End If

	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	Call MakeKeyStream("X")

	
	strVal = ""
	lGrpCnt = 0
	

    With frm1
        If .chkMaterial.checked = True Then
			strVal = strVal & "usp_CostSimulatation_ban_ko441" &Parent.gColSep &Parent.gRowSep
			lGrpCnt = lGrpCnt + 1		
		END If
		if .chkProcess.checked = True Then
			strVal = strVal & "usp_CostSimulatation_mem_ko441" &Parent.gColSep &Parent.gRowSep
			lGrpCnt = lGrpCnt + 1		
		END If
		if .chkIndirect.checked = True Then
			strVal = strVal & "usp_c_prod_in_qty_sum_batch_ko441" &Parent.gColSep &Parent.gRowSep
			lGrpCnt = lGrpCnt + 1		
		END If

 
       .txtMode.value        = Parent.UID_M0006
       .txtKeyStream.value   = lgKeyStream
	   .txtMaxRows.value     = lGrpCnt
	   .txtSpread.value      = strVal
	End With
	
'    With frm1
'        If .chkMaterial.checked = True Then
'			strVal = strVal & "usp_c_std_cost_by_material" &Parent.gColSep &Parent.gRowSep
'			lGrpCnt = lGrpCnt + 1		
'		END If
'		if .chkProcess.checked = True Then
'			strVal = strVal & "usp_c_std_cost_by_process" &Parent.gColSep &Parent.gRowSep
'			lGrpCnt = lGrpCnt + 1		
'		END If
'		if .chkIndirect.checked = True Then
'			strVal = strVal & "usp_c_std_cost_by_overhead" &Parent.gColSep &Parent.gRowSep
'			lGrpCnt = lGrpCnt + 1		
'		END If
'		if .chkRollUp.checked = True Then
'			strVal = strVal & "usp_c_std_cost_by_rollup" &Parent.gColSep &Parent.gRowSep
'			lGrpCnt = lGrpCnt + 1		
'		END If
' 
'       .txtMode.value        = Parent.UID_M0006
'       .txtKeyStream.value   = lgKeyStream
'	   .txtMaxRows.value     = lGrpCnt
'	   .txtSpread.value      = strVal
'	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	                                       '☜: 비지니스 ASP 를 가동 
	
    ExeStdCost = True         
End Function

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)

    '------ Developer Coding part (Start ) --------------------------------------------------------------



	'lgKeyStream = Trim(frm1.txtYyyymm.TEXT)	& Parent.gColSep
	lgKeyStream =  Trim(Replace(frm1.txtYyyymm.TEXT, "-", "")) & Parent.gColSep	

	'MSGBOX lgKeyStream
	
	'lgKeyStream = Trim(frm1.txtPlantCd.Value)	& Parent.gColSep
	
   ' IF Trim(frm1.txtItemCd.value) = "" Then
   '     lgKeyStream = lgKeyStream & "*" & Parent.gColSep
   ' ELSE
   '     lgKeyStream = lgKeyStream & Trim(frm1.txtItemCd.value) & Parent.gColSep
  '  END IF
	
     '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub



Sub Form_Load()

    Call LoadInfTB19029

    Call ggoOper.LockField(Document, "N")
    
    Call InitVariables
    
    Call SetDefaultVal
    Call SetToolbar("10000000000011")
    frm1.txtYyyymm.focus
    'frm1.txtPlantCd.focus
     
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Function FncQuery() 
    On Error Resume Next
End Function

Function FncSave() 
    On Error Resume Next
End Function

Function FncCopy() 
    On Error Resume Next
End Function

Function FncCancel() 
    On Error Resume Next
End Function

Function FncInsertRow() 
    On Error Resume Next
End Function

Function FncDeleteRow() 
    On Error Resume Next
End Function

Function FncPrint()
    Call parent.FncPrint()
End Function

Function FncPrev() 
	On Error Resume Next
End Function

Function FncNext() 
	On Error Resume Next
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
Dim IntRetCD
	FncExit = False
	
    'If lgBlnFlgChgValue = True Then
	'	IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
	'	If IntRetCD = vbNo Then
	'		Exit Function
	'	End If
    'End If

    FncExit = True
End Function

'========================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtYyyymm.focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재공수량 정리작업</font></td>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>

							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="작업년월" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>					
							<!--TR>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>품목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="1XXXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
							</TR !-->
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkMaterial" ID="chkMaterial" tag="11X" Class="RADIO" VALUE="Y"><LABEL FOR="chkMaterial">반도체 MES IF 정리</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkProcess" ID="chkProcess" tag="11X" Class="RADIO" VALUE="Y"><LABEL FOR="chkProcess">Mem MES IF  정리</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkIndirect" ID="chkIndirect" tag="11X" Class="RADIO" VALUE="Y"><LABEL FOR="chkIndirect">MES 입고자료집계</LABEL></TD>
							</TR>
							<!--TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME="chkRollUp" ID="chkRollUp" tag="11X" Class="RADIO" VALUE="Y"><LABEL FOR="chkRollUp">적상</LABEL></TD>
							</TR!-->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width=10> &nbsp; </TD>
					<TD><BUTTON NAME="btnExeStdCost" CLASS="CLSSBTN" onclick="ExeStdCost()" Flag=1>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

