<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c4101ba1
'*  4. Program Name         : 실제원가데이터삭제 
'*  5. Program Desc         : 실제원가데이터삭제 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit											'☜: indicates that All variables must be declared in advance 

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_ID = "C4107BB1.asp"		

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------


'@Global_Var
Dim lgBlnFlgChgValue           'Variable is for Dirty flag
Dim lgIntGrpCount              'Group View Size를 조사할 변수 
Dim lgIntFlgMode               'Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop
Dim lgSortKey      
Dim lgKeyStream    

'======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim StartDate
	Dim EndDate
	Dim IntRetCD
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)

	frm1.txtFrom_YYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtFrom_YYYYMM, Parent.gDateFormat, 2)
    
	frm1.txtTo_YYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtTo_YYYYMM, Parent.gDateFormat, 2)
    
	
   call  CommonQueryRs("max(yyyymm)","c_batch_job_check","flag = 'A' and work_step = '12' and progress_yn = 'Y' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

   frm1.txtCostYYYYMM.value = Replace(lgF0, Chr(11), "")

   
   
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '1') and grp = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM1.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate1.value = Replace(lgF1, Chr(11), "")

   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '2') and grp = '2' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM2.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate2.value = Replace(lgF1, Chr(11), "")   
   
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '3') and grp = '3' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM3.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate3.value = Replace(lgF1, Chr(11), "") 
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '4') and grp = '4' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM4.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate4.value = Replace(lgF1, Chr(11), "")   	      
	
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '5') and grp = '5' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM5.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate5.value = Replace(lgF1, Chr(11), "")   	
    
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA")%>
End Sub




'======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	
	Call InitVariables                                                     '⊙: Setup the Spread sheet
 
    Call SetDefaultVal
    Call SetToolbar("10000000000011")										'⊙: 버튼 툴바 제어 
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================

'======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)						'☜: 화면 유형 
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

End Sub


'======================================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'=======================================================================================================
Function FncExit()
	FncExit = True
End Function


Sub MakeKeyStream(pOpt)

	Dim sYear,sMon,sDay,sYYYYMM
	Dim eYear,eMon,eDay,eYYYYMM
    '------ Developer Coding part (Start ) --------------------------------------------------------------


	Call parent.ExtractDateFromSuper(frm1.txtFrom_YYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
	Call parent.ExtractDateFromSuper(frm1.txtTo_YYYYMM.Text, parent.gDateFormat,eYear,eMon,eDay)	
	
	sYYYYMM = sYear & sMon
	eYYYYMM = eYear & eMon
	
		
	lgKeyStream = sYYYYMM	& Parent.gColSep 
	lgKeyStream = lgKeyStream & eYYYYMM	& Parent.gColSep 	
	

     '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


Function FncBtnExe() 
	FncBtnExe = False   

	Dim lGrpCnt,strVal
	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	Call MakeKeyStream("X")


	lGrpCnt=0
	
	strVal = ""
	

    With frm1
        If .chkLevel1.checked = True Then
			strVal = strVal & "1" & Parent.gColSep 
			lGrpCnt = lGrpCnt + 1		
		END If
		if .chkLevel2.checked = True Then
			strVal = strVal & "2" & Parent.gColSep 
			lGrpCnt = lGrpCnt + 1		
		END If
		if .chkLevel3.checked = True Then
			strVal = strVal & "3" & Parent.gColSep 
			lGrpCnt = lGrpCnt + 1		
		END If
		if .chkLevel4.checked = True Then
			strVal = strVal & "4" & Parent.gColSep 
			lGrpCnt = lGrpCnt + 1		
		END If
 
 		if .chkLevel5.checked = True Then
			strVal = strVal & "5" & Parent.gColSep 
			lGrpCnt = lGrpCnt + 1		
		END If
 
       .txtMode.value        = Parent.UID_M0006
       .txtKeyStream.value   = lgKeyStream
	   .txtMaxRows.value     = lGrpCnt
	   .txtSpread.value      = strVal
	End With
	
	IF lGrpCnt = 0 Then
		Call DisplayMsgBox("232520","x","x","x")
		IF LayerShowHide(0) = False Then
			Exit Function
		END IF		
		Exit Function
	End IF
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'☜: 비지니스 ASP 를 가동 
	
	FncBtnExe = True                                      	                    '⊙: Processing is OK
End Function

Function FncBtnExeOK()				            '☆: 저장 성공후 실행 로직 

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

   	Call DisplayMsgBox("990000","X","X","X")
  
    call  CommonQueryRs("max(yyyymm)","c_batch_job_check"," flag = 'a' and work_step = '12' and progress_yn = 'Y' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

   frm1.txtCostYYYYMM.value = Replace(lgF0, Chr(11), "")

   
   
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '1') and grp = '1' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM1.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate1.value = Replace(lgF1, Chr(11), "")

   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '2') and grp = '2' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM2.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate2.value = Replace(lgF1, Chr(11), "")   
   
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '3') and grp = '3' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM3.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate3.value = Replace(lgF1, Chr(11), "") 
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '4') and grp = '4' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM4.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate4.value = Replace(lgF1, Chr(11), "")   	      
	
   call  CommonQueryRs(" to_yyyymm, convert(varchar(10),job_dt,120 )","C_DATA_DEL_HISTORY"," SEQ = (select max(seq) from C_DATA_DEL_HISTORY where grp = '5') and grp = '5' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    

	frm1.txtDelYYYYMM5.value = Replace(lgF0, Chr(11), "")
	frm1.txtLastDate5.value = Replace(lgF1, Chr(11), "")   	
	
	frm1.chkLevel1.checked = False
    frm1.chkLevel2.checked = False
    frm1.chkLevel3.checked = False
    frm1.chkLevel4.checked = False
    frm1.chkLevel5.checked = False
End Function

Sub txtFrom_YYYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrom_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrom_YYYYMM.focus
	End If
End Sub

Sub txtTo_YYYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtTo_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTo_YYYYMM.focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
'======================================================================================================= -->



<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>실제원가Data삭제</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
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
									<TD CLASS="TD5" NOWRAP>최종원가계산년월</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCostYYYYMM" SIZE=15 MAXLENGTH=6 tag="14XXXU" ALT="최종원가계산년월"></TD>
									<TD CLASS="TD6" NOWRAP>&nbsp</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrom_YYYYMM" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT="시작작업년월" id=txtFrom_YYYYMM></OBJECT>');</SCRIPT>
															&nbsp;~&nbsp<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtTo_YYYYMM" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT="시작작업년월" id=txtTo_YYYYMM></OBJECT>');</SCRIPT>	</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp</TD>
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
								<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>LEVEL1 계산용 Data삭제</LEGEND>
								<TABLE WIDTH=600 <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>원가계산용 임시성Data,삭제무방 </TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종삭제년월</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDelYYYYMM1" SIZE=15  tag="14XXXU" ALT="최종삭제년월"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>&nbsp;&nbsp;관련:실제원가계산에러내역,공정별수불Batch</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종작업일</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLastDate1" SIZE=15  tag="14XXXU" ALT="최종작업일"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP><LABEL FOR="chkLevel1">선택</LABEL><INPUT TYPE=CHECKBOX NAME="chkLevel1" ID="chkLevel1" tag="11X" Class="RADIO" VALUE="Y"></TD>
								</TR>								
								</TABLE>
															
							</FIELDSET>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
			
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>LEVEL2 실제원가구성Rate 삭제</LEGEND>
								<TABLE WIDTH=600 <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>원가계산용 Data, 이월되는Data, 최종원가계산월Data는 삭제안됨</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종삭제년월</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDelYYYYMM2" SIZE=15 tag="14XXXU" ALT="최종삭제년월"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>&nbsp;&nbsp;관련:공정별원가집계,배부요소Data(자동),C/C,공정별배부내역</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종작업일</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLastDate2" SIZE=15 tag="14XXXU" ALT="최종작업일"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP><LABEL FOR="chkLevel2">선택</LABEL><INPUT TYPE=CHECKBOX NAME="chkLevel2" ID="chkLevel2" tag="11X" Class="RADIO" VALUE="Y"></TD>
								</TR>								
								</TABLE>
															
							</FIELDSET>
							</TR>
						</TABLE>
					</TD>
				</TR>					
				
				<TR>
			
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>LEVEL3 집계성 Data삭제</LEGEND>
								<TABLE WIDTH=600 <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>조회용 집계Data,일부화면 조회불가능</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종삭제년월</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDelYYYYMM3" SIZE=15 tag="14XXXU" ALT="최종삭제년월"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>&nbsp;&nbsp;관련:공정별원가집계,배부요소Data(자동),C/C,공정별배부내역,등</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종작업일</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLastDate3" SIZE=15 tag="14XXXU" ALT="최종작업일"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP><LABEL FOR="chkLevel3">선택</LABEL><INPUT TYPE=CHECKBOX NAME="chkLevel3" ID="chkLevel3" tag="11X" Class="RADIO" VALUE="Y"></TD>
								</TR>								
								</TABLE>
															
							</FIELDSET>
							</TR>
						</TABLE>
					</TD>
				</TR>												
				
				<TR>
			
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>LEVEL4 원가구성상세정보 Data삭제</LEGEND>
								<TABLE WIDTH=600 <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>원가구성정보Data,대부분화면조회불가능 </TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종삭제년월</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDelYYYYMM4" SIZE=15 tag="14XXXU" ALT="최종삭제년월"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>&nbsp;&nbsp;관련:공정별실적Data,회계가공비집계조회,자품목투입정보,실제BOM출력,평가내역,단가추이,C/C,공정,품목별 원가분석,등</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종작업일</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLastDate4" SIZE=15 tag="14XXXU" ALT="최종작업일"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP><LABEL FOR="chkLevel4">선택</LABEL><INPUT TYPE=CHECKBOX NAME="chkLevel4" ID="chkLevel4" tag="11X" Class="RADIO" VALUE="Y"></TD>
								</TR>								
								</TABLE>
															
							</FIELDSET>
							</TR>
						</TABLE>
					</TD>
				</TR>			
				<TR>
			
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=20 VALIGN="top"  WIDTH="100%">
							<FIELDSET CLASS="CLSFLD">
								<LEGEND>LEVEL5 원가정보 Data삭제</LEGEND>
								<TABLE WIDTH=600 <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>원가정보,매출이익정보</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종삭제년월</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDelYYYYMM5" SIZE=15 tag="14XXXU" ALT="최종삭제년월"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD WIDTH=400 BGCOLOR="#E7E5CE" ALIGN="LEFT" CELLPADDING=5 NOWRAP>&nbsp;&nbsp;관련:품목/오더별실제원가,매출이익정보,등</TD>
									<TD WIDTH=100 BGCOLOR="#F7F7F7" ALIGN="RIGHT" CELLPADDING=5 NOWRAP>최종작업일</TD>
									<TD WIDTH=50 BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLastDate5" SIZE=15 tag="14XXXU" ALT="최종작업일"></TD>
									<TD WIDTH=50 BGCOLOR="#e6e6fa" ALIGN="RIGHT" CELLPADDING=5 NOWRAP><LABEL FOR="chkLevel5">선택</LABEL><INPUT TYPE=CHECKBOX NAME="chkLevel5" ID="chkLevel5" tag="11X" Class="RADIO" VALUE="Y"></TD>
								</TR>								
								</TABLE>
															
							</FIELDSET>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnExe()" Flag=1>실 행</BUTTON>&nbsp;</TD>
				<TD>&nbsp</TD>
				<TD>&nbsp</TD>				
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=150><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=150 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
