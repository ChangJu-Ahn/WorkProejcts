<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : m6112ma2
'*  4. Program Name         : 부대비일괄배부 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/09/13
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Byun Jee Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->       
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             

Const BIZ_PGM_ID   = "m6112mb2.asp"

'Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
'Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim IsOpenPop          
'Dim lgCboKeyPress      
'Dim lgplantchk
'Dim lgDbSaveOkOccurFlag

<!-- #Include file="../../inc/lgvariables.inc" -->

'==============================================================================================================================
Sub InitVariables()

IF Parent.gPlant = "" Then 
	Call SetToolbar("1000000000001111")	
Else
	Call SetToolbar("1000100000001111")	
End If

    lgIntFlgMode = parent.OPMD_CMODE                                            
    lgBlnFlgChgValue = False                                                
    
    IsOpenPop = False														
    'lgCboKeyPress = False
    'lgDbSaveOkOccurFlag	=	False
    
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub SetDefaultVal()									'⊙: 버튼 툴바 제어 
    frm1.txtPlantCd.value=Parent.gPlant
	frm1.txtPlantNm.value=Parent.gPlantNm
    frm1.txtPlantCd.focus
	
	Call ggoOper.FormatDate(frm1.txtDisbEndDt, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtFrDisbQryDt, Parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtToDisbQryDt, Parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtDisbDt, Parent.gDateFormat, 2)
    Set gActiveElement = document.activeElement
    If frm1.txtPlantCd.value <> "" then 
		Call checkflg()
	End if
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub
'==============================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CONVERT(VARCHAR(10),INV_CLS_DT)"

    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    arrHeader(2) = "작업년월"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)	
		'lgDbSaveOkOccurFlag	=	false
		Call checkflg()
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenStep()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "경비발생단계"					
	arrParam(1) = "B_minor"						
	arrParam(2) = frm1.txtStep.value	
	arrParam(3) = ""							
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""			
	arrParam(5) = "경비발생단계"			
	
    arrField(0) = "minor_cd"					
    arrField(1) = "minor_nm"					
    
    arrHeader(0) = "경비발생단계"				
    arrHeader(1) = "경비발생단계명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtStep.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtStep.value		= arrRet(0)
		frm1.txtStepNm.value	= arrRet(1)
		frm1.txtStep.focus
		Set gActiveElement = document.activeElement
	End If
	
	if frm1.txtPlantCd <> "" then
		Call checkflg()	
	end if

End Function
'=================================================================================
Sub checkflg()

	Dim iCodeArr, iNameArr, iRateArr
    Dim plant_cd, process_step
    Dim strdt,stwhere, strdt1
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Err.Clear
    
    plant_cd = frm1.txtPlantCd.value
    process_step = frm1.txtStep.value
	
	stwhere = "c.plant_cd = " & FilterVar(plant_cd, "''", "S") & " and b.process_step like " & FilterVar(process_step, "'%'", "S") 
	stwhere = stwhere & " AND A.BATCH_FLG = " & FilterVar("Y", "''", "S") 
	stwhere = stwhere & " AND (A.DISB_TYPE = " & FilterVar("B", "''", "S") & " OR A.DISB_TYPE = " & FilterVar("", "''", "S") & ") "
	
	Call CommonQueryRs(" isnull(max(batch_job_dt), '1900-01-01') "," m_purchase_expense_by_gm a inner join m_purchase_charge b on (a.charge_no = b.charge_no) inner join m_purchase_expense_by_item c on (a.charge_no = c.charge_no and a.seq = c.seq) ",stwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Sub
	End If

   if iCodeArr(0) = "1900-01-01" then
		strdt = "<%=GetSvrDate%>"
		frm1.txtDisbEndDt.text    = ""
   else 
       strdt = iCodeArr(0)
       frm1.txtDisbEndDt.text    = UniConvDateAToB(strdt,Parent.gDateFormat,Parent.gDateFormatYYYYMM)     
   end if 
   
   strdt1 =  uniDateAdd("M", 1, strdt, parent.gDateFormat)
   
   frm1.txtFrDisbQryDt.text  = UniConvDateAToB(strdt,Parent.gDateFormat,Parent.gDateFormatYYYYMM)    
   frm1.txtToDisbQryDt.text  = UniConvDateAToB(strdt1,Parent.gDateFormat,Parent.gDateFormatYYYYMM)      
   frm1.txtDisbDt.text	 = UniConvDateAToB(strdt1,Parent.gDateFormat,Parent.gDateFormatYYYYMM)	
   
   Call SetToolbar("1000100000001111")	
	
End Sub

'==============================================================================================================================
Sub Form_Load()

    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart, _
							Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    '----------  Coding part  -------------------------------------------------------------
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables	    
End Sub
'==============================================================================================================================
Sub txtFrDisbQryDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrDisbQryDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtFrDisbQryDt.Focus
    End If
End Sub
'==============================================================================================================================
Sub txtToDisbQryDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDisbQryDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtToDisbQryDt.Focus
    End If
End Sub
'=============================================================================
Sub txtDisbDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDisbDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDisbDt.Focus
    End If
End Sub
'=============================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False                                                       
    
    Err.Clear		                                                   

	If Not chkField(Document, "2") Then                          
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	If CheckRunningBizProcess = True Then
		Exit Function
	End If

    If DbSave = False Then Exit Function            
    
    FncSave = True    
    Set gActiveElement = document.activeElement 
        
End Function
'==============================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
    Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
    Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO_CANCEL,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
    Set gActiveElement = document.activeElement 
End Function
'==============================================================================================================================
Function DbSave() 
	
    'Err.Clear																<%'☜: Protect system from crashing%>

    DbSave = False		
    
    If LayerShowHide(1) = False Then Exit Function													<%'⊙: Processing is NG%>
        
    With frm1
		.txtMode.value = parent.UID_M0002	
		.txtFlgMode.value = lgIntFlgMode	
		 Call LayerShowHide(1) 
		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)									
	End With
	
    DbSave = True   
End Function
'==============================================================================================================================
Function DbSaveOk()	
	Call DisplayMsgBox("990000", "X", "X", "X")													<%'☆: 저장 성공후 실행 로직 %>
	Call InitVariables
	call checkflg()
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% >
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>부대비일괄배부</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" ONChange="vbscript:checkflg()" tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="공장" tag="14X">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>배부참조번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDistRefNo" SIZE=20 MAXLENGTH=18 ALT="배부참조번호"  tag="24XXXU">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>경비발생단계</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtStep" SIZE=10 MAXLENGTH=5 ALT="경비발생단계" ONChange="vbscript:checkflg()" tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenStep()">
													   <INPUT TYPE=TEXT NAME="txtStepNm" SIZE=20 MAXLENGTH=20 ALT="경비발생단계" tag="14X">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종배부년월</TD>
								<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=txtDisbEndDt title=FPDATETIME CLASSID=<%=gCLSIDFPDT%> CLASS=FPDTYYYYMM name=txtDisbEndDt tag="14N1" ALT="최종배부년월"></OBJECT>');</SCRIPT>
								</TD>	
							</TR>				
							<TR>
								<TD CLASS="TD5" NOWRAP>배부대상기간</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
									<tr>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrDisbQryDt name=txtFrDisbQryDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="배부대상기간" tag="12XXXU"><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</SCRIPT>
										</td>
										<TD> ~ </TD>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtToDisbQryDt name=txtToDisbQryDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="배부대상기간" tag="12XXXU"><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</SCRIPT>
										</td>
									</tr>
									</table>
								</TD>								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>배부년월</TD>
								<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtDisbDt name=txtDisbDt CLASS=FPDTYYYYMM title=FPDATETIME ALT="배부년월" tag="12XXXU"><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
