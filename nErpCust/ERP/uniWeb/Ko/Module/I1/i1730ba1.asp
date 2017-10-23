
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 수불마감작업 
'*  3. Program ID           : i2232ba1
'*  4. Program Name         : 수불마감 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/12/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================  -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 
<Script Language="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "i1730bb1.asp"

Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop          
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6


Dim iJobOK

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    
End Sub

Sub SetDefaultVal()
	iJobOK = "Y"
	Call GenCloseDateInfo
End Sub

Sub GenCloseDateInfo()
	On Error Resume Next
	Dim IntRetCd
	Dim CloseDt
	Dim CloseYYYYMM
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strwhere
	Dim strClsFg
	
	strwhere =  " close_flag in (" & Filtervar("Y","''","S") & "," & Filtervar("V","''","S") & "," & Filtervar("I","''","S") & ") " 

	IntRetCD = CommonQueryRs("isnull(convert(Char(10),max(closed_date),21),'')","c_close_status",strwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If IntRetCD = False Then
		Exit Sub	    
	Else
	    CloseDt = Trim(Replace(lgF0,Chr(11),""))
	    If CloseDt = "" Then
			strClsFg = "N"
		Else
			strClsFg = "Y"
	    End If
	End If

	' 수불마감이력이 없고 재고마감이력이 있는 경우는 재고마감의 최소월이 최종마감년월이 됨 
	' 수불마감이력이 없고 재고마감이력도 없는 경우는 처리불가(재고마감을 선행해야 함)
	' 수불마감이력이 있는 경우는 최종수불마감월이 최종마감년월이 됨.
	If strClsFg = "N" Then
		IntRetCD = CommonQueryRs("convert(char(8),convert(datetime,min(b.mnth_inv_year+b.mnth_inv_month)+'01',112),112)", _ 
		                         "b_plant a left join i_inv_closing_history  b on a.plant_cd=b.plant_cd", _
		                         "b.close_create_flag='Y'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If IntRetCD = False Then
			Exit Sub	    
		Else
			CloseDt = Trim(Replace(lgF0,Chr(11),""))
			If 	CloseDt = "" Then	' 수불마감없고 , 재고마감이력 없으므로 처리불가 
				iJobOK = "N"
				Exit Sub
			Else
				iJobOK = "Y"		' 수불마감은 없어도 재고마감이력이 있으므로 처리 가능 
				frm1.txtMoveClsDt.text  = CloseDt
			End If	
		End If
	Else
		Call ExtractDateFrom(CloseDt,Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
		frm1.txtMoveClsDt.Year = strYear
		frm1.txtMoveClsDt.Month = strMonth
		frm1.txtMoveClsDt.text  =  CloseDt
	End If

	Call ggoOper.FormatDate(frm1.txtMoveClsDt, Parent.gDateFormat, 2)			
	
	'수불마감이력이 없은 경우는 공장마감의 최소값이 작업대상년월이 되고 
	'수불마감이력이 있는 경우는 최종수불마감월 +1 월이 작업대상년월이 됨 
	If 	strClsFg = "N" Then
		IntRetCD = CommonQueryRs("convert(char(8),dateadd(m,+1,convert(datetime,min(b.mnth_inv_year+b.mnth_inv_month)+'01',112)),112)", _ 
		                         "b_plant a left join i_inv_closing_history  b on a.plant_cd=b.plant_cd", _
		                         "b.close_create_flag='Y'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If IntRetCD = False Then
			frm1.txtWorkingDt.text = ""
			Exit Sub	    
		Else
		    CloseYYYYMM = Trim(Replace(lgF0,Chr(11),""))
		    frm1.txtWorkingDt.text = CloseYYYYMM
		End If	
	Else	
		frm1.txtWorkingDt.text = UNIDateAdd("M",1,CloseDt,Parent.gServerDateFormat)
	End If

	Call ggoOper.FormatDate(frm1.txtWorkingDt, Parent.gDateFormat, 2)
	
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA")%>
End Sub

Function ExeReflect(ByVal iWhere) 
	Dim strVal
	Dim strYyyymm
	Dim	strYear, strMonth, strDay
	
	On Error Resume Next
	Err.Clear 

	Dim IntRetCD
	
	If Trim(frm1.txtWorkingDt.Text) = "" Then
		IntRetCD = DisplayMsgBox("160822","X","X","X") '재고마감확정을 수행하십시오 
		Exit Function
	End If
	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
	ExeReflect = False

    Call ExtractDateFrom(frm1.txtMoveClsDt.Text,frm1.txtMoveClsDt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

	Select Case iWhere
		Case 1
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtWorkingDt=" & strYYYYMM
		Case 2 	
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002
			strVal = strVal & "&txtWorkingDt=" & strYYYYMM
	End Select 

	Call RunMyBizASP(MyBizASP, strVal)
	ExeReflect = True
    
End Function

Function ExeReflectOk()
	Call ggoOper.FormatDate(frm1.txtMoveClsDt, Parent.gDateFormat, 2)
	Call DisplayMsgBox("990000","X","X","X")
End Function

Function ExeReflectOk1()
	Call ggoOper.FormatDate(frm1.txtWorkingDt, Parent.gDateFormat, 2)
End Function

Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables
	Call SetToolbar("10000000000000")
	Call SetDefaultVal
	frm1.txtMoveClsDt.focus
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub



Function FncQuery()

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
	FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수불마감작업</font></td>
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
								<TD CLASS="TD5" NOWRAP>최종마감년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i1730ba1_fpDateTime1_txtMoveClsDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업대상년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i1730ba1_fpDateTime1_txtWorkingDt.js'></script></TD>
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
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width = 10> &nbsp </TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSMBTN" onclick="ExeReflect(1)" Flag=1>수불마감</BUTTON>&nbsp;<BUTTON NAME="btnCancel" CLASS="CLSMBTN" onclick="ExeReflect(2)" Flag=2>수불마감취소</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

