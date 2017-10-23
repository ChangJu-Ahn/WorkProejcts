
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 원가마감작업 
'*  3. Program ID           : c3980ba1
'*  4. Program Name         : 원가마감 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
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

Const BIZ_PGM_ID = "C3980bb1.asp"

Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop          
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6


Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    
End Sub

Sub SetDefaultVal()
	On Error Resume Next
	Dim IntRetCd
	Dim CloseDt
	Dim CloseYYYYMM
	Dim strYear
	Dim strMonth
	Dim strDay


	IntRetCD = CommonQueryRs("isnull(convert(Char(10),max(closed_date),21),'')","c_close_status","close_flag =" & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	If IntRetCD = False Then
		Exit Sub	    
	Else
	    CloseDt = Trim(Replace(lgF0,Chr(11),""))
	End If

	If CloseDt <> "" Then
		Call ExtractDateFrom(CloseDt,Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

		frm1.txtCostClsDt.Year  =  strYear
		frm1.txtCostClsDt.Month =  strMonth
		Call ggoOper.FormatDate(frm1.txtCostClsDt, Parent.gDateFormat, 2)
	End if


	
	IntRetCD = CommonQueryRs("isnull(max(YYYYMM)," & FilterVar("X", "''", "S") & " )","c_close_status","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = False Then
		Exit Sub	    
	Else
	    CloseYYYYMM = Trim(Replace(lgF0,Chr(11),""))
	End If
	
	
	IF CloseYYYYMM <> "X" Then
		frm1.txtCostRefDt.Year  =  Mid(CloseYYYYMM,1,4)
		frm1.txtCostRefDt.Month =  Mid(CloseYYYYMM,5,2)
		Call ggoOper.FormatDate(frm1.txtCostRefDt, Parent.gDateFormat, 2)
	End If
	
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA")%>
End Sub

Function ExeReflect(ByVal iWhere) 
	Dim strVal
	Dim strYyyymm
	Dim	strYear, strMonth, strDay
	

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If


	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	ExeReflect = False
    
	On Error Resume Next

	select case iWhere

		case 1
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		case 2 	
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002
	end select 

	Call RunMyBizASP(MyBizASP, strVal)
	ExeReflect = True
    
End Function

Function ExeReflectOk()
	Call ggoOper.FormatDate(frm1.txtCostClsDt, Parent.gDateFormat, 2)
	Call DisplayMsgBox("990000","X","X","X")
End Function

Function ExeReflectOk1()
	Call ggoOper.FormatDate(frm1.txtCostRefDt, Parent.gDateFormat, 2)
End Function

Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables
	Call SetToolbar("10000000000000")
	Call SetDefaultVal
	frm1.txtCostClsDt.focus
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원가마감작업</font></td>
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
								<TD CLASS="TD5" NOWRAP>최종차이반영년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/c3980ba1_fpDateTime1_txtCostRefDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종마감년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/c3980ba1_fpDateTime1_txtCostClsDt.js'></script></TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSMBTN" onclick="ExeReflect(1)" Flag=1>원가마감</BUTTON>&nbsp;<BUTTON NAME="btnCancel" CLASS="CLSMBTN" onclick="ExeReflect(2)" Flag=2>원가마감취소</BUTTON></TD>
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

