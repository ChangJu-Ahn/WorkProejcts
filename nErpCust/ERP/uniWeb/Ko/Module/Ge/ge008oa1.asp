
<%@ LANGUAGE="VBSCRIPT" %>
<!--
**********************************************************************************************
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        : 
*  3. Program ID           : GC007OA1
*  4. Program Name         : ǰ��׷� ��������ǥ��� 
*  5. Program Desc         : ǰ��׷� ��������ǥ��� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/05
*  8. Modified date(Last)  : 2001/12/05
*  9. Modifier (First)     : Lee Kang Yeong
* 10. Modifier (Last)      : Lee Kang Yeong
* 11. Comment      :
* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
*                            this mark(��) Means that "may  change"
*                            this mark(��) Means that "must change"
* 13. History              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                        '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim IsOpenPop
Dim lgOldRow

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
         
End Sub

'========================================================================================================
Sub SetDefaultVal()
	Dim StartDate
	StartDate = "<%=GetSvrDate%>"
	
	frm1.txtYyyymm.focus()    
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 3) 
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("P", "G", "NOCOOKIE", "PA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables 
    
'    Call ggoOper.FormatDate(frm1.txtpay_yymm, Parent.gDateFormat, 2)                    '�̱ۿ��� ����� �Է��ϰ� ������� ���� �Լ��� ���Ѵ�.
    
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")

End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================
Function FncQuery()
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    FncQuery = True                                                              '��: Processing is OK

End Function

'========================================================================================
Function txtGrade_onKeyPress(Key)    
    
    frm1.action = "../../blank.htm"       
    
End Function
	
'=======================================================================================================
Function FncBtnPrint() 
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim var1,var2
    Dim strYear, strMonth, strDay
    	
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If
	
	StrEbrFile = "ge008oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	
	
    <%'--��������� �����ϴ� �κ� ���� %>
	
	condvar = "YYYY|" & var1
	condvar = condvar & "|ITEM_GROUP_CD|" & var2	
	
	Call FncEBRPrint(EBAction,ObjName,condvar)				

End Function

'========================================================================================
Function FncBtnPreview()
'On Error Resume Next                                                    '��: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'��: This function check indispensable field%>
       Exit Function
    End If
	
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim var1, var2	
    Dim strYear, strMonth, strDay

	StrEbrFile = "ge008oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP "," ITEM_GROUP_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	

	condvar = "YYYY|" & var1
	condvar = condvar & "|ITEM_GROUP_CD|" & var2

	Call FncEBRPreview(ObjName,condvar)

End Function

'========================================================================================================
Function FncPrint()
	Call FncPrint()                                                      '��: Protect system from crashing
End Function

'========================================================================================================
Function FncFind() 
	Call FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
Function OpenPopUp()
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


			arrParam(0) = "ǰ��׷�"					' �˾� ��Ī 
			arrParam(1) = "B_ITEM_GROUP"						' TABLE ��Ī 
			arrParam(2) = UCase(Trim(frm1.txtFr_dept_cd.Value))	' Code Condition
			arrParam(3) = ""							' Name Cindition
			'arrParam(4) = ""	
			arrParam(5) = "ǰ��׷��ڵ�"			
	
   			arrField(0) = "ITEM_GROUP_CD"	     				' Field��(0)
			arrField(1) = "ITEM_GROUP_NM"			    		' Field��(1)
		
			arrHeader(0) = "ǰ��׷�"					' Header��(0)
			arrHeader(1) = "ǰ��׷��"				' Header��(1)
    
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	   Frm1.txtFr_dept_cd.value = arrRet(0)
	   frm1.txtFr_dept_nm.value = arrRet(1)
	End If	

End Function

'========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
	End If
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ��׷캰 ��������ǥ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD CLASS=TD5  NOWRAP>�����</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/ge008oa1_fpDateTime1_txtYyyymm.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="ǰ��׷��ڵ�" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp()">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="ǰ��׷��" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU"></TD>
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
	<TR HEIGHT=20>
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>�μ�</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME type=hidden NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

