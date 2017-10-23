<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : ���������� ������ ��� 
'*  3. Program ID           : i1821oa1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/04/30
'*  8. Modified date(Last)  : 2002/05/20
'*  9. Modifier (First)     : LeeSeungWook
'* 10. Modifier (Last)      : LeeSeungWook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2001/09/12 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->	
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<!--'==========================================  1.1.2 ���� Include   =====================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE = VBSCRIPT>
Option Explicit         

'==========================================  1.2.1 Global ��� ����  ======================================
<%
EndDate   = GetSvrDate
%>

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim hPosSts
Dim IsOpenPop          

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE  
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
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "OA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim StartDate
	StartDate = UNIDateAdd("m", -1,"<%=EndDate%>", parent.gServerDateFormat)
	frm1.txtMovFrDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtMovToDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtMovFrDt, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtMovToDt, parent.gDateFormat, 2)
	frm1.txtPlantCd.focus

End Sub

'--------------------------------------------------------------------------------------------------------- 
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "����"		
	arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	

End Function

 '------------------------------------------  OpenItemAcct()  --------------------------------------------------
'	Name : OpenItemAcct()
'	Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "ǰ����� �˾�"	
	arrParam(1) = "B_MINOR"				
	arrParam(2) = Trim(frm1.txtItemAcct.Value)
	arrParam(3) = ""						
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""		
	arrParam(5) = "ǰ�����"			
	
	arrField(0) = "MINOR_CD"		
	arrField(1) = "MINOR_NM"		
	
	arrHeader(0) = "ǰ�����"	
	arrHeader(1) = "ǰ�������"	
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemAcct.focus
		Exit Function
	Else
		Call SetItemAcct(arrRet)
	End If	
	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus	
	lgBlnFlgChgValue	  	 = True	
End Function

 '------------------------------------------  SetItemAcct()  --------------------------------------------------
'	Name : SetItemAcct()
'	Description : ItemAcct Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byval arrRet)
	frm1.txtItemAcct.Value	    =arrRet(0)
	frm1.txtItemAcctNm.Value	=arrRet(1)
	frm1.txtItemAcct.focus
	Set gActiveElement = document.activeElement
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call InitVariables		
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)		
	Call ggoOper.LockField(Document, "N")			
	Call SetToolbar("10000000000011")
	Call SetDefaultVal
End Sub

'=======================================================================================================
'   Event Name : txtMovFrDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMovFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtMovFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMovToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMovToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtMovToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMovFrDt_KeyPress()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtMovToDt_KeyPress()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Call BtnPreview()
    FncQuery = True       
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
    Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
	Dim var1, var2, var3, var4, var5, var6
    Dim condvar 
   	Dim ObjName
 
	If Not chkField(Document, "1") Then	
		Exit Function
	End If
	
	If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
    
    If Plant_Or_ItemAcct_Check = False Then 
		Exit Function
	End If

    If Trim(frm1.txtPlantCd.value) = "" Then
		var1 = "%"
	Else 
		var1 = UCase((Trim(frm1.txtPlantCd.value) & "%"))
	End if 
   
    If Trim(frm1.txtItemAcct.value) = "" Then
		var2 = "%"
	Else 
		var2 = (Trim(frm1.txtItemAcct.value) & "%")
	End if

 	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
 
    var3 = strYear
    var4 = strMonth
    var5 = strYear1
    var6 = strMonth1
    
    condvar = condvar & "PLANTCD|"    & var1
    condvar = condvar & "|ITEMACCT|" & var2
    condvar = condvar & "|FromYy|"    & var3
    condvar = condvar & "|FromMm|"    & var4
    condvar = condvar & "|ToYy|"      & var5
    condvar = condvar & "|ToMm|"      & var6

	ObjName = AskEBDocumentName("i1821oa1", "ebr")
	Call FncEBRprint(EBAction, ObjName, condvar) 	

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()
	Dim var1, var2, var3, var4, var5, var6
    Dim condvar 
	Dim ObjName

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then       
       Exit Function
    End If
    
    If Plant_Or_ItemAcct_Check = False Then 
		Exit Function
	End If

    If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
	
    If Plant_Or_ItemAcct_Check = False Then 
		Exit Function
	End If

    If Trim(frm1.txtPlantCd.value) = "" Then
		var1 = "%"
	Else 
		var1 = UCase((Trim(frm1.txtPlantCd.value) & "%"))
	End if 
    
    If Trim(frm1.txtItemAcct.value) = "" Then
		var2 = "%"
	Else 
		var2 = (Trim(frm1.txtItemAcct.value) & "%")
	End if
  
    var3 = frm1.txtMovFrDt.Year
    var4 = right("0" & frm1.txtMovFrDt.Month, 2)
    var5 = frm1.txtMovToDt.Year
    var6 = right("0" & frm1.txtMovToDt.Month, 2)
    
    
	condvar = condvar & "PLANTCD|"      & var1
    condvar = condvar & "|ITEMACCT|"    & var2
    condvar = condvar & "|FromYy|"      & var3
    condvar = condvar & "|FromMm|"      & var4
    condvar = condvar & "|ToYy|"        & var5
    condvar = condvar & "|ToMm|"        & var6

	ObjName = AskEBDocumentName("i1821oa1", "ebr")
	Call FncEBRPreview(ObjName, condvar)    
	
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
    Call parent.FncExport(parent.C_SINGLE)        
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , True)        
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : Plant_Or_ItemAcct_Check
' Function Desc : 
'========================================================================================
Function Plant_Or_ItemAcct_Check()
	
	If Trim(frm1.txtPlantCd.Value)	<> "" Then
		'-----------------------
		'Check Plant CODE
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus
			Plant_Or_ItemAcct_Check = False
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	End If

	If Trim(frm1.txtItemAcct.Value)	<> "" Then
		'-----------------------
		'Check ItemAcct CODE
		'-----------------------
		If 	CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD= " & FilterVar(frm1.txtItemAcct.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("169952","X","X","X")
			frm1.txtItemAcctNm.Value = ""
			frm1.txtItemAcct.focus
			Plant_Or_ItemAcct_Check = False
			Exit function
		End If
        lgF0 = Split(lgF0, Chr(11))
		frm1.txtItemAcctNm.Value = lgF0(0)
	End If 

    Plant_Or_ItemAcct_Check = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������� ������</font></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%> >
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="11XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=20 tag="14" ALT="�����">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ұⰣ</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i1821oa1_I204729908_txtMovFrDt.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/i1821oa1_I573897516_txtMovToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">ǰ�����</TD>
								<TD CLASS="TD6">
								<input TYPE=TEXT NAME="txtItemAcct" SIZE="8" MAXLENGTH="2" tag="11XXXU" ALT="ǰ�����"  ><IMG align=top height=20 name="btnItemAcct" onclick="vbscript:OpenItemAcct()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 MAXLENGTH=40 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5"></TD>
								<TD CLASS="TD6"></TD>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreView()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>                  
				</TR>
			</TABLE>
		</TD>
	</TR>                  
	<TR>                  
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>                  
		</TD>                  
	</TR>                  
</TABLE>                  
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">                  
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">                  
<INPUT TYPE=HIDDEN NAME="hPosSts" tag="24" TABINDEX="-1">                  
<INPUT TYPE=HIDDEN NAME="hItemDocumentNo" tag="24" TABINDEX="-1">                  
</FORM>                  
<DIV ID="MousePT" NAME="MousePT">                  
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>                  
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">                 
</FORM>
</BODY>                  
</HTML>
  
