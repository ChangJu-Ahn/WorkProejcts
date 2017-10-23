<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : �� ���� ���� ��� 
'*  3. Program ID           : i2251oa1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*			       i21511Post Phy Inv Svr
'*			       I21119Lookup Phy inv Svr
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : LeeSeungWook
'* 10. Modifier (Last)      : LeeSeungWook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   ********************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE = VBSCRIPT>
Option Explicit                                                             

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
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
	frm1.txtMovDt.Text  = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtMovDt, parent.gDateFormat, 2)
End Sub

'------------------------------------------ OpenPlant()  --------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

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

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Tracking No."	
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackingNo.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "Tracking No."			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking_No"		
    arrHeader(1) = "ǰ��"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus	
	lgBlnFlgChgValue	  	 = True	
End Function

 '------------------------------------------  SetItemAcct()  --------------------------------------------------
'	Name : SetItemAcct()
'	Description : ItemAcct Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byRef arrRet)
	frm1.txtItemAcct.Value		= arrRet(0)
	frm1.txtItemAcctNm.Value	= arrRet(1)
	frm1.txtItemAcct.focus
	lgBlnFlgChgValue	  		= True	
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
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("10000000000011")
	Call SetDefaultVal
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtMovDt.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtMovDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMovDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtMovDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMovDt_KeyPress()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtMovDt_KeyPress(KeyAscii)
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
	Dim strYear, strMonth, strDay
	Dim var1, var2, var3, var4, var5
	Dim condvar
	Dim ObjName	 
    
    If Not chkField(Document, "1") Then				
		Exit Function
	End If

    '�����ڵ� �� ǰ������ڵ� üũ �Լ� ȣ�� 
    If Plant_Or_ItemAcct_Check = False Then 
		Exit Function
	End If

 	Call ExtractDateFrom(frm1.txtMovDt.Text,frm1.txtMovDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
   
		var1 = UCase(Trim(frm1.txtPlantCd.value))
		var2 = Trim(frm1.txtItemAcct.value)
		var3 = strYear
		var4 = strMonth
		var5 = "%" & Trim(frm1.txtTrackingNo.value) & "%"
    
	condvar = condvar & "PLANTCD|"      & var1
	condvar = condvar & "|ItemAcct|"    & var2
	condvar = condvar & "|InvYy|"       & var3
	condvar = condvar & "|InvMm|"       & var4
	condvar = condvar & "|TRACKINGNO|"  & var5
	
	ObjName = AskEBDocumentName("i2251oa1", "ebr")
	Call FncEBRprint(EBAction, ObjName, condvar)

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 
	Dim strYear, strMonth, strDay
	Dim var1, var2, var3, var4, var5
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

	Call ExtractDateFrom(frm1.txtMovDt.Text,frm1.txtMovDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

    On Error Resume Next
 	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = Trim(frm1.txtItemAcct.value)
	var3 = strYear
	var4 = strMonth
	var5 = "%" & Trim(frm1.txtTrackingNo.value) & "%"
		
	condvar = condvar & "PLANTCD|"      & var1
	condvar = condvar & "|ItemAcct|"    & var2
	condvar = condvar & "|InvYy|"       & var3
	condvar = condvar & "|InvMm|"       & var4
	condvar = condvar & "|TRACKINGNO|"  & var5
	
	ObjName = AskEBDocumentName("i2251oa1", "ebr")
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
	'-----------------------
	'Check Plant CODE		'�����ڵ尡 �ִ� �� üũ 
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

	'-----------------------
	'Check ItemAcct CODE	''ǰ������ڵ尡 �ִ� �� üũ 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����Ҵ������</font></td>
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
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=20 tag="24">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ҳ��</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i2251oa1_I537819713_txtMovDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">ǰ�����</TD>
								<TD CLASS="TD6">
								<input TYPE=TEXT NAME="txtItemAcct" SIZE="8" MAXLENGTH="2" ALT="ǰ�����" tag="12XXXU" ><IMG align=top height=20 name="btnItemAcct" onclick="vbscript:OpenItemAcct()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=20 MAXLENGTH=40 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">Tracking No.</TD>      
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()">
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
<INPUT TYPE=HIDDEN NAME="hPosSts" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemDocumentNo" tag="24" TABINDEX="-1">
</FORM>                  
<DIV ID="MousePT" NAME="MousePT">                  
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>                  
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="HIDDEN" name="uname" TABINDEX="-1">
	<input type="HIDDEN" name="dbname" TABINDEX="-1">
	<input type="HIDDEN" name="filename" TABINDEX="-1">
	<input type="HIDDEN" name="condvar" TABINDEX="-1">
	<input type="HIDDEN" name="date" TABINDEX="-1">
</FORM>                  
</BODY>                  
</HTML>
