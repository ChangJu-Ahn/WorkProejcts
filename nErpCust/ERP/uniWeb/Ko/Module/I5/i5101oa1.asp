<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI ������Ȳ ��� 
'*  3. Program ID           : i5101oa1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/27
'*  8. Modified date(Last)  : 
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

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		


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
	Dim StartDate
	StartDate = UNIDateAdd("d", -6,"<%=EndDate%>", parent.gServerDateFormat)

	frm1.txtMovFrDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtMovToDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtMovFrDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtMovToDt, parent.gDateFormat, 1)
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
	
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus	
		lgBlnFlgChgValue	  	 = True
	End If	

End Function


'------------------------------------------  OpenBp()  -------------------------------------------------
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = "BP_CD in (select distinct BP_CD from B_STORAGE_LOCATION where PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & ")"		
	arrParam(5) = "����ó"			
	
	arrField(0) = "BP_CD"	
	arrField(1) = "BP_NM"	
	
	arrHeader(0) = "����ó"		
	arrHeader(1) = "����ó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)	
		frm1.txtBpCd.focus
	End If	
End Function



'------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If


	If Plant_Bp_Item_Check = False Then 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "ǰ�� �˾�"
	arrParam(1) = "B_ITEM_BY_PLANT P,B_ITEM I"	
	arrParam(2) = Trim(frm1.txtItemCd.Value)	
	arrParam(3) = ""							
	arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.MATERIAL_TYPE= " & FilterVar("30", "''", "S") & " AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "ǰ��"			
	
	arrField(0) = "I.ITEM_CD"
	arrField(1) = "I.ITEM_NM"
	
	arrHeader(0) = "ǰ���ڵ�"
	arrHeader(1) = "ǰ���"	
	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	=arrRet(0)
		frm1.txtItemNm.Value	=arrRet(1)
		frm1.txtItemCd.focus
	End If	
	
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
	Call InitComboBox
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtMovFrDt.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
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
	Dim var1, var2, var3, var4, var5
    Dim condvar 
  	Dim ObjName

	If Not chkField(Document, "1") Then	
		Exit Function
	End If
	
	If Plant_Bp_Item_Check = False Then 
		Exit Function
	End If
    
	If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
    
 
    var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UNIConvDateToYYYYMMDD(frm1.txtMovFrDt.Text,parent.gDateFormat,"")
	var3 = UNIConvDateToYYYYMMDD(frm1.txtMovToDt.Text,parent.gDateFormat,"")
    var4 = Trim(frm1.txtBPCd.Value)
    var5 = Trim(frm1.txtItemCd.value)

    
    condvar = condvar & "PLANTCD|"		& var1
    condvar = condvar & "|FromDt|"		& var2
    condvar = condvar & "|ToDt|"		& var3
    condvar = condvar & "|BPCD|"		& var4
    condvar = condvar & "|ITEMCD|"		& var5
    
	ObjName = AskEBDocumentName("i5101oa1", "ebr")
	Call FncEBRprint(EBAction, ObjName, condvar) 	

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview()
    Dim var1, var2, var3, var4, var5
    Dim condvar 
  	Dim ObjName

	If Not chkField(Document, "1") Then	
		Exit Function
	End If
	
	If Plant_Bp_Item_Check = False Then 
		Exit Function
	End If
    
	If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
    
 
    var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UNIConvDateToYYYYMMDD(frm1.txtMovFrDt.Text,parent.gDateFormat,"")
	var3 = UNIConvDateToYYYYMMDD(frm1.txtMovToDt.Text,parent.gDateFormat,"")
    var4 = Trim(frm1.txtBPCd.Value)
    var5 = Trim(frm1.txtItemCd.value)

    condvar = condvar & "PLANTCD|"		& var1
    condvar = condvar & "|FromDt|"		& var2
    condvar = condvar & "|ToDt|"		& var3
    condvar = condvar & "|BPCD|"		& var4
    condvar = condvar & "|ITEMCD|"		& var5
        
	ObjName = AskEBDocumentName("i5101oa1", "ebr")
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
' Function Name : Plant_Bp_Item_Check
' Function Desc : 
'========================================================================================
Function Plant_Bp_Item_Check()
	'-----------------------
	'Check Plant CODE	
	'-----------------------
    If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Plant_Bp_Item_Check = False
		Exit function
		
    End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	'-----------------------
	'Check BP CODE			
	'-----------------------
	If frm1.txtBPCd.value <> "" Then
		If 	CommonQueryRs(" A.BP_NM "," B_BIZ_PARTNER A, B_STORAGE_LOCATION B ", " A.BP_CD = B.BP_CD AND B.BP_CD = " & FilterVar(frm1.txtBPCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If 	CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", " BP_CD = " & FilterVar(frm1.txtBPCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
						
				Call DisplayMsgBox("229927","X","X","X")
				frm1.txtBPNm.Value = ""
				frm1.txtBPCd.focus 
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtBPNm.Value = lgF0(0)

			Call DisplayMsgBox("162064","X","X","X")
			frm1.txtBPCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtBPNm.Value = lgF0(0)
	End If

	'-----------------------
	'Check Item CODE	
	'-----------------------
    If frm1.txtItemCd.Value <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
		Else
			frm1.txtItemNm.Value = ""
		End If
	End If

    Plant_Bp_Item_Check = True
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>VMI ������Ȳ���</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=20 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i5101oa1_I102438020_txtMovFrDt.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/i5101oa1_I387102526_txtMovToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=8 MAXLENGTH=10 tag="11XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpPopup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=30 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCd" SIZE="18" MAXLENGTH="18" ALT="ǰ��" tag="11XXXU" ><IMG align=top height=20 name="btnItemNm" onclick="vbscript:OpenItemCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="14">
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>                  
		</TD>                  
	</TR>                  
</TABLE>                  
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">                  
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">                  
<INPUT TYPE=HIDDEN NAME="hPosSts" tag="24">                  
<INPUT TYPE=HIDDEN NAME="hItemDocumentNo" tag="24">                  
</FORM>                  
<DIV ID="MousePT" NAME="MousePT">                  
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>                  
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
</FORM>
</BODY>                  
</HTML>
