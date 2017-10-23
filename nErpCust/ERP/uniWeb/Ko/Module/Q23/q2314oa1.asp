<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2314OA1
'*  4. Program Name         : �̷�ī����� 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit																		'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim strInspItemCd
Dim strInspItemCd_temp
Dim strPrintFlag
Dim strLen

Dim IsOpenPop          

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID	= "..\Q21\q2114ob1.asp"                         '��: �����Ͻ� ���� ASP�� 
'--------------- ������ coding part(��������,End)-----------------------------------------------------------

 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboInspClassCd.value		= "F"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","OA") %>
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))
End Sub

 '------------------------------------------  OpenPlant() -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
	
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

    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End If	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

 '------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "ǰ��"																	' �˾� ��Ī 
	arrParam(1) = "B_Item_By_Plant,B_Item"												' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)													' Code Condition
	arrParam(3) = ""												' Name Condition
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
	arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " " 			' Where Condition
	arrParam(5) = "ǰ��"																	' TextBox ��Ī 
	
	arrField(0) = "B_Item_By_Plant.Item_Cd"					' Field��(0)
	arrField(1) = "B_Item.Item_NM"				' Field��(1)
	arrField(2) = "B_Item.SPEC"					' Field��(2)
	
	arrHeader(0) = "ǰ���ڵ�"						' Header��(0)
	arrHeader(1) = "ǰ���"					' Header��(1)
	arrHeader(2) = "�԰�"						' Header��(2)
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
	End If	
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement
	OpenItem = true	
End Function

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	
	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables                                                      '��: Initializes local global variables
	
	 '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetDefaultVal
	Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtPlantCd.focus 
    Else
		frm1.txtPlantCd.focus 
    End If

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery()
	FncQuery = true
	If FncBtnPreview = False Then Exit function               '�̸����� Call
	
	Call CancelRestoreToolBar()
	FncQuery = true
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = false
	
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
    
    FncFind = true
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = false 
	
	Call Parent.FncPrint()
	
	FncPrint = true
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is related to Print Button
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim strHeaderList
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------------------------------------
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
							& "&txtInspClassCd=" & Trim(.cboInspClassCd.value) _
							& "&txtItemCd=" & Trim(.txtItemCd.value)
	'--------------- ������ coding part(�������,End)------------------------------------------------------------------------------
        
	    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()															'��: ��ȸ ������ ������� 
		

	DbQueryOk = false

	dim var1, var2, var3, var4, var5, var6, var7, var8
	dim condvar

	Dim strEbrFile
	Dim objName
	Dim strUrl
	
	If Not chkField(Document, "1") Then	Exit Function
		
	var1 = Trim(frm1.txtPlantCd.value)

	var2 = Trim(frm1.cboInspClassCd.value)

	var3 = Trim(frm1.txtItemCd.value)

	strEbrFile = "Q2314OA1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	strLen = len(strInspItemCd)
	strInspItemCd_temp = left(strInspItemCd,strLen-1)

		
	strUrl = strUrl & "PlantCd|" & var1 
	strUrl = strUrl & "|InspClassCd|" & var2 
	strUrl = strUrl & "|ItemCd|" & var3
	strUrl = strUrl & "|" & strInspItemCd_temp

	
	If strPrintFlag = "P" Then
		Call FncEBRprint(EBAction, objName, strUrl)
	Else
		call FncEBRPreview(objName, strUrl)
	End if

	
	DbQueryOk = true
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
	FncBtnPrint = false
	
	If Not chkField(Document, "1") Then	Exit Function

	If Plant_Item_Check = False Then Exit Function

	strPrintFlag = "P"
	If DbQuery = false then Exit Function

	FncBtnPrint = true	
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	FncBtnPreview = false
		
	If Not chkField(Document, "1") Then	Exit Function

	If Plant_Item_Check = False Then Exit Function
	
	strPrintFlag = "V"
	
	If DbQuery = false then Exit Function
	
	FncBtnPreview = true
End Function


'========================================================================================
' Function Name : Plant_Item_Check
'========================================================================================
Function Plant_Item_Check()

	Plant_Item_Check = False
	
	With frm1
	
 		If  CommonQueryRs(" B.ITEM_NM, C.PLANT_NM "," B_ITEM_BY_PLANT A, B_ITEM B, B_PLANT C ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("125000","X","X","X")
				.txtPlantNm.Value = ""
				.txtPlantCd.focus 
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			.txtPlantNm.Value = lgF0(0)


			If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
				Call DisplayMsgBox("122600","X","X","X")
				.txtItemNm.Value = ""
				.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				lgF0 = Split(lgF0, Chr(11))
				.txtItemNm.Value = lgF0(0)
				Call DisplayMsgBox("122700","X","X","X")
				.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If

 		lgF0 = Split(lgF0, Chr(11))
 		lgF1 = Split(lgF1, Chr(11))
		.txtPlantNm.Value = lgF1(0)
		.txtItemNm.Value = lgF0(0)
	End With 
	Plant_Item_Check = True
End Function
'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�̷�ī��</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" HEIGHT=* colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 STYLE="HEIGHT: 100%">	
	    		<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 100%">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0 STYLE="HEIGHT: 100%">
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantCd ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X"></TD>
                                    				</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�˻�з�</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="�˻�з�" STYLE="WIDTH: 150px" tag="14"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="ǰ��" tag="12XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
												<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	    <TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
				    	<BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;              
		                <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>                
		            </TD>                  
				</TR>
			</TABLE>
		</TD>  
	</TR>                        
	<TR>                        
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm "  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>                        
		</TD>                        
	</TR>                        
</TABLE>                        
</FORM>                        
<DIV ID="MousePT" NAME="MousePT">                        
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>                        
</DIV>                        
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST"> 
    <input type="hidden" name="uname" tabindex=-1>
    <input type="hidden" name="dbname" tabindex=-1>
    <input type="hidden" name="filename" tabindex=-1>
    <input type="hidden" name="condvar" tabindex=-1>
	<input type="hidden" name="date" tabindex=-1>
</FORM> 
</BODY>                        
</HTML>

