<%@ LANGUAGE="VBSCRIPT" %>

<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>����� �޴� �߰�</TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* %>

<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>

<%'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================%>
<!-- #Include file="../inc/uni2kcm.inc" -->
<!-- #Include file="../inc/IncServer.asp" -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/EventPopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/ccm.vbs">       </Script>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incUni2KTV.vbs"></Script>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js">   </SCRIPT>



<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_ID          = "BizAddMenu.asp"		
Const BIZ_PGM_USERMENU_ID = "LoadUserMenu.asp"		

Const C_Sep           = "/"

Const C_IMG_Folder    = "../../CShared/image/CloseDir.gif"
Const C_IMG_Open      = "../../CShared/image/OpenDir.gif"
Const C_IMG_URL       = "../../CShared/image/Program.gif"
Const C_IMG_None      = "../../CShared/image/Program_d.gif"
Const C_IMG_Const     = "../../CShared/image/c_const.gif"

Const C_MNU_SEP       = "::"
Const C_MNU_ID        = 0
Const C_MNU_UPPER     = 1
Const C_MNU_LVL       = 2
Const C_MNU_TYPE      = 3
Const C_MNU_NM        = 4
Const C_MNU_AUTH      = 5

Const C_USER_MENU     = "����� �޴�"
Const C_USER_MENU_KEY = "*"
Const C_USER_MENU_STR = "UM_"
Const C_UNDERBAR      = "_"

Const C_NEW_FOLDER    = "�� ����"

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnBizLoadMenu
Dim StrVal
Dim OpMode
Dim arrParam
Dim strPgName
Dim strURL
Dim XXXX   'khy200304

arrParam  = window.dialogArguments

set XXXX = arrParam(0)'khy200304
strPgName = arrParam(1)
strURL    = arrParam(2)

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgBlnBizLoadMenu = False
'--------------------------------------------  Coding part  -----------------------------------------------
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************** 
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim NodX
	Set NodX = frm2.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, C_USER_MENU, C_USFolder, C_USFolder)
	NodX.ExpandedImage = C_USOpen
	NodX.Tag = 0			
End Sub
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'   ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'==========================================================================================
'   Function Name : GetIndex
'   Function Desc : ���� �ԷµǴ� Node�� ���° ������ ��ȯ�Ѵ�.
'==========================================================================================
Function GetIndex(Node)
    
    If Node.Key = gDragNode.Key Then
        GetIndex = gDropNode.Children + 1
        Exit Function
    End If

	GetIndex = GetSeq(Node)    

End Function

'==========================================================================================
'   Function Name : GetNodeLvl
'   Function Desc : ���� ����� Level�� ã�´�.
'==========================================================================================
Function GetNodeLvl(Node)
    Dim tempNode
    
    Set tempNode = Node
    
    GetNodeLvl = 0
    
    if tempNode.Key <> "*" Then
	    Do    	
    		GetNodeLvl = GetNodeLvl + 1
    		Set tempNode = tempNode.Parent
    	Loop Until tempNode.Key = "*"
	End If
	Set tempNode = Nothing
End Function

'======================================================================================================
'	�޴��� �о� TreeView�� ���� 
'======================================================================================================
Sub DisplayBizMenu()
	Dim strVal

	Call SetDefaultVal()
		
	frm2.uniTree1.MousePointer = 11
	
	strVal = BIZ_PGM_USERMENU_ID & "?txtKey=$"								'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtUKey=$"											'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(PopBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
	strVal = ""

End Sub

'========================================================================================
' Function Name : GetImage
' Function Desc : �̹��� ���� 
'========================================================================================
Function GetImage(Byval arrLine)
	Dim strImg
	Select Case arrLine(C_MNU_AUTH)
		Case "A"
			If arrLine(C_MNU_TYPE) = "M" Then
				strImg = C_USFolder
			Else
				strImg = C_USURL
			End If
		Case "I"
			strImg = C_USConst
		Case "N"
			strImg = C_USNone
	End Select
	GetImage = strImg
End Function

'==========================================================================================
'   Function Name : GetMenuType
'   Function Desc : �ԷµǴ� Node�� �޴�Ÿ���� ��ȯ�Ѵ�.
'==========================================================================================
Function GetMenuType(imgStyle)
    If imgStyle = C_USURL Then
        getMenuType = "P"
    Else
        getMenuType = "M"
    End If
End Function

'==========================================================================================
'   Function Name : GetPgID
'   Function Desc : ���� ȭ���� ID�� ��´�.
'==========================================================================================
Function getPgID()

'	Dim tmpPgID
	
'	tmpPgID = strURL
	
'	If InStrRev(tmpPgID, ".") > 0 Then tmpPgID = Left(tmpPgID, InStrRev(tmpPgID, ".") - 1)
'	If InStrRev(tmpPgID, "/") > 0 Then tmpPgID = Mid(tmpPgID, InStrRev(tmpPgID, "/") + 1)
	
'	getPgID = tmpPgID

	getPgID = strURL

End Function


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'==========================================================================================
'   Event Name : uniTree1_AfterLabelEdit
'   Event Desc : Add�ϰ� Label�� �������� DB����� ȣ���� �̺�Ʈ 
'==========================================================================================
Sub uniTree1_AfterLabelEdit(Cancel, NewString)

	Dim NodX
	
	Set NodX = frm2.uniTree1.SelectedItem
	
	StrVal = StrVal & "U"                     & gColSep 
	StrVal = StrVal & NodX.Key                & gColSep 
	StrVal = StrVal & NodX.Parent.Key         & gColSep 
	StrVal = StrVal & NewString               & gColSep 
	StrVal = StrVal & getMenuType(NodX.Image) & gColSep 
	StrVal = StrVal & GetNodeLvl(NodX)        & gColSep 
	StrVal = StrVal & NodX.Parent.Tag         & gColSep 
	StrVal = StrVal & NodX.Key                & gColSep 
	StrVal = StrVal & NodX.Parent.Key         & gRowSep
	Set NodX = Nothing	
	
	
End Sub
'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node�� Ŭ���ϸ� �߻� �̺�Ʈ 
'==========================================================================================
Sub uniTree1_NodeClick(Node)
	if Node.Key = "*" Then
		frm2.uniTree1.LabelEdit = 1
	Else
		frm2.uniTree1.LabelEdit = 0
	End If				
End Sub

'======================================  uniTree1_onAddImgReady()  ====================================
'	Name : uniTree1_onAddImgReady()
'	Description : SetAddImageCount���� Image�� �ٿ�ε� �Ϸ�ǰ� TreeView�� ImageList�� �߰��Ǹ� �߻��ϴ� �̺�Ʈ 
'========================================================================================================= 

Sub uniTree1_onAddImgReady()
	Call DisplayBizMenu()
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
	
	Dim NodX, lHwnd
	Dim IntRetCD
	
	If Trim(getPgID()) = "" Or Trim(getPgId()) = "N/A" Then 'khy200304
		IntRetCD = DisplayMsgBox("211429", "x","x","x")			    '���� ������ ����� �޴��� ����Ҽ� �����ϴ�.    			
    	Self.Returnvalue = ""
  		Self.Close()      			
		Exit Sub		
	End if
	
	With frm2
		.uniTree1.SetAddImageCount = 5
		.uniTree1.Indentation      = "200"	                                            ' �� ���� 
		.uniTree1.AddImage C_IMG_Folder, C_USFolder, 0									'��: TreeView�� ���� �̹��� ���� 
		.uniTree1.AddImage C_IMG_Open  , C_USOpen  , 0
		.uniTree1.AddImage C_IMG_URL   , C_USURL   , 0
		.uniTree1.AddImage C_IMG_None  , C_USNone  , 0
		.uniTree1.AddImage C_IMG_Const , C_USConst , 0
		
		.uniTree1.OLEDragMode = 0														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
		.uniTree1.OLEDropMode = 0

		.txtName.Value = strPgName
	End With
	
	Call InitVariables()
	
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************

'======================================================================================================
'	��� ��ư Ŭ�� 
'======================================================================================================

Function btnCancel_onclick()
	Self.Returnvalue = ""
	Self.Close()
End Function

'======================================================================================================
'	Ȯ�� ��ư Ŭ���� 
'======================================================================================================

Function btnOK_onclick()

	Dim NodX	
	Dim strTmp1, strTmp2, tmp_cnt
	Dim IntRetCD
	Dim NodeCnt
	
	' �������� ������� ����� New folder�� ������ ���� 
	strTmp1 = split(StrVal,gRowSep)
	StrVal = ""
  
	For tmp_cnt = 0 to ubound(strTmp1,1) - 1
	    strTmp2 = split(strTmp1(tmp_cnt),gColSep)
	    If Trim(strTmp2(3)) > "" then
	        StrVal = StrVal & strTmp1(tmp_cnt) & gRowSep
	    End If
	Next


	If frm2.txtName.value ="" Then 		
		IntRetCD = DisplayMsgBox("211428", "x","x","x")			    '���α׷����� �Է��ϼ���.
    			If IntRetCD = vbNo Then
      				Exit Function
    			End If	
		Exit function
	End If 

	With frm2
		Set NodX = .uniTree1.SelectedItem		 
		NodeCnt = split(NodX.fullpath,"\")               'khy
		StrVal = StrVal & "C"                  & gColSep 
		StrVal = StrVal & getPgID()            & gColSep 
		StrVal = StrVal & NodX.Key             & gColSep 
		StrVal = StrVal & .txtName.Value       & gColSep 
		StrVal = StrVal & "P"                  & gColSep  
		StrVal = StrVal & UBound(NodeCnt) + 1  & gColSep 'khy
		StrVal = StrVal & (Cint(NodX.Tag) + 1) & gColSep 
		StrVal = StrVal & ""                   & gColSep 
		StrVal = StrVal & ""                   & gRowSep

	End With


	frm2.uniTree1.MousePointer = 11
	
	'khy200304
	Dim StrRec
	Dim StrTmp
	Dim CntRec
	Dim i 
	Dim ArrRec
	
	StrTmp = Split(Strval,chr(12))
	
	CntRec = UBound(strtmp)
		
	For i = 0 To CntRec-1
		ArrRec= Split(StrTmp(i),chr(11))		
		Call XXXX.AddMenuPopup (ArrRec(0),ArrRec(1),ArrRec(2),ArrRec(3),ArrRec(4),ArrRec(5),ArrRec(6))
    	
	Next
	
	Call XXXX.RefreshUsrXml()
	
	Call DbSaveOk
	
	strval =""
	 
	Set NodX = Nothing	

End Function


'======================================================================================================
'	������ ��ư Ŭ�� 
'======================================================================================================

Function btnNewFolder_onclick()

	Dim NodX, Node, tmpCnt, strNewKey
	
	Set Node = frm2.uniTree1.SelectedItem
	
	tmpCnt = 0
	
	On Error Resume Next

	'��ȣ�� ���� 2002.09.13 �޴����� ����.
	'If (GetNodeLvl(Node) + 1) > 10 then
	'	IntRetCD = DisplayMsgBox("211437", "x","x","x")			    '����� �޴��� 10������ �ʰ��� �� �����ϴ�.
	'	exit function
	'End If
	
	Do
		strNewKey = C_USER_MENU_STR & cStr(tmpCnt)
		Set NodX = frm2.uniTree1.Nodes(strNewKey)
        If Err.Number <> 0 Then
            Exit Do
        Else
        	Err.Clear
            tmpCnt = tmpCnt + 1
        End If
    Loop
    
    Set NodX             = frm2.uniTree1.Nodes.Add(Node.Key, tvwChild, strNewKey, C_NEW_FOLDER, C_USFolder, C_USFolder)
    NodX.ExpandedImage   = C_USOpen
    NodX.tag             = 0
	NodX.Parent.Tag      = Cint(NodX.Parent.Tag) + 1
	NodX.Parent.Expanded = True
	NodX.Selected        = True

	StrVal = StrVal &  "C"                    & gColSep
	StrVal = StrVal &  strNewKey              & gColSep
	StrVal = StrVal &  Node.Key               & gColSep
	StrVal = StrVal &  C_NEW_FOLDER           & gColSep
    StrVal = StrVal &  "M"                    & gColSep
    StrVal = StrVal &  (GetNodeLvl(Node) + 1) & gColSep
    StrVal = StrVal &  (Cint(Node.Tag) + 1)   & gColSep
    StrVal = StrVal &  ""                     & gColSep
    StrVal = StrVal &  ""                     & gRowSep
    
	Set NodX = Nothing
	Set Node = Nothing
	
	frm2.uniTree1.SetFocus
	Call frm2.uniTree1.StartLabelEdit

End Function

'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
'#########################################################################################################

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQueryOk�� �������϶� ���� 
'========================================================================================
Function DbQueryOk()
	lgBlnBizLoadMenu = True
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	frm2.uniTree1.MousePointer = 0
    Self.close
End Function
</SCRIPT>
</HEAD>

<BODY SCROLL=no rightmargin="0" bgcolor="#ffffff">
<FORM NAME="frm2" TARGET="PopBizASP" METHOD="POST">
<TABLE WIDTH="85%" HEIGHT="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
	<TR>
		<TD WIDTH="100%" VALIGN="BOTTOM">
			<TABLE WIDTH="100%" id="TABLE1">
				<TR HEIGHT=20>
					<TD WIDTH="25%"></TD>
					<TD WIDTH="75%">����� �޴��� �� ȭ���� �߰��մϴ�.</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD VALIGN="TOP">
			<FIELDSET>
			<TABLE CELLSPACING="0" CELLPADDING="2" WIDTH="100%" id="TABLE1">
				<TR HEIGHT="25">
					<TD CLASS="TD5" WIDTH="20%" ALIGN="RIGHT"><LABEL>�̸�</LABEL>&nbsp;&nbsp;</TD>
					<TD CLASS="TD6" WIDTH="90%"><INPUT TYPE="TEXT" id=txtName style="font-size:9pt;HEIGHT: 100%; WIDTH: 100%" TAG =11></TD>
				</TR>
			</TABLE>
			</FIELDSET>
			<FIELDSET>
			<TABLE CELLSPACING="0" CELLPADDING="2" WIDTH="100%" id="TABLE1">
				<TR HEIGHT="20">
					<TD CLASS="TD5" WIDTH="20%" ALIGN="RIGHT" VALIGN="bottom"><LABEL>��ġ</LABEL>&nbsp;&nbsp;</TD>
					<TD CLASS="TD6" WIDTH="90%" ROWSPAN="18">
						<script language =javascript src='./js/folderadd_uniTree1_N237411767.js'></script>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" WIDTH="20%">&nbsp;</TD>
				</TR>
			</TABLE>
			</FIELDSET>
			<TABLE>
				<TR Class="Toolbar">
					<TD>
					<IFRAME ID="PopBizASP" NAME="PopBizASP" SRC='../blank.htm' width="100%" height="0" STYLE="display: ''"></IFRAME>
					</TD>
				</TR>
			</TABLE>
			<TABLE WIDTH="100%">
			 <TR>
				<TD ALIGN="CENTER">
					<IMG SRC="../image/bu_confirm_off.gif" BORDER="0" Style="CURSOR: hand" ALT="Ȯ  ��" NAME="btnOK"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_confirm_on.gif',1)">&nbsp;&nbsp;
					<IMG SRC="../image/bu_cancel_off.gif"  BORDER="0" Style="CURSOR: hand" ALT="��  ��" NAME="btnCancel"    onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_cancel_on.gif',1)">&nbsp;&nbsp;
					<IMG SRC="../image/bu_new_off.gif"     BORDER="0" Style="CURSOR: hand" ALT="������" NAME="btnNewFolder" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/bu_new_on.gif',1)">
				</TD>
			 </TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<TEXTAREA NAME=txtMulti STYLE="display: none"></TEXTAREA>
</FORM>
</BODY>
</HTML>
