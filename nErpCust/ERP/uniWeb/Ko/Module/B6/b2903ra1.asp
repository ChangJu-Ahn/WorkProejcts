
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : �λ�/�޿� 
*  2. Function Name        : ��������ȸ 
*  3. Program ID           : B2903ma1
*  4. Program Name         : ������ ��ȸ 
*  5. Program Desc         : �������� Ʈ���� ���·� �����ش� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001//
*  8. Modified date(Last)  : 2002/12/17
*  9. Modifier (First)     : �̼��� 
* 10. Modifier (Last)      : Sim Hae Young
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE>���������泻��</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE = "VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<Script Language = "VBScript"   SRC="../../inc/incUni2KTV.vbs"></Script>
<SCRIPT LANGUAGE = "JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language = "VBScript">

Option Explicit                                                        '��: indicates that All variables must be declared in advance


'==========================================================================================================

Const BIZ_PGM_ID = "b2903rb1.asp"


Dim IsOpenPop
'----treeView constants
Dim lgIsClicked
Dim RootNode

Const IsNode           = 1
Const IsNodeKey        = 0

Const C_PAR_DEPT_CD    = 0
Const C_DEPT_CD        = 1
Const C_INTERNAL_CD    = 2
Const C_DEPT_FULL_NM   = 3

'----Spread constants
Const C_MaxCols        = 4
Const C_NAME           = 1
Const C_EMP_NO         = 2
Const C_ROLE_CD        = 3
Const C_PAY_GRADE1     = 4
Const C_TEL_NO         = 5


Dim ArrParent, PopupParent

ArrParent = window.dialogArguments

Set PopupParent = ArrParent(0)


'==========================================  3.3.1 MakeTree()  ======================================
'	Name : makeTree()
'	Desc : �������� �����Ѵ� 
'          ������� key�� "ID" & INTERNAL_CD�� �Ѵ� 
'========================================================================================================= 


Sub makeTree(DeptList,CoName)
	dim val,i
	dim arrDept, arrLine
	dim NodX 
	dim KeyVal, CD, dName, ParentKey

	On Error Resume Next
	val = replace(DeptList,vbCrLf,chr(12))
		
    arrDept = Split(val, chr(12),-1,1)
	
	arrLine = Split(" " & arrDept(0), chr(11))
	
	Set RootNode = frm1.uniTree1.Nodes.add(,tvwChild,"ID1" ,CoName)  'ȸ����� ��Ʈ�� ��ġ��Ų�� 
       For i = 0 To UBound(arrDept, 1) 
           
           If arrDept(i) = "" Then 
    	      Exit For
           End If   
           arrLine = Split(arrDept(i), chr(11))
           CD      = arrLine(C_INTERNAL_CD)
           KeyVal  = "ID" & CD
           dName   = arrLine(C_DEPT_FULL_NM)
           
           ParentKey = GetParentKey(KeyVal)

        '   Set NodX = frm1.uniTree1.Nodes.Add(Left(KeyVal,Len(KeyVal)-1), tvwChild,KeyVal,dName) 'interanl_cd�� Ʈ���� �����Ѵ� 
			
			Set NodX = frm1.uniTree1.Nodes.Add(ParentKey, tvwChild,KeyVal,dName) 'interanl_cd�� Ʈ���� �����Ѵ� 
			
			if err.number <> 0 then
		      err.Clear 
		      msgbox "�μ������� �߸��Ǿ� �ֽ��ϴ� : " & dName
		   end if
	
       Next
       RootNode.Expanded = true
       
End Sub



'======================================================================================================
' Function Name : GetParentKey
' Function Desc : When making treeview, Searching Parent key of a node
'======================================================================================================
function GetParentKey(Nodx )
	
	Dim LenNodx, i
	Dim PnodKey
	
	LenNodx = Len(Nodx)
	
	i = 0
	PnodKey = Nodx

	do while i <= LenNodx  
		if mid(nodx, LenNodX - i - 1 , 1) <> "0" then
			PnodKey = left(nodx,LenNodX-i-1)
			exit do
		end if
		
		i = i + 1
	Loop
		
	GetParentKey = PnodKey
	
	
End function

'======================================================================================================
' Function Name : allExpand(nodx)
' Function Desc : ��üȮ���ư�� ������ ��� �μ��� ���δ� 
'======================================================================================================


Sub allExpand(nodx)
	
	with frm1.uniTree1
		.nodes(nodx.key).expanded=true
		if .nodes(nodx.key).children > 0 then
			allExpand(.nodes(nodx.key).child)
		
		end if
		
		if .nodes(nodx.key) <> .nodes(nodx.key).LastSibling then
			allExpand(.nodes(nodx.key).next)
		
		end if
		
	end with
		
End Sub


'======================================================================================================
' Function Name : allCollapse(nodx)
' Function Desc : ��ü��ҹ�ư�� ������ ���θ�� �ٷι� �μ��� ���δ� 
'======================================================================================================
Sub allCollapse(nodx)

	with frm1.uniTree1
		.nodes(nodx.key).expanded=false
		if .nodes(nodx.key).children > 0 then
			allCollapse(.nodes(nodx.key).child)
		else
			Exit Sub
		end if
		
		if .nodes(nodx.key) <> .nodes(nodx.key).LastSibling then
			allCollapse(.nodes(nodx.key).next)
		else 
			Exit sub
		end if
		
	end with
	
End Sub


'======================================================================================================
' Function Name : allExpand_ButtonClicked()
' Function Desc : Ʈ���� Ȯ�带 ���� ��ư �̺�Ʈ 
'======================================================================================================
sub allExpand_ButtonClicked()
    If CheckDeptID() = True Then
        Exit Sub
    End If

	call allExpand(RootNode)		
		
End sub

'======================================================================================================
' Function Name : allCollapse_ButtonClicked()
' Function Desc : Ʈ���� ��Ҹ� ���� ��ư �̺�Ʈ 
'=====================================================================================================
sub allCollapse_ButtonClicked()
    If CheckDeptID() = True Then
        Exit Sub
    End If

	call allCollapse(RootNode)
	RootNode.expanded = true
End Sub


sub InitVariables()
	frm1.unitree1.nodes.clear
End Sub

'------------------------------------------  OpenOrgId()  -------------------------------------------
'	Name : OpenOrgID()
'	Description : OrgId PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOrgId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�μ�����ID �˾�"			' �˾� ��Ī 
	arrParam(1) = "horg_abs"					' TABLE ��Ī 
	arrParam(2) = frm1.txtOrgId.value		    ' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "�μ�����ID"				' �����ʵ��� �� ��Ī 
	
    arrField(0) = "orgid"					    ' Field��(0)%>
    arrField(1) = "orgnm"					    ' Field��(1)%>
    
    arrHeader(0) = "�μ�����ID"				' Header��(0)%>
    arrHeader(1) = "�μ������"				' Header��(1)%>
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetOrgId(arrRet)
	End If	
	
End Function


'------------------------------------------  SetOrgId()  --------------------------------------------
'	Name : SetOrgId()
'	Description : OrgId Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetOrgId(Byval arrRet)
	With frm1
		.txtOrgId.value = arrRet(0)
		.txtOrgNm.value = arrRet(1)
	End With
End Function

Function CheckDeptID()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim blnResult

    If  CommonQueryRs(" ORGNM "," HORG_ABS "," ORGID= " & FilterVar(frm1.txtOrgId.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False then
        blnResult = True
    Else
        blnResult = False
    End If

    CheckDeptID = blnResult
End Function


'==========================================  3.1.2 Form_Load()  ======================================
'	Name : Form_Load()
'	Desc : 
'========================================================================================================= 
Sub Form_Load()

	Call ggoOper.LockField(Document, "N")   
	lgIsClicked = False
	IsOpenPop   = False
	frm1.btnCb_allCollapse.disabled = true
	frm1.btnCb_allExpand.disabled = true 

end sub

Function FncQuery() 
    Dim IntRetCD 
    Dim Tree
    
    FncQuery = False                                                        

    Err.Clear                                                               'Protect system from crashing
	
	Call InitVariables 
    															
    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
        frm1.txtOrgNm.value = ""
	    frm1.btnCb_allCollapse.disabled = true
	    frm1.btnCb_allExpand.disabled = true 
        frm1.txtOrgId.focus
       Exit Function
    End If
    
    If CheckDeptID() = True Then
    	Call DisplayMsgBox("970000", "X", "�μ�����ID", "X")     
        frm1.txtOrgNm.value = ""
	    frm1.btnCb_allCollapse.disabled = true
	    frm1.btnCb_allExpand.disabled = true 
        frm1.txtOrgId.focus
        Exit Function    
    End If
    
  '-----------------------
  'Query function call area
  '----------------------- 

	
	Call DBQuery()
	
       
    FncQuery = True															
    
End Function

Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing

	Call LayerShowHide(1)

	
    With frm1
    
		strVal = BIZ_PGM_ID & "?txtOrgId=" & frm1.txtOrgId.value

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
		
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()														'��: ��ȸ ������ ������� 
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed

End Function



Function OKClick()
	Self.Close()
End Function


Function Document_onKeyUp()
	Dim objEl, KeyCode

	Set objEl = window.event.srcElement
	
	KeyCode = window.event.keycode

	If KeyCode = 27 Then
       Call OKClick()
    End If
End Function

</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5">�μ�����ID</TD>
						<TD CLASS="TD656">
							<INPUT TYPE=TEXT NAME="txtOrgId" SIZE=10 MAXLENGTH=5 tag="12XXXU"  ALT="�μ�����ID" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrgId" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOrgId()" >
							<INPUT TYPE=TEXT NAME="txtOrgNm" Size=40 tag="14">
						</TD>
					</TR>
			</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/b2903ra1_OBJECT1_uniTree1.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20 Valign=TOP>
		<TD  colspan=2>
			<TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
		            <TD WIDTH=10>&nbsp;</TD>
				    <TD Align="center"><BUTTON NAME="btnCb_allExpand" CLASS="CLSMBTN" ONCLICK="VBScript: allExpand_ButtonClicked()">��üȮ��</BUTTON>&nbsp;&nbsp;
						<BUTTON NAME="btnCb_allCollapse" CLASS="CLSMBTN" ONCLICK="VBScript: allCollapse_ButtonClicked()">��ü���</BUTTON></TD>
					<TD WIDTH=* ALIGN="right"></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
				
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ONCLICK="FncQuery()" ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)" ONCLICK="OkClick()"></IMG>
							                  </TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


  

