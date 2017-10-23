

<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 인사/급여 
*  2. Function Name        : 조직도조회 
*  3. Program ID           : B2903ma1
*  4. Program Name         : 조직도 조회 
*  5. Program Desc         : 조직도를 트리뷰 형태로 보여준다 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001//
*  8. Modified date(Last)  : 2002/12/17
*  9. Modifier (First)     : 이석민 
* 10. Modifier (Last)      : Sim Hae Young
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		


<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBScript" SRC="../../inc/incUni2KTV.vbs"></Script>
<Script Language="JavaScript" SRC="../../inc/incImage.js"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID = "b2903mb1.asp"
Const Jump_PGM_ID = "H2001ma1"

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

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_NAME      
Dim C_EMP_NO    
Dim C_ROLE_CD   
Dim C_PAY_GRADE1
Dim C_TEL_NO    

Sub InitSpreadPosVariables()
    C_NAME       = 1
    C_EMP_NO     = 2
    C_ROLE_CD    = 3
    C_PAY_GRADE1 = 4
    C_TEL_NO     = 5
End Sub


'==========================================  3.3.1 MakeTree()  ======================================
'	Name : makeTree()
'	Desc : 조직도를 구성한다 
'          각노드의 key는 "ID" & INTERNAL_CD로 한다 
'========================================================================================================= 
Sub makeTree()
	Dim val,i
	Dim arrDept, arrLine
	Dim NodX 
	Dim KeyVal, CD, dName ,ParentKey
	Dim erDept1, erDept2

    On Error Resume Next
	Err.Clear 
	
	erDept1 = "부서코드 충돌되는 부서-------" & chr(13)
	erDept2 = "내부부서코드가 잘못된부서----" & chr(13)
	
	val =replace(frm1.DeptList.value ,vbCrLf,chr(12))
		
    arrDept = Split(val, chr(12),-1,1)
	
	arrLine = Split(" " & arrDept(0), chr(11))
	
	Set RootNode = frm1.uniTree1.Nodes.add(,tvwChild,"ID1" ,frm1.CoName.value)  '회사명을 루트에 위치시킨다 
		For i = 0 To UBound(arrDept, 1) 
           
			If arrDept(i) = "" Then 
    			Exit For
			End If   

			arrLine = Split(arrDept(i), chr(11))
			CD      = arrLine(C_INTERNAL_CD)
			KeyVal  = "ID" & CD

			dName   = arrLine(C_DEPT_FULL_NM)

			ParentKey = GetParentKey(KeyVal)

			Set NodX = frm1.uniTree1.Nodes.Add(ParentKey, tvwChild,KeyVal,dName) 'interanl_cd로 트리를 구성한다 

			If Err.Clear  <> 0 Then
			   Err.Clear 
			   Msgbox "부서정보가 잘못되어 있습니다 : " & dName
			End If
		Next

       RootNode.Expanded = true

       Call uniTree1_nodeClick(RootNode)

End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData	

        ggoSpread.Spreadinit "V20021216",,parent.gAllowDragDropSpread    
             
        .ReDraw = false

        .MaxCols = C_TEL_NO + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
        .Col = .MaxCols														'☆: 사용자 별 Hidden Column
        .ColHidden = True    
        	           
        .MaxRows = 0
        ggoSpread.ClearSpreadData
        	
        Call GetSpreadColumnPos("A")  

        ggoSpread.SSSetEdit   C_NAME,             "성명",      20,,,50,2
        ggoSpread.SSSetEdit   C_EMP_NO,           "사번",      12,,,15,2
        ggoSpread.SSSetEdit   C_ROLE_CD,          "직책",      17,,,20,2
        ggoSpread.SSSetEdit   C_PAY_GRADE1,       "급호",      17,,,20,2
        ggoSpread.SSSetEdit   C_TEL_NO,           "전화",      14,,,20,2
        
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
	
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_NAME       = iCurColumnPos(1)
            C_EMP_NO     = iCurColumnPos(2)
            C_ROLE_CD    = iCurColumnPos(3)
            C_PAY_GRADE1 = iCurColumnPos(4)
            C_TEL_NO     = iCurColumnPos(5)
    End Select    
End Sub

Sub uniTree1_NodeClick(Node) 
	Dim strPar
	strPar = BIZ_PGM_ID & "?Nodekey="
	strPar = strPar & Node.key
	strPar = strPar & "&fnc=EMP"
	call RunMyBizAsp(MyBizAsp, strPar)

End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

	frm1.vspdData.col = C_EMP_NO
	frm1.vspdData.row = row
	
	if row <> 0 then
		call open_Emp_Detail(frm1.vspdData.text)	
	end if
	
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

function GetParentKey(Nodx )
	
	Dim LenNodx, i
	Dim PnodKey
	
	LenNodx = Len(Nodx)
	i = 0
	PnodKey = Nodx
	do while i <= LenNodx  
		if mid(nodx, LenNodX-i-1 , 1) <> "0" then
			PnodKey = left(nodx,LenNodX-i-1)
			exit do
		end if
		
		i = i + 1
	Loop
		
	GetParentKey = PnodKey
	
	
End function

Sub open_Emp_Detail(EMP)
    Dim IntRetCD 
	Dim arrRet
	Dim iCalledAspName
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("Emp_Detail_popup")
	
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "Emp_Detail_popup", "X")
		lgIsOpenPop = False
		Exit Sub
	End If
	
	arrRet = window.showModalDialog(iCalledAspName&"?EMP_NO="&Cstr(EMP),Array(window.parent,"","",""), "dialogWidth=400px; dialogHeight=500px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
	if arrRet = true then
		CookiePage 1, EMP
		pgmJump(Jump_PGM_ID)
	end if	
End Sub

Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    With frm1
    .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_NAME, -1, 1
        ggoSpread.SpreadLock    C_EMP_NO, -1, 1
        ggoSpread.SpreadLock    C_ROLE_CD, -1, 1
        ggoSpread.SpreadLock    C_PAY_GRADE1, -1, 1
        ggoSpread.SpreadLock    C_TEL_NO, -1, 1
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With

End Sub

Sub allExpand(nodx)
	
	with frm1.uniTree1
		.nodes(nodx.key).expanded=true
		if .nodes(nodx.key).children > 0 then
			allExpand(.nodes(nodx.key).child)
		end if
		
		if .nodes(nodx.key) <> .nodes(nodx.key).LastSibling then
			allExpand(.nodes(nodx.key).next)
		else 
			Exit sub
		end if
		
	end with
		
End Sub

Sub allCollapse(nodx)

	with frm1.uniTree1
		.nodes(nodx.key).expanded=false
		if .nodes(nodx.key).children > 0 then
			allCollapse(.nodes(nodx.key).child)
		end if
		
		if .nodes(nodx.key) <> .nodes(nodx.key).LastSibling then
			allCollapse(.nodes(nodx.key).next)
		else 
			Exit sub
		end if
		
	end with
	
End Sub

sub allExpand_ButtonClicked()

	call allExpand(RootNode)		
		
End sub

sub allCollapse_ButtonClicked()
	call allCollapse(RootNode)
	RootNode.expanded = true
End Sub

Sub Form_Load()
	lgIsClicked = False
	IsOpenPop   = False

	frm1.unitree1.HideSelection = false
	
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>

	
	call FncQuery(1)
	
end sub
	

Sub Form_QueryUnLoad( Cancel , UnloadMode)
	
End Sub

Function FncQuery(par)
    Dim strPar
	strPar = BIZ_PGM_ID & "?fnc=TREE"
	call RunMyBizAsp(MyBizAsp, strPar)
End Function


Function FncExit()
    FncExit = True
End Function


Function OpenDeptHistory()
	Dim arrRet
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B2903ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B2903ra1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,"","",""), "dialogWidth=500px; dialogHeight=700px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

Function CookiePage(ByVal flgs,EMP)
	On Error Resume Next

	Const CookieSplit = 4877	
	Const DeptcookieSplit = 5877
	
	Dim strTemp
	Dim strPar, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim tempNode
	If flgs = 1 Then
		WriteCookie CookieSplit , cstr(EMP)
		
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(DeptcookieSplit)
		
		If strTemp = "" then Exit Function
			
		
		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			WriteCookie DetpCookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		WriteCookie DeptCookieSplit , ""	
		
		strPar = " ORG_CHANGE_ID = CUR_ORG_CHANGE_ID	AND DEPT_CD =  " & FilterVar(strTemp, "''", "S") & " " 
		
		Call CommonQueryRs(" internal_cd "," B_ACCT_DEPT, B_COMPANY ", strPar , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
		Set TempNode = frm1.unitree1.nodes("ID" & Trim(replace(lgF0,chr(11),"")))
		
		TempNode.EnsureVisible
		call uniTree1_NodeClick(TempNode)
		
		TempNode.selected = true
	
				
	End If
End Function

sub DbQueryOk()
	call MakeTree()
'	Call CookiePage (0,0)                                                             '☜: Check Cookie
end sub


</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL=no rightmargin=0 bgColor="#FFFFFF">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="GET">
<TABLE HEIGHT=100% WIDTH=100%  <%=LR_SPACE_TYPE_00%> >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->
	<TR HEIGHT=23>
		<TD WIDTH="100%" >
			<TABLE <%=LR_SPACE_TYPE_10%> class="BasicTB">
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:OpenDeptHistory()">조직도변경내역</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR >
		<TD height=* WIDTH=100% VALIGN=TOP class="TAB11">
			<TABLE <%=LR_SPACE_TYPE_20%> class="TB3">
				
				<TR>		
					<TD width = 30%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT <%=UNI2KTV_IDVER%> name=uniTree1 width=100% height=100% TAG="2"> <PARAM NAME="LineStyle" VALUE="0"> <PARAM NAME="Style" VALUE="6"> <PARAM NAME="LabelEdit" VALUE="1"> <PARAM NAME="indentation" VALUE="350"> </OBJECT>');</SCRIPT>
					</TD>
					<TD width = 70%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=EMPLIST> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</Table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<!-- space Area-->
	<TR HEIGHT=20 Valign=TOP>
		<TD  colspan=2>
			<TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
		            <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnCb_allExpand" CLASS="CLSMBTN" ONCLICK="VBScript: allExpand_ButtonClicked()">전체확장</BUTTON>&nbsp;&nbsp;
						<BUTTON NAME="btnCb_allCollapse" CLASS="CLSMBTN" ONCLICK="VBScript: allCollapse_ButtonClicked()">전체축소</BUTTON></TD>
					<TD WIDTH=* ALIGN="right"></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
				
	</TR>
			
	
	<TR >
		<TD HEIGHT=<%=BizSize%> WIDTH=100% COLSPAN=2>
			<IFRAME ID="MyBizASP" NAME="MyBizASP" SRC='../../blank.htm' width=100% height=1 STYLE="display: ''"></IFRAME> 
		</TD>
	</TR>
	
</TABLE>
<input type=hidden name=DeptList value="">
<input type=hidden name=CoName  value="" >
</FORM>
</BODY>
</HTML>


