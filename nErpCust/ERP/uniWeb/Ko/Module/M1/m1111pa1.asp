<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1111PA1
'*  4. Program Name         : 품목정보팝업 
'*  5. Program Desc         : 품목정보팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : KimTaeHyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit                            

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "m1111pb1.asp"                                      '☆: 비지니스 로직 ASP명 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const POPUP_TITLE = 0                                                           '--- Index of POP-UP Title
Const TABLE_NAME  = 1                                                           '--- Index of DB table name to query
Const CODE_CON    = 2                                                           '--- Index of Code Condition value
Const NAME_CON    = 3                                                           '--- Index of Name Condition value
Const WHERE_CON   = 4                                                           '--- Index of Where Clause
Const TEXT_NAME   = 5                                                           '--- Index of Textbox Name
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgStrCodeKey
Dim lgStrNameKey
Dim lgSortKey

Dim arrParent
Dim arrParam				 '--- First Parameter Group
Dim arrTblField				 '--- Second Parameter Group(DB Table Field Name)

Dim arrReturn				 '--- Return Parameter Group
Dim gintDataCnt				 '--- Data Counts to Query
Dim arrFieldType
Dim lgIsOpenPop 

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec			
		'------ Set Parameters from Parent ASP ------ 
		arrParent = window.dialogArguments
		Set PopupParent = arrParent(0)
		arrParam = arrParent(1)
		arrTblField = arrParent(2)
		top.document.title = arrParam(POPUP_TITLE)			'Common Popup과 같이 사용함(2003.02.10)

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCD	=  1
	C_ItemNm	=  2
	C_Spec		=  3
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With vspdData
		ggoSpread.Source = vspdData
        ggoSpread.Spreadinit "V20021127",, PopupParent.gAllowDragDropSpread
		
		.ReDraw = False

		.MaxCols = C_Spec + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_ItemCD, "품목코드", 25
		ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 25
		ggoSpread.SSSetEdit 	C_Spec, "규격", 25
		
		Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
		Call SetSpreadLock()
		.ReDraw = True
	End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
    End Select    
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	txtCd.value = arrParam(CODE_CON)
	txtNm.value = arrParam(NAME_CON)
			
	Self.Returnvalue = Array("")
End Sub

'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
  ggoSpread.Source = vspdData
  ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call Search_OnClick()
End Sub

'++++++++++++++++++++++++++++++++++++++++++++  OpenJnlItem()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenJnlItem()																				+
'+	Description : Sales Order Type PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenJnlItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "품목계정"								<%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"										<%' TABLE 명칭 %>
	arrParam(2) = Trim(txtJnlItem.value)						<%' Code Condition%>
	arrParam(3) = ""											<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("P1001", "''", "S") & ""							<%' Where Condition%>
	arrParam(5) = "품목계정"								<%' TextBox 명칭 %>

	arrField(0) = "MINOR_CD"									<%' Field명(0)%>
	arrField(1) = "MINOR_NM"									<%' Field명(1)%>

	arrHeader(0) = "품목계정"								<%' Header명(0)%>
	arrHeader(1) = "품목계정명"								<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		txtJnlItem.focus
		Exit Function
	Else
		txtJnlItem.value = arrRet(0)
		txtJnlItemNm.value = arrRet(1)
		call SetCookieJnlItem()
		txtJnlItem.focus
	End If
End Function
	
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetCookieJnlItem()																			+
'+	Description : JnlItem Code is saved at Cookie					 									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetCookieJnlItem
	Dim 	ExpDate
	ExpDate = "Sun 31-Jan-2999 12:00:00 GMT" '한번 세팅된 품목계정 코드는 계속 저장되어 있게 변경함.
	Document.Cookie = "m1311JnlItem" & "=" & FilterVar(Trim(txtJnlItem.value),"","SNM") & "; path=" & "/; expires=" & ExpDate
		
End Function

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../comasp/ComLoadInfTB19029.asp" -->
End Sub	

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()
	call Search_OnClick()	
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function OKClick()
	Dim intColCnt
		
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
				
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
		
	Self.Close()
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub Search_OnClick()
    vspdData.MaxRows = 0
    lgStrCodeKey = ""
    lgStrNameKey = ""

	Call DbQuery()
End Sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call Search_OnClick()
	End If
End sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or vspdData.MaxRows = 0 Then 
         Exit Sub
    End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Sub

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
       Exit Sub
    End If
	    
    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
       If lgStrCodeKey <> "" Or lgStrNameKey <> "" Then
 		  DbQuery
       End If
    End if
End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"   

    Call SetPopupMenuItemInf("0000111111")
    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub	

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Description   : 
'========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strPreWhere, strWhere
    Dim iLoop
    Dim arrStrVal
    Dim arrStrDT
	
	DbQuery = False    
    
    strPreWhere = ""
    strWhere = ""
	arrStrVal = ""
	arrStrDT = ""
	
	If UCase(Trim(arrParam(WHERE_CON))) <> "" Then
		strPreWhere = UCase(Trim(arrParam(WHERE_CON))) 
    End If
    
    If Trim(txtJnlItem.value) <> ""  then 		    
		strWhere = vbCr & "WHERE " & strPreWhere & " And B_ITEM_BY_PLANT.ITEM_ACCT =  " & FilterVar(txtJnlItem.value, "''", "S") '& ", " '& " AND "			
	else
		strWhere = vbCr & "WHERE " & strPreWhere 		
	End if
	
   '----- Code가 있을 경우는 Name에 상관없이 Code로만 조회하고, Code가 없는 경는 Name으로 조회한다.
	If Trim(txtCd.value) = "" AND Trim(txtNm.value) = "" Then 'All
		strWhere = strWhere & " AND " & Trim(arrTblField(0)) & ">= " & FilterVar(UCase(lgStrCodeKey), "''", "S") & "  Order by " & Trim(arrTblField(0)) & vbCr
	ElseIf  Trim(txtCd.value) <> "" Then 'Code
		strWhere = strWhere & " AND " & Trim(arrTblField(0)) & ">= " & FilterVar(UCase(lgStrCodeKey), "''", "S") & " " & vbCr
		strWhere = strWhere & " AND " & Trim(arrTblField(0)) & ">= " & FilterVar(UCase(txtCd.value), "''", "S") & "  Order by " & Trim(arrTblField(0))  & vbCr
	Else 'Name
		strWhere = strWhere & " AND " & Trim(arrTblField(0)) & ">= " & FilterVar(UCase(lgStrCodeKey), "''", "S") & " " & vbCr
		strWhere = strWhere & " AND " & Trim(arrTblField(1)) & ">= " & FilterVar(UCase(lgStrNameKey), "''", "S") & " " & vbCr
		strWhere = strWhere & " AND " & Trim(arrTblField(1)) & ">= " & FilterVar(UCase(txtNm.value), "''", "S") & "  Order by " & Trim(arrTblField(1)) '& "," & Trim(arrTblField(1))  & vbCr
	End IF

	For iLoop = 0 To 1
	    arrStrVal = arrStrVal & Trim(arrTblField(iLoop)) & PopupParent.gColSep
	Next
	arrStrVal = arrStrVal & "B_ITEM.SPEC" & PopupParent.gColSep 

    arrStrDT = "ED" & PopupParent.gColSep & "ED" & PopupParent.gColSep & "ED" & PopupParent.gColSep
	    
	strVal = BIZ_PGM_ID & "?txtTable=" & Trim(arrParam(TABLE_NAME)) 
	strVal = strVal & "&txtWhere="    & strWhere
	strVal = strVal & "&txtJnlItem=" & Trim(txtJnlItem.value)
	strVal = strVal & "&gintDataCnt=" & 2
	strVal = strVal & "&arrField="    & arrStrVal
	strVal = strVal & "&arrStrDT="    & arrStrDT

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
		
	DbQuery = True                                                          '⊙: Processing is NG
End Function
	
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function DbQueryOk()
   Dim IntRetCD

   If vspdData.MaxRows = 0 Then
      IntRetCD = DisplayMsgBox("900014","X","X","X") 
      If Trim(txtCd.value) > "" Then
         txtCd.Select 
         txtCd.Focus
      Else   
         txtNm.Select 
         txtNm.Focus
     End If   
   End If  
End Function

	
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>						
			<TR>	
				<TD CLASS=TD5 NOWRAP>품목</TD>
				<TD CLASS=TD6 NOWRAP>
					<INPUT NAME="txtCd" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="품목"></TD>
				<TD CLASS=TD5 NOWRAP>품목계정</TD>
				<TD CLASS=TD6 NOWRAP>
					<INPUT NAME="txtJnlItem" TYPE="Text" MAXLENGTH="20" SIZE=10 tag="11XXXU" ALT="품목계정" onChange = "vbscript:SetCookieJnlItem"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnJnlItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenJnlItem">&nbsp;
					<INPUT NAME="txtJnlItemNm" TYPE="Text" SIZE=20 tag="14X">
				</TD>
			</TR>	
			<TR>
				<TD CLASS=TD5 NOWRAP>품목명</TD>
				<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNm" TYPE="Text" SIZE=30 MAXLENGTH="50" ALT="품목명" tag="11"></TD>
				<TD CLASS=TD5 NOWRAP></TD>
				<TD CLASS=TD6 NOWRAP></TD>
			</TR>
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<script language =javascript src='./js/m1111pa1_vaSpread1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>		
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
