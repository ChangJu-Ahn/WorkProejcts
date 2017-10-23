<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3112pa2.asp
'*  4. Program Name         : 품목팝업(수주내역등록)
'*  5. Program Desc         : 품목팝업(수주내역등록)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop                                              
Dim lgBlnItemGroupCdChg
Dim lgMark                                                  


Dim arrReturn					
Dim arrParam

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3112pb2.asp"
Const C_MaxKey          = 7                                                                         
Const C_PopItemGroupCd	=	1
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'=============================================================================================================
Sub InitVariables()
	lgPageNo = ""
    lgBlnFlgChgValue = False                                        
    lgSortKey        = 1	
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Sub

'=============================================================================================================
Sub SetDefaultVal()
	arrParam = arrParent(1)
	frm1.txtItem.value = arrParam(0)
	frm1.txtPlant.value = arrParam(1) 
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=============================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3112pa1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock      
End Sub


'=============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	.vspdData.OperationMode = 3
    .vspdData.ReDraw = True
    End With
End Sub

'=============================================================================================================
Function OpenJnlItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "품목계정"								
	arrParam(1) = "b_minor"									
	arrParam(2) = Trim(frm1.txtJnlItem.value)					
	arrParam(3) = ""											
	arrParam(4) = "major_cd = " & FilterVar("P1001", "''", "S") & ""								
	arrParam(5) = "품목계정"								

	arrField(0) = "MINOR_CD"										
	arrField(1) = "MINOR_NM"										

	arrHeader(0) = "품목계정"								
	arrHeader(1) = "품목계정명"								

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetJnlItem(arrRet)
	End If
End Function

'==================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere

	Case C_PopItemGroupCd
		iArrParam(1) = "dbo.B_ITEM_GROUP "					<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtItemGrp.value)	<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = "LEAF_FLG = " & FilterVar("Y", "''", "S") & "  AND DEL_FLG = " & FilterVar("N", "''", "S") & " "	<%' Where Condition%>
		iArrParam(5) = frm1.txtItemGrp.alt			'"품목그룹"		<%' TextBox 명칭 %>
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "ITEM_GROUP_CD"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "ITEM_GROUP_NM"	<%' Field명(1)%>
		    
		iArrHeader(0) = "품목그룹"						<%' Header명(0)%>
		iArrHeader(1) = "품목그룹명"					<%' Header명(1)%>

		frm1.txtItemGrp.focus

	End Select
 
	iArrParam(0) = iArrParam(5)								<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
	
End Function

'=====================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopItemGroupCd
		frm1.txtItemGrp.value = pvArrRet(0) 
		frm1.txtItemGrpNm.value = pvArrRet(1)
		lgBlnItemGroupCdChg = False
	End Select

	SetConPopup = True

End Function


'=============================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function


	lgIsOpenPop = True

	arrParam(0) = "공장"				
	arrParam(1) = "B_PLANT"							
	arrParam(2) = Trim(frm1.txtPlant.value)		
	arrParam(4) = ""							
	arrParam(5) = "공장"				
		
	arrField(0) = "PLANT_CD"				
	arrField(1) = "PLANT_NM"				
	    
	arrHeader(0) = "공장"					
	arrHeader(1) = "공장명"				

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
End Function

'=============================================================================================================
Function OpenSortPopup()
	Dim arrRet

    On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
	
'=============================================================================================================
Function SetJnlItem(arrRet)
	frm1.txtJnlItem.value = arrRet(0)
	frm1.txtJnlItemNm.value = arrRet(1)
	frm1.txtJnlItem.focus
End Function

'=============================================================================================================
Function SetPlant(Byval arrRet)
	With frm1
		.txtPlant.value = arrRet(0) 
		.txtPlantNm.value = arrRet(1)   
		.txtPlant.focus
	End With
End Function

'=============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														
	
    Call ggoOper.LockField(Document, "N")                                   
  
  	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	
	If frm1.txtItem.value <> "" Or frm1.txtPlant.value <> "" Then
		Call FncQuery()
	End If
	
End Sub


'=============================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'=============================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgPageNo <> "" Then							
				DbQuery
			End If
		End If
	End With
End Sub
	

'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
	If CheckRunningBizProcess = True Then Exit Sub	
			
    	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub


'=============================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then   'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'=============================================================================================================
Function OKClick()
		
	Redim arrReturn(2)
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		arrReturn(0) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",2)
		arrReturn(1) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",6)
		arrReturn(2) = frm1.vspdData.Text
			
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'=============================================================================================================
Function CancelClick()
	Self.Close()
End Function


'=============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                                

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 
    Call DbQuery															

    FncQuery = True		
End Function

'=============================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)		
		strVal = strVal & "&txtItemNm=" & Trim(frm1.txtItemNm.value)	
		strVal = strVal & "&txtJnlItem=" & Trim(frm1.txtJnlItem.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
		strVal = strVal & "&txtItemGrp=" & Trim(frm1.txtItemGrp.value)
		strVal = strVal & "&txtItemSpec=" & Trim(frm1.txtItemSpec.value)
		
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        strVal = strVal & "&lgPageNo="   & lgPageNo                    
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
  
        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True
End Function

'=============================================================================================================
Function DbQueryOk()														

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtItem.focus
	End If

End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>	
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" ALT="품목"></TD>
						<TD CLASS="TD5" NOWRAP>품목그룹</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtItemGrp" ALT="품목그룹" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPopup(C_PopItemGroupCd)">&nbsp;<INPUT NAME="txtItemGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>품목명</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" TYPE="Text" SIZE=30 MAXLENGTH="50" ALT="품목명" tag="11"></TD>
						<TD CLASS=TD5 NOWRAP>품목계정</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtJnlItem" TYPE="Text" MAXLENGTH="20" SIZE=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnJnlItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenJnlItem">&nbsp;
							<INPUT NAME="txtJnlItemNm" TYPE="Text" SIZE=20 tag="24">
						</TD>
						
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>규격</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemSpec" TYPE="Text" SIZE=30 MAXLENGTH="50" ALT="규격" tag="11"></TD>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>												
					</TR>	
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3112pa2_vaSpread_vspdData.js'></script>
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
					<TD>&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"     ONCLICK="FncQuery()"     ></IMG>
									<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG></TD>
					<TD ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"        ONCLICK="OkClick()"      ></IMG>
							        <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"    ONCLICK="CancelClick()"  ></IMG></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
