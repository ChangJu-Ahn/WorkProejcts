<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s1912ra2.asp	
'*  4. Program Name         : 재고현황참조 
'*  5. Program Desc         : 재고현황참조 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop 
Dim IscookieSplit 

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
Const BIZ_PGM_ID        = "S1912rb2.asp"
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'=============================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
	lgPageNo         = ""
	Redim arrReturn(0, 0)
	Self.Returnvalue = arrReturn
End Sub

'=============================================================================================================
Sub SetDefaultVal()
	
	'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	Dim strArguments
	strArguments = arrParent(1)

	frm1.txtItem.Value		= strArguments(0)
	frm1.txtItemNm.value	= strArguments(1)
	frm1.txtPlant.value		= strArguments(2)
	frm1.txtPlantNm.value	= strArguments(3)
	frm1.txtSL.value		= strArguments(4)
	frm1.txtSLNm.value		= strArguments(5)
	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
End Sub

'=============================================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
		'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=============================================================================================================
Sub InitSpreadSheet()		
	Call SetZAdoSpreadSheet("S1912ra2","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock     
End Sub

Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

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
Function OpenSoDtl(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim TempItem, TempPlant

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0	
		arrParam(1) = "b_item item, b_plant plant, b_item_by_plant item_plant"				
		arrParam(2) = Trim(frm1.txtItem.value)										
		arrParam(4) = "item.item_cd=item_plant.item_cd and plant.plant_cd=item_plant.plant_cd"
		arrParam(5) = "품목"		
	
		arrField(0) = "item.item_cd"	
		arrField(1) = "item.item_nm"	
		arrField(2) = "plant.plant_cd"	
		arrField(3) = "plant.plant_nm"	
    
		arrHeader(0) = "품목"		
		arrHeader(1) = "품목명"		
		arrHeader(2) = "공장"		
		arrHeader(3) = "공장명"		

	Case 1	
		TempItem = frm1.txtITem.value 

		arrParam(1) = "b_plant plant, b_item_by_plant item_plant"			
		arrParam(2) = Trim(frm1.txtPlant.value)								
		arrParam(4) = "plant.plant_cd=item_plant.plant_cd and item_plant.item_cd =  " & FilterVar(TempItem, "''", "S") & ","	
		arrParam(5) = "공장"						
	
		arrField(0) = "plant.plant_cd"					
		arrField(1) = "plant.plant_nm"					
		    
		arrHeader(0) = "공장"						
		arrHeader(1) = "공장명"						
		
	Case 2	
		If frm1.txtPlant.value = "" Then
			Call DisplayMsgBox("189220", "x", "x", "x")
			lgIsOpenPop = False
			Exit Function
		End If	

		TempPlant = frm1.txtPlant.value 

		arrParam(1) = "b_storage_location"				
		arrParam(2) = Trim(frm1.txtSL.value)			
		arrParam(4) = "plant_cd = " & FilterVar(TempPlant, "''", "S") & ","	
		arrParam(5) = "창고"						
	
		arrField(0) = "sl_cd"							
		arrField(1) = "sl_nm"							
    
		arrHeader(0) = "창고"						
		arrHeader(1) = "창고명"						
	End Select

	arrParam(0) = arrParam(5)							

	Select Case iWhere
	Case 0	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSoDtl(arrRet, iWhere)
	End If	
	
End Function

'=============================================================================================================
Function SetSoDtl(Byval arrRet,ByVal iWhere)
	With frm1
		Select Case iWhere
		Case 0	
			.txtItem.value = arrRet(0) 
			.txtItemNm.value = arrRet(1) 
		Case 1	
			.txtPlant.value = arrRet(0)
			.txtPlantNm.value = arrRet(1)
		Case 2	
			.txtSL.value = arrRet(0)
			.txtSLNm.value = arrRet(1)
		Case Else
			Exit Function
		End Select
		
	End With
End Function

'=============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format	
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub


'=============================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then			
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
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function


'=============================================================================================================
Function OKClick()
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=============================================================================================================
Function CancelClick()
	Self.Close()
End Function


'=============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
        
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=============================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtItem=" & Trim(frm1.txtItem.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
		strVal = strVal & "&txtSL=" & Trim(frm1.txtSL.value)
		
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag        
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		
		strVal = strVal & "&lgPageNo=" & lgPageNo
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
 
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function

'=============================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
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
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>품목</TD>
						<TD CLASS=TD6 Colspan = 3>
							<INPUT TYPE=TEXT NAME="txtItem" SIZE=18 MAXLENGTH=18 TAG="14XXXU" ALT="품목">&nbsp;
							<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=50 TAG="14" ALT="품목명">
						</TD>
					</TR>	
					<TR>	
						<TD CLASS=TD5>공장</TD>
						<TD CLASS=TD6>
							<INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" OnClick="vbscript:OpenSoDtl 1">&nbsp;
							<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5>창고</TD>
						<TD CLASS=TD6>
							<INPUT TYPE=TEXT NAME="txtSL" SIZE=10 MAXLENGTH=7 TAG="11XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" OnClick="vbscript:OpenSoDtl 2">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 TAG="14">
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
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s1912ra2_vaSpread_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                     <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHSL" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
