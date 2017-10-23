<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S3112RA4
'*  4. Program Name         : 수주내역참조(반품)
'*  5. Program Desc         : 수주내역참조(반품)
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
Dim lgMark                                                  

Dim arrReturn					

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
Const BIZ_PGM_ID        = "s3112rb4.asp"
Const C_MaxKey          = 30                                   '☆☆☆☆: Max key value
Const gstPaytermsMajor = "B9004"                                         
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'==============================================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
	lgPageNo         = ""
	Redim arrReturn(0, 0)
	Self.Returnvalue = arrReturn
End Sub

'==============================================================================================================
Sub SetDefaultVal()	
	Dim strArguments
	strArguments = Split(arrParent(1), PopupParent.gRowSep)

	' 박정순 수정 (2007-03-16) 거래처 , 화폐단위 Default 처리.
	frm1.txtSONo.value = strArguments(0)
	frm1.txtSoldtoParty.value = strArguments(1)
	frm1.txtSoldtoPartyNm.value = strArguments(2)
	frm1.txtHCur.value = strArguments(4)

	frm1.txtSOFrDt.text = StartDate
	frm1.txtSOToDt.text = EndDate
End Sub

'==============================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'==============================================================================================================
Sub InitSpreadSheet()	
	Call SetZAdoSpreadSheet("S3112ra4","S","A","V20030408", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
	Call SetSpreadLock           
End Sub


'==============================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	.vspdData.OperationMode = 5
    .vspdData.ReDraw = True
    End With
End Sub


'==============================================================================================================
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



'==============================================================================================================
Function OpenSONo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주번호"											
	arrParam(1) = "S_SO_HDR A, B_BIZ_PARTNER B, B_SALES_GRP C"				
	arrParam(2) = Trim(frm1.txtSONo.value)									
	arrParam(3) = ""														
	arrParam(4) = "A.SOLD_TO_PARTY = B.BP_CD AND A.SALES_GRP = C.SALES_GRP" 
	arrParam(5) = "수주번호"											

	arrField(0) = "ED12" & PopupParent.gColSep & "A.SO_NO"								
	arrField(1) = "ED10" & PopupParent.gColSep & "A.SOLD_TO_PARTY"						
	arrField(2) = "ED15" & PopupParent.gColSep & "B.BP_NM"								
	arrField(3) = "DD10" & PopupParent.gColSep & "CONVERT(char(11),A.SO_DT)"			
	arrField(4) = "ED15" & PopupParent.gColSep & "C.SALES_GRP_NM"						
	arrField(5) = "ED10" & PopupParent.gColSep & "A.PAY_METH"							
		
	arrHeader(0) = "수주번호"											
	arrHeader(1) = "주문처"												
	arrHeader(2) = "주문처명"											
	arrHeader(3) = "수주일"												
	arrHeader(4) = "영업그룹명"											
	arrHeader(5) = "결제방법"											

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSONo(arrRet)
	End If
End Function

'==============================================================================================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "주문처"							
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.txtSoldtoParty.value)		
	arrParam(3) = ""									
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
	arrParam(5) = "주문처"							
		
	arrField(0) = "BP_CD"								
	arrField(1) = "BP_NM"								
		
	arrHeader(0) = "주문처"							
	arrHeader(1) = "주문처명"						
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function


'==============================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "품목"							
	arrParam(1) = "B_ITEM"								
	arrParam(2) = Trim(frm1.txtItem.value)				
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "품목"							
		
	arrField(0) = "ITEM_CD"								
	arrField(1) = "ITEM_NM"								
		
	arrHeader(0) = "품목"							
	arrHeader(1) = "품목명"							
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItem(arrRet)
	End If
End Function

'==============================================================================================================
Function OpenSoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주형태"					
	arrParam(1) = "S_SO_TYPE_CONFIG"				
	arrParam(2) = Trim(frm1.txtSoType.value)		
	arrParam(3) = Trim(frm1.txtSoTypeNm.value)		
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "				
	arrParam(5) = "수주형태"					
		
    arrField(0) = "ED10" & PopupParent.gColSep & "SO_TYPE"			
    arrField(1) = "ED20" & PopupParent.gColSep & "SO_TYPE_NM"		
    arrField(2) = "ED9" & PopupParent.gColSep & "EXPORT_FLAG"		
    arrField(3) = "ED9" & PopupParent.gColSep & "RET_ITEM_FLAG"	
    arrField(4) = "ED15" & PopupParent.gColSep & "AUTO_DN_FLAG"	
	    
    arrHeader(0) = "수주형태"					
    arrHeader(1) = "수주형태명"					
    arrHeader(2) = "수출여부"					
    arrHeader(3) = "반품여부"					
    arrHeader(4) = "자동출하생성여부"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=570px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSoType.value = arrRet(0)
		frm1.txtSoTypeNm.value = arrRet(1)
		frm1.txtSoType.focus
	End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenTrackingNo()
	Dim iCalledAspName
	Dim strRet
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	'2002-10-07 s3135pa1.asp 추가 
	Dim arrTNParam(5), i

	If Len(frm1.txtSoldtoParty.value) Then
		arrTNParam(0) = frm1.txtSoldtoParty.value
	End If
	
	If Len(frm1.txtItem.value) Then
		arrTNParam(3) = frm1.txtItem.value
	End If
	
	If Len(frm1.txtSONo.value) Then
		arrTNParam(4) = frm1.txtSONo.value
	End If

	arrTNParam(5) = "SO"
		
	iCalledAspName = AskPRAspName("s3135pa3")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa3", "x")
		lgIsOpenPop = False
		exit Function
	end if
	lgIsOpenPop = True

	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrTNParam), _
		"dialogWidth=655px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo.value = strRet 
	End If		
		
	frm1.txtTrackingNo.focus
End Function	


'==============================================================================================================
Function SetSONo(arrRet)
	frm1.txtSONo.Value = arrRet(0)
End Function

'==============================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtSoldtoParty.value = arrRet(0)
	frm1.txtSoldtoPartyNm.value = arrRet(1)
	frm1.txtSoldtoParty.focus
End Function

'==============================================================================================================
Function SetItem(arrRet)
	frm1.txtItem.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItem.focus
End Function

'==============================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                     

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'==============================================================================================================
Sub txtSOFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOFrDt.Action = 7 
        Call SetFocusToDocument("P")
       frm1.txtSOFrDt.Focus
    End If
End Sub

'==============================================================================================================
Sub txtSOToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSOToDt.Action = 7 
        Call SetFocusToDocument("P")
		frm1.txtSOToDt.Focus
    End If
End Sub

'==============================================================================================================
Sub txtSOFrDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtSoToDt_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'==============================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
    If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then  
        Exit Function
    End If
        
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'==============================================================================================================
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
	
'==============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
	If OldLeft <> NewLeft Then Exit Sub    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
		
		If lgPageNo <> "" Then
	       Call DBQuery          
		End If
	End If    
End Sub

'==============================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'==============================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1

			frm1.vspdData.Row = intRowCnt + 1

			If frm1.vspdData.SelModeSelected Then	
					
				For intColCnt = 1 To frm1.vspdData.MaxCols - 1						
					frm1.vspdData.Col = GetKeyPos("A", intColCnt)
					arrReturn(intInsRow, intColCnt - 1) = frm1.vspdData.Text						
				Next			

				intInsRow = intInsRow + 1

			End IF
		Next
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'==============================================================================================================
Function CancelClick()
	Self.Close()
End Function


'==============================================================================================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function

    Call DbQuery															

    FncQuery = True		
End Function

'==============================================================================================================
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
		strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)				
		strVal = strVal & "&txtSoldtoParty=" & Trim(frm1.txtSoldtoParty.value)
		strVal = strVal & "&txtItem=" & Trim(frm1.txtITem.value)
		strVal = strVal & "&txtSOFrDt=" & Trim(frm1.txtSOFrDt.text)
		strVal = strVal & "&txtSoToDt=" & Trim(frm1.txtSoToDt.text)
		strVal = strVal & "&txtSoType=" & Trim(frm1.txtSoType.value)
		strVal = strVal & "&txtHCur=" & Trim(frm1.txtHCur.value)
		
		'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        strVal = strVal & "&lgPageNo="   & lgPageNo      
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")

        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True


End Function

'==============================================================================================================
Function DbQueryOk()														
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
						<TD CLASS=TD5 NOWRAP>수주번호</TD>
						<TD CLASS=TD6 NOWRAP>							
							<INPUT TYPE=TEXT NAME="txtSONo" SIZE=20 MAXLENGTH=18 TAG="14XXXU" ALT="수주번호"></TD>
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSoldtoParty" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" OnClick="vbscript:OpenBizPartner">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSoldtoPartyNm" SIZE=20 TAG="14">
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>수주일</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSOFrDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="수주시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtSOToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="수주종료일"></OBJECT>');</SCRIPT>
						</TD>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtITem" SIZE=20 MAXLENGTH=18 TAG="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" OnClick="vbscript:OpenItem">&nbsp;
							<INPUT TYPE=TEXT NAME="txtITemNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>수주형태</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoType" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoType">&nbsp;<INPUT NAME="txtSoTypeNm" TYPE="Text" SIZE=20 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=30 MAXLENGTH=25 TAG="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTRno" align=top TYPE="BUTTON" OnClick="vbscript:OpenTrackingNo">&nbsp;</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCur" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
