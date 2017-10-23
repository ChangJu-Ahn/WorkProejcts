<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'**********************************************************************************************
'*  1. Module Name          : Cast
'*  2. Function Name        :
'*  3. Program ID           : P6110Mma1
'*  4. Program Name         : 금형점검계획수립
'*  5. Program Desc         : 금형점검계획수립
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/19
'*  8. Modified date(Last)  : 2005/01/21
'*  9. Modifier (First)     : Lee sang-ho
'* 10. Modifier (Last)      : Lee Sang-hO
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID  = "P6210bb1.asp"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
<%'========================================================================================================%>
Dim IsOpenPop          

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtProd_Dt_Fr.focus

	frm1.txtProd_Dt_Fr.Year = strYear 		 '년월일 default value setting
	frm1.txtProd_Dt_Fr.Month = strMonth 
	frm1.txtProd_Dt_Fr.Day = "01"
	
	frm1.txtProd_Dt_To.Year = strYear 		 '년월일 default value setting
	frm1.txtProd_Dt_To.Month = strMonth 
	frm1.txtProd_Dt_To.Day = strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "BA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
<%'========================================================================================================%>

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
'    Call ggoOper.FormatDate(frm1.txtFinAj_dt, Parent.gDateFormat, 1)
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtCastCd.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()


End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
End Function

'========================================================================================================
' Name : OpenCast_Popup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenCast_Popup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
		Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
		Case "1"
			arrParam(0) = "금형코드"					' 팝업 명칭 
			arrParam(1) = "Y_CAST"							' TABLE 명칭 
			arrParam(2) = Trim(frm1.txthCast_Cd.Value)		' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "금형코드"					' TextBox 명칭 
	
			arrField(0) = "ED15" & parent.gcolsep & "CAST_CD"							' Field명(0)
			arrField(1) = "ED15" & parent.gcolsep & "CAST_NM"							' Field명(1)
			arrField(2) = "ED20" & parent.gcolsep & "(SELECT ITEM_GROUP_NM FROM B_ITEM_GROUP WHERE ITEM_GROUP_CD = CAR_KIND )"						' Field명(2)
			arrField(3) = "ED20" & parent.gcolsep & "(SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD = ITEM_CD_1 )"						' Field명(3)
			arrField(4) = "F3"   & parent.gcolsep & "EXT1_QTY"						' Field명(4)
    
			arrHeader(0) = "금형코드"					' Header명(0)
			arrHeader(1) = "금형코드명"					' Header명(1)
			arrHeader(2) = "모델명"						' Header명(2)
			arrHeader(3) = "품목명"						' Header명(3)
			arrHeader(4) = "차수"						' Header명(4)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
	     Frm1.txthCast_Cd.focus
		 Exit Function
	Else
		 Call SetCondArea(arrRet,iWhere)
	End If 
 
End Function

'======================================================================================================
' Name : SetCondArea()           
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCondArea(Byval arrRet, Byval iWhere) 
	With Frm1
		Select Case iWhere
			Case "1"
			    .txthCast_Cd.value = arrRet(0)
			    .txthCast_Nm.value = arrRet(1)
		End Select
	End With
End Sub


'========================================================================================================
'   Event Name : txthCast_Cd_Onchange()            '<==코드만 입력해도 앤터키,탭키를 치면 코드명을 불러준다 
'   Event Desc :
'========================================================================================================
Function txthCast_Cd_Onchange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txthCast_Cd.value = "" THEN
        frm1.txthCast_Nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" Cast_Nm "," Y_CAST "," Cast_cd = " & FilterVar(frm1.txthCast_Cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 'unicode
        If IntRetCd = false then
			Call DisplayMsgBox("Y60040","X","X","X")
            frm1.txthCast_Nm.value = ""
            frm1.txthCast_Cd.focus
            txthCast_Cd_Onchange = true
        ELSE    
            frm1.txthCast_Nm.value = Trim(Replace(lgF0,Chr(11),""))
        END IF
    END IF 
End Function 

'======================================================================================================
' Function Name : ExeReflect
' Function Desc : 
'=======================================================================================================
Function ExeReflect()
	Dim strVal
	Dim strFinAj_dt, stracct_dt, strPlantCd
	Dim IntRetCD
    Dim tempStr
    Dim tmpCastCd
    Dim strReport_dt_fr, strReport_dt_to

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then	
       Call BtnDisabled(0)
       Exit Function            								         '☜: This function check required field
    End If
    
    if txthCast_Cd_Onchange() then
		Exit Function
	end if

    IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If   LayerShowHide(1) = False Then
	     Call BtnDisabled(0)
	     Exit Function
	End If
	
	Call BtnDisabled(1) 
	
    strPlantCd = Trim(frm1.txtPlantCd.value)
    
    strReport_dt_fr = UniConvDateToYYYYMMDD(frm1.txtProd_dt_fr.text, Parent.gDateFormat, Parent.gComDateType)
    strReport_dt_to = UniConvDateToYYYYMMDD(frm1.txtProd_dt_to.text, Parent.gDateFormat, Parent.gComDateType)
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0006
	strVal = strVal & "&txtPlantCd=" & strPlantCd
	strVal = strVal & "&txthCast_Cd=" & Trim(frm1.txthCast_Cd.value)
	strVal = strVal & "&txtReportDtFr=" & Trim(strReport_dt_fr)
	strVal = strVal & "&txtReportDtTo=" & Trim(strReport_dt_to)

	If IsNull(frm1.txthCast_Cd.value) Or Trim(frm1.txthCast_Cd.value) = "" Then
		tmpCastCd = "%"
	Else
		tmpCastCd = frm1.txthCast_Cd.value
	End If

	Call CommonQueryRs("COUNT(*)","Y_CAST","MAKE_DT <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtProd_dt_fr.text, Parent.gDateFormat, Parent.gComDateType), "''", "s") & " AND CAST_CD LIKE " & FilterVar(tmpCastCd, "''", "S") & " AND (CLOSE_DT IS NULL OR CLOSE_DT = " & FilterVar("1900-01-01", "''", "S")  & ") " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If Replace(lgF0, Chr(11), "") <= 0  then
   	   Call DisplayMsgBox("800076","x","x","x")
   	   Call BtnDisabled(0)
   	   Call LayerShowHide(0)
   	   Exit Function
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
	Call BtnDisabled(0)
End Function

'======================================================================================================
' Function Name : ExeReflectOk
' Function Desc : ExeReflect가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End Function

Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
End Function

Sub txtFinAj_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtFinAj_dt.Action = 7
		frm1.txtFinAj_dt.focus
	End If
End Sub

 '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True   Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_Plant"


	arrParam(2) = Trim(frm1.txtPlantCd.Value)

	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_CD"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		
		frm1.txtPlantCd.focus
		
		Exit Function
	Else
		frm1.txtPlantCd.Value  = arrRet(0)
		frm1.txtPlantNm.Value  = arrRet(1)
		frm1.txtPlantCd.focus
	End If
End Function

'==========================================================================================
'   Event Name : txtProd_dt_Fr
'   Event Desc :
'==========================================================================================

 Sub txtProd_Dt_Fr_DblClick(Button)
	if Button = 1 then
		frm1.txtProd_dt_Fr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtProd_dt_Fr.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtProd_dt_To
'   Event Desc :
'==========================================================================================

 Sub txtProd_dt_To_DblClick(Button)
	if Button = 1 then
		frm1.txtProd_dt_To.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtProd_dt_To.Focus
	End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
 

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>타수적용</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>						
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
									<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>금형코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txthCast_Cd" MAXLENGTH="18"  SIZE="18" ALT ="금형코드" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCast_Popup('1')">
								                       <INPUT NAME="txthCast_Nm" MAXLENGTH="40" SIZE="25" ALT ="금형명" tag="14"></TD>

	                        </TR>
	                        <TD CLASS="TD5" NOWRAP>실적일자</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="시작일" NAME="txtProd_Dt_Fr" CLASSID=<%=gCLSIDFPDT%> id=OBJECT1 style="HEIGHT: 20px; WIDTH: 100px" tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
													
												</td>
												<td>&nbsp;~&nbsp;</td>
												<td>
												    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="시작일" NAME="txtProd_Dt_To" CLASSID=<%=gCLSIDFPDT%> id=OBJECT2 style="HEIGHT: 20px; WIDTH: 100px" tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											<tr>
										</table>

									</TD>
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
  					     <BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1 SRC="../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


