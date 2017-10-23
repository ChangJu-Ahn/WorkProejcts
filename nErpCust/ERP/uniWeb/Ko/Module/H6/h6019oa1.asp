<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 급여관리 
'*  3. Program ID           : h6019oa1
'*  4. Program Name         : 은행이체LIST출력 
'*  5. Program Desc         : 은행이체LIST출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : Shin Kwang-Ho
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
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
	
	frm1.txtPay_yymm.Focus			'년월 default value setting
	frm1.txtPay_yymm.Year = strYear 
	frm1.txtPay_yymm.Month = strMonth

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
	Call ggoOper.FormatDate(frm1.txtPay_yymm, Parent.gDateFormat, 2)

    Call InitVariables 
        
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")
 
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
        
    Err.Clear                                                                    '☜: Clear err status
    
    FncQuery = true
    With frm1

        If txtProv_type_OnChange()  Then
            Exit Function
        End If
         If txtFr_bank_cd_Onchange()  Then
            Exit Function
        End If
        If txtTo_bank_cd_Onchange()  Then
            Exit Function
        End If               
        
    End With    
End Function

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

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case "FR_BANK_POP"
	        arrParam(0) = "은행코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_bank"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtFr_bank_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtFr_bank_nm.value									' Name Cindition
	    	arrParam(4) = ""	                		    	' Where Condition
	    	arrParam(5) = "은행코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "bank_cd"						    	' Field명(0)
	    	arrField(1) = "bank_full_nm"    				  	' Field명(1)
	    	arrField(2) = "bank_nm"    				        	' Field명(2)
    
	    	arrHeader(0) = "은행코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "은행명"	          		        ' Header명(1)
	    	arrHeader(2) = "은행약어명"	    		        ' Header명(1)
	    Case "TO_BANK_POP"
	        arrParam(0) = "은행코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_bank"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtTo_bank_cd.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtFr_bank_nm.value 				' Name Cindition
	    	arrParam(4) = ""	                		    	' Where Condition
	    	arrParam(5) = "은행코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "bank_cd"						    	' Field명(0)
	    	arrField(1) = "bank_full_nm"    			    	' Field명(1)
	    	arrField(2) = "bank_nm"           					' Field명(2)
    
	    	arrHeader(0) = "은행코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "은행명"	          		        ' Header명(1)
	    	arrHeader(2) = "은행약어명"	    		        ' Header명(1)
	    Case "PROV_TYPE"
			arrParam(0) = "지급구분 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtProv_type.value     			' Code Condition
	    	arrParam(3) = ""'frm1.txtProv_type_nm.value				' Name Cindition
	    	arrParam(4) = "major_cd = " & FilterVar("H0040", "''", "S") & ""	   		    	' Where Condition
	    	arrParam(5) = "지급코드"  			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"						    ' Field명(0)
	    	arrField(1) = "minor_nm"    					  	' Field명(1)
	    	arrField(2) = ""    				        		' Field명(2)
    
	    	arrHeader(0) = "지급구분코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "지급구분코드명"        		        ' Header명(1)
	    	arrHeader(2) = ""	    							' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then

		Select Case iWhere
		    Case "FR_BANK_POP"
		    	frm1.txtFr_bank_cd.focus
   	        Case "TO_BANK_POP"
		    	frm1.txtTo_bank_cd.focus
			Case "PROV_TYPE"
				frm1.txtProv_type.focus
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case "FR_BANK_POP"
		        .txtFr_bank_cd.value = arrRet(0) 
		    	.txtFr_bank_nm.value = arrRet(1) 
		    	.txtFr_bank_cd.focus
   	        Case "TO_BANK_POP"
   				.txtTo_bank_cd.value = arrRet(0) 
		    	.txtTo_bank_nm.value = arrRet(1) 
		    	.txtTo_bank_cd.focus
			Case "PROV_TYPE"
				.txtProv_type.value = arrRet(0) 
				.txtProv_type_nm.value  = arrRet(1) 
				.txtProv_type.focus
        End Select
	End With

End Function
'========================================================================================
' Function Name : FncBtnPrint()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPrint() 


	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName
	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Call BtnDisabled(0)
	   Exit Function
    End If

	dim pay_yymm, prov_type, gigup_type1, fr_bank_cd, to_bank_cd, stand_amt 
	
	StrEbrFile = "h6019oa1"
	
    Pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)

	prov_type = frm1.txtProv_type.value
	fr_bank_cd = frm1.txtFr_bank_cd.value
	to_bank_cd = frm1.txtTo_bank_cd.value
	stand_amt = UNICDbl(frm1.txtStand_amt.Text)
	stand_amt = Replace(stand_amt, Parent.gClientNumDec, ".")

	If frm1.txtGigup_type(0).checked Then 
		gigup_type1 = "1"
	Elseif frm1.txtGigup_type(1).checked Then 
		gigup_type1 = "2"
	Else
		gigup_type1 = "3"		
	End if		
	
	If IsNull(gigup_type1) Or Trim(gigup_type1) = "" Or gigup_type1 = "3" Then
	    gigup_type1 = "1"	    
	    stand_amt = 0
	End If
	
        If txtProv_type_OnChange()  Then
            Exit Function
        End If
         If txtFr_bank_cd_Onchange()  Then
            Exit Function
        End If
        If txtTo_bank_cd_Onchange()  Then
            Exit Function
        End If  

	if fr_bank_cd = "" then
		fr_bank_cd = " "
		frm1.txtFr_bank_nm.value = ""
	End if	
	
	if to_bank_cd = "" then
		to_bank_cd = "ZZZZZZZZZZ"
		frm1.txtTo_bank_nm.value = ""
	End if		
    
    If (fr_bank_cd= "") AND (to_bank_cd="") Then       
    Else
        If fr_bank_cd > to_bank_cd then
	        Call DisplayMsgbox("970025","X","시작은행코드","종료은행코드")	'시작은행은 종료은행보다 작아야 합니다.
			frm1.txtFr_bank_cd.focus
            Set gActiveElement = document.activeElement
			Call BtnDisabled(0)
            Exit Function
        End IF 
        
    END IF   

  	Call BtnDisabled(1)  

	strUrl = "pay_yymm|" & pay_yymm
	strUrl = strUrl & "|prov_type|" & prov_type
	strUrl = strUrl & "|gigup_type1|" & gigup_type1
	strUrl = strUrl & "|fr_bank|" & fr_bank_cd
	strUrl = strUrl & "|to_bank|" & to_bank_cd
	strUrl = strUrl & "|stand_amt|" & stand_amt
   
    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

 	call FncEBRPrint(EBAction , ObjName , strUrl)

End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview()
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile, ObjName
		
	dim pay_yymm, prov_type, gigup_type1, fr_bank_cd, to_bank_cd, stand_amt 


    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Call BtnDisabled(0)
	   Exit Function
    End If
	
	StrEbrFile = "h6019oa1"
	
    Pay_yymm = frm1.txtPay_yymm.year & Right("0" & frm1.txtPay_yymm.month , 2)

	prov_type = frm1.txtProv_type.value
	fr_bank_cd = frm1.txtFr_bank_cd.value
	to_bank_cd = frm1.txtTo_bank_cd.value
	stand_amt = UNICDbl(frm1.txtStand_amt.Text)
	stand_amt = Replace(stand_amt, Parent.gClientNumDec, ".")

	If frm1.txtGigup_type(0).checked Then 
		gigup_type1 = "1"
	Elseif frm1.txtGigup_type(1).checked Then 
		gigup_type1 = "2"
	Else 
		gigup_type1 = "3"
	End if		
	
	If IsNull(gigup_type1) Or Trim(gigup_type1) = "" Or gigup_type1 = "3" Then
	    gigup_type1 = "1"	
	    stand_amt = 0
	End If
	
    If txtProv_type_OnChange()  Then
        Exit Function
    End If
     If txtFr_bank_cd_Onchange()  Then
        Exit Function
    End If
    If txtTo_bank_cd_Onchange()  Then
        Exit Function
    End If  

	if fr_bank_cd = "" then
		fr_bank_cd = " "
		frm1.txtFr_bank_nm.value = ""
	End if	
	
	if to_bank_cd = "" then
		to_bank_cd = "ZZZZZZZZZZ"
		frm1.txtTo_bank_nm.value = ""
	End if	

    If (fr_bank_cd= "") AND (to_bank_cd="") Then       
    Else
        If fr_bank_cd > to_bank_cd then
	        Call DisplayMsgbox("970025","X","시작은행코드","종료은행코드")	'시작은행은 종료은행보다 작아야 합니다.
			frm1.txtFr_bank_cd.focus
            Set gActiveElement = document.activeElement
			Call BtnDisabled(0)
            Exit Function
        End IF 
        
    END IF     

	Call BtnDisabled(1)
    
	strUrl = "pay_yymm|" & pay_yymm
	strUrl = strUrl & "|prov_type|" & prov_type
	strUrl = strUrl & "|gigup_type1|" & gigup_type1
	strUrl = strUrl & "|fr_bank|" & fr_bank_cd
	strUrl = strUrl & "|to_bank|" & to_bank_cd
	strUrl = strUrl & "|stand_amt|" & stand_amt
   
    ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	call FncEBRPreview(ObjName , strUrl)

End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
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
	FncExit = True
End Function


'========================================================================================================
'   Event Name : txtFr_bank_cd_Onchange
'   Event Desc : 은행코드에러체크 
'========================================================================================================
Function txtFr_bank_cd_onchange()

    Dim IntRetCd
    
    If frm1.txtFr_bank_cd.value = "" Then
		frm1.txtFr_bank_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" bank_nm "," B_BANK "," bank_cd =  " & FilterVar(frm1.txtFr_bank_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("800137", "x","x","x")                   
            frm1.txtFr_bank_nm.value = ""            
            frm1.txtFr_bank_cd.focus 
            Set gActiveElement = document.ActiveElement
            txtFr_bank_cd_onchange = true
            Exit Function
        Else
            frm1.txtFr_bank_nm.value = Trim(Replace(lgF0, Chr(11), ""))
        End If
    End If
    
End Function

'========================================================================================================
'   Event Name : txtTo_bank_cd_Onchange
'   Event Desc : 은행코드에러체크 
'========================================================================================================
Function txtTo_bank_cd_onchange()

    Dim IntRetCd    
    
    If frm1.txtTo_bank_cd.value = "" Then
		frm1.txtTo_bank_nm.value = ""
	ELSE
        IntRetCd = CommonQueryRs(" bank_nm "," B_BANK "," bank_cd =  " & FilterVar(frm1.txtTo_bank_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("800137", "x","x","x")                   
            frm1.txtTo_bank_nm.value = ""            
            frm1.txtTo_bank_cd.focus 
            Set gActiveElement = document.ActiveElement
            txtTo_bank_cd_onchange = true
            Exit Function
        Else
            frm1.txtTo_bank_nm.value = Trim(Replace(lgF0, Chr(11), ""))
        End If
    End If
    
End Function
'======================================================================================================
'   Event Name : txtProv_type_OnChange
'   Event Desc : 지급구분 에러체크 
'=======================================================================================================
Function txtProv_type_OnChange()

    Dim iDx
    Dim IntRetCd
    
    If Trim(frm1.txtProv_type.value) = "" Then
        frm1.txtProv_type_nm.value = ""
    ELSE
        IntRetCd = CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtProv_type.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        IF IntRetCd = False THEN
            Call DisplayMsgBox("800054", "x","x","x")   '등록되지 않은 코드입니다                
	        frm1.txtProv_type_nm.value = ""    	        
	        frm1.txtProv_type.focus 
	        Set gActiveElement = document.ActiveElement	        
	        txtProv_type_OnChange = true
	        Exit Function
	    Else
	        frm1.txtProv_type_nm.value = Trim(Replace(lgF0, Chr(11), ""))
	    End If
    End If	
End Function

'========================================================================================================
' Name : txtPay_yymm_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtPay_yymm_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")
		frm1.txtPay_yymm.Action = 7
		frm1.txtPay_yymm.focus
	End If
End Sub
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
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>은행이체LIST출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
								<TD CLASS=TD5  NOWRAP>해당년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h6019oa1_txtPay_yymm_txtPay_yymm.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtProv_type" NAME="txtProv_type" SIZE=10 MAXLENGTH=1 tag="12XXXU" ALT="지급구분코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'PROV_TYPE', 'x')">
								                       <INPUT TYPE="Text" NAME="txtProv_type_nm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="지급구분코드명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>은행코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_bank_cd" NAME="txtFr_bank_cd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="시작은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'FR_BANK_POP', 'x')">
								                       <INPUT TYPE="Text" NAME="txtFr_bank_nm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="시작은행코드명">&nbsp;~&nbsp;</TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtTo_bank_cd" NAME="txtTo_bank_cd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="종료은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCode('x', 'TO_BANK_POP', 'x')">
								                       <INPUT TYPE="Text" NAME="txtTo_bank_nm" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="종료은행코드명">&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급방식</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type VALUE = "1" ID=Rb_tot Checked tag="12"><LABEL FOR=Rb_tot>기준금액 제외한 은행이체</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type VALUE = "2" ID=Rb_dur tag="12"><LABEL FOR=Rb_dur>기준금액 미만 금액만 은행이체</LABEL></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtGigup_type VALUE = "3" ID=Rb_dept tag="12"><LABEL FOR=Rb_dept>모든 금액 은행이체</LABEL></TD>
							</TR>
							
	    					<TR>
              				    <TD CLASS="TD5" NOWRAP>기준금액</TD>
	                   			<TD CLASS="TD6"><script language =javascript src='./js/h6019oa1_txtStand_amt_txtStand_amt.js'></script>&nbsp;원</TD>
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>

