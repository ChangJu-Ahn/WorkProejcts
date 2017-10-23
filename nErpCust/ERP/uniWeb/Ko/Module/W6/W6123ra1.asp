
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<!--'**********************************************************************************************
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
'*
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Common.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->		
Dim arrParent
dIM arrReturn		
Dim iStrFlag

Dim sFiscYear, sRepType, sCoCd,strMajor,strMinor
	sCoCd		  = Trim("<%=Request("sCoCd")%>")
	sFiscYear	  = Trim("<%=Request("sFiscYear")%>")
	sRepType	  = Trim("<%=Request("sRepType")%>")
	strMajor      = Trim("<%=Request("strMajor")%> ")

	

arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)

	top.document.title = "계산내역"

Self.Returnvalue = Array("CANCEL")

'=================================================================================================
Function OKClick()
    Dim strVal

		Redim arrReturn(5)
		
			
				arrReturn(0) =  txtTAX.text         '감면세액 
				arrReturn(1) = txtAmt.text         '감면소득 
				arrReturn(2) = txtw3_113.text    '과세표준 
				arrReturn(3) = txtwRateValue.value  ' 감면율 
				arrReturn(4) = txtwRateView.value    '감면율 

				arrReturn(5) = txtw3_120.text    '산출세액 

		'3호 서식 (120) 산출세액 × (감면(면제)소득/3호서식 (113) 과세표준)× 	감면율 
		Self.Returnvalue = arrReturn
	
		CALL CancelClick()
		
End Function


Function InitVariables()
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
	
	'========================================================================================================
' Name : OpenCon()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenCon(Byval iWhere)
    Dim arrRet,IsOpenPop ,strFrom
    Dim arrParam(5), arrField(6), arrHeader(6) ,strWhere

    If IsOpenPop = True  Then
       Exit Function
    End If
    
    IsOpenPop = True
    Select Case iWhere
       Case "1"
       strWhere = "Co_cd ='" & sCoCd &"' and   fisc_year  ='" & sFiscYear &"' and   rep_type  ='" & sRepType &"' "

			strFrom =		"("
			strFrom = strFrom  & "(select '감면사업1'  nm,  w4_1 data   "
			strFrom = strFrom  & "	 from TB_48H2   "
			strFrom = strFrom  & "	 where w1_cd='25' and " & strWhere &" )"
				
			strFrom = strFrom  & "	union "
			strFrom = strFrom  & "	(select '감면사업2'  nm,  w4_2 data "
			strFrom = strFrom  & "	 from TB_48H2  where w1_cd='25' and  " & strWhere &"  )"
				
			strFrom = strFrom  & "	union "
			strFrom = strFrom  & "	(select '감면사업3'  nm,  w4_3  data "
			strFrom = strFrom  & "	 from TB_48H2  where w1_cd='25' and " & strWhere &"  ) "
				
			strFrom = strFrom  & "	union "
			strFrom = strFrom  & "	(select '감면사업4'  nm,   w4_4   data "
			strFrom = strFrom  & "	 from TB_48H2  where w1_cd='25' and  " & strWhere &"  ) "
				
			strFrom = strFrom  & "	union "
			strFrom = strFrom  & "	(select '감면사업5'  nm,   w4_5 data  "
			strFrom = strFrom  & " 	 from TB_48H2  where w1_cd='25' and  " & strWhere &"  ) "
				
			strFrom = strFrom  & "	union "
			strFrom = strFrom  & "	(select '감면사업6'  nm,   w4_6 data "
			strFrom = strFrom  & "	  from TB_48H2  where w1_cd='25'  and  " & strWhere &"  ) "
			strFrom = strFrom  & "  ) k "

 
 
 
           arrParam(0) = "감사사업소득"                         ' 팝업 명칭 
           arrParam(1) = strFrom                                        ' TABLE 명칭 
           arrParam(2) = ""                ' Code Condition
           arrParam(3) = ""                                        ' Name Cindition
           arrParam(4) = ""                                          ' Where Condition
           arrParam(5) = ""                        ' TextBox 명칭 
           
           arrField(0) = "nm"                     ' Field명(0)
           arrField(1) = "data"                     ' Field명(1)
           
           arrHeader(0) = ""                 ' Header명(0)
           arrHeader(1) = ""                 ' Header명(1)
    End Select
        
        
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetCond(arrRet,iWhere)
    End If
    
End Function


'======================================================================================================
' Name : SetCondArea()           
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCond(Byval arrRet, Byval iWhere)
        Select Case iWhere
            Case "1"
                txtAmt.text = arrRet(1)
                
                call txtAmt_Change()
        
        End Select
End Sub


Function txtAmt_Change()
       Call Fn_Recal()
End Function


Function Fn_Recal()
dim dblamt
   if  unicdbl(txtw3_113.text) <> 0 then
       dblAmt = (unicdbl(txtAmt.text) / unicdbl(txtw3_113.text) )       
   else
      dblAmt = 0
   end if    
       
        txtTax.value =unicdbl(txtw3_120.value) * dblAmt *  unicdbl(txtwRateValue.value)
        
        
        
        
End Function

Function cboRATE_OnChange()
    DIM  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim  strColNm , strTableid ,  strNmwhere ,strArr
		 strColNm  = "   reference_1 , reference_2 "
		 strTableid  = " ufn_TB_Configuration('" & strMajor &"','" & C_REVISION_YM & "')  "
		 strNmwhere = "  minor_cd  = '" & trim(cboRATE.value)  &"'   "

	 Call CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

      txtwRateView.value   = Trim(replace(lgF1,Chr(11),""))
      txtwRatevalue.value   = Trim(replace(lgF0,Chr(11),""))
      Call Fn_Recal
  
End Function




Function SetDefaultVal()
dim arrW1
dim arrW2
      call CommonQueryRs("W3_120,W3_113","dbo.ufn_TB_8_2_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

           

		If lgF0 = "" Then	 Exit Function

		    arrW1 = REPLACE(lgF0, chr(11),"")
		    arrW2 = REPLACE(lgF1, chr(11),"")
	        txtW3_120.value =  cdbl(arrW1)
		    txtW3_113.text =  cdbl(arrW2)
		   
 
       
End Function


Sub InitComboBox()
    DIM  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim  strColNm , strTableid ,  strNmwhere
	 strColNm  = "   b.minor_cd,b.minor_nm"
	 strTableid  = " ufn_TB_Configuration('" & strMajor &"','" & C_REVISION_YM & "') b "
	 strNmwhere = "  "

	 Call CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
     Call SetCombo2(cboRATE ,lgF0  ,lgF1  ,Chr(11))
End Sub



'===========
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub


'=================================================================================================
Function CancelClick()
	Self.Close()
End Function
	
'=================================================================================================
Sub Form_Load()
      
    CALL InitVariables
     Call GetGlobalVar()
                      
    Call LoadInfTB19029	
	Call ggoOper.LockField(Document, "N")  
	
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)

    Call SetDefaultVal 
     Call InitComboBox 
    txtAmt.focus
End Sub

'=================================================================================================
Sub Window_onLoad()
    Call Form_Load()    
End Sub


'=================================================================================================
Sub RunMyBizASP(objIFrame, strURL)
	Call BtnDisabled(True)
	objIFrame.location.href = GetUserPath & strURL

End Sub

'=================================================================================================
Function GetUserPath()
	If gURLPath = "" or isEmpty(gURLPath) Then
		Dim strLoc, iPos , iLoc, strPath
		strLoc = window.location.href
                iLoc = inStr(1, strLoc, "?")
            
                If iLoc > 0 Then
                   strLoc = Left(strLoc, iLoc - 1)
                End If
		
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		gURLPath = Left(strLoc, iPos)
	End If
	GetUserPath = gURLPath
End Function

Function document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function

Sub TextKeypress(pos)
	If window.event.keyCode = 13 Then
		Select Case pos
			Case 3
				Call OKClick()
		End Select
	End If
End sub
	

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=10 CLASS="basicTB">
	<TR>
		<TD HEIGHT=*>
			<FIELDSET>
			<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
	
	
	
	   <TR>
						<TD BGCOLOR=#d1e8f9 COLSPAN=2>
						<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
						 	<TR>
						 		<TD ALIGN=CENTER WIDTH=30%>3호(120)산출세액 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw3_120" name=txtw3_120 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 80%></OBJECT>');</SCRIPT></TD>
						 		<TD ALIGN=CETER WIDTH=30%>
						 		<TABLE  CLASS="BasicTB" CELLSPACING=0 border="0">
						 			<TR>
						 				<TD ROWSPAN=3 WIDTH=30% ALIGN=RIGHT> x&nbsp;&nbsp;</TD>
						 				<TD ALIGN=CENTER>감면(면제)소득 
						 				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtAmt" name=txtAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="22X2" width = 80%></OBJECT>');</SCRIPT><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAmt" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCon('1')">
						 				
						 				</TD>
						 				<TD ROWSPAN=3 ></TD>
						 			</TR>
						 			<TR>
						 				<TD HEIGHT=1 BGCOLOR=BLACK></TD>
						 			</TR>
						 			<TR>
						 				<TD ALIGN=CENTER>3호(113)과세표준						 				
						 				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtw3_113" name=txtw3_113 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 80%></OBJECT>');</SCRIPT></TD>
						 			</TR>
																
						 		</TABLE>
						 		<TD CLASS="TD5" NOWRAP>x&nbsp;&nbsp;</TD>
                                <TD CLASS="TD5" NOWRAP >감면율<SELECT NAME="CboRate" ALT="감면율" CLASS ="cbonormal" TAG="12" width =100><OPTION VALUE=0></OPTION></SELECT><INPUT NAME="txtwRateView" ALT="" TYPE="Text"  MAXLENGTH=50 SiZE=5 tag=24></TD>		
                        	
						 			<TD CLASS="TD6">=</TD>
						 		    <TD CLASS="TD6">감면대상세액<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtTax" name=txtTax CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" SiZE=5 ></OBJECT>');</SCRIPT></TD>
							 		
						 	</TR>
						 </TABLE></TD>															
												  
											
				</tr>							 
										
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtwRateValue"     TAG="24">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm" tabindex=-1></iframe>
</DIV>
</BODY>
</HTML>
