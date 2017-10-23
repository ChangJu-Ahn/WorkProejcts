<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 
*  2. Function Name        : importExcel
*  3. Program ID           : Import Excel file  Popup
*  4. Program Name         : ImportExcelfile Popup
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/07/5
*  8. Modified date(Last)  : 2002/07/5
*  9. Modifier (First)     : Lee Seok min
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/eventpopup.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js">  </SCRIPT>

<Script Language="VBScript">
Option Explicit            


Sub InitVariables()
    
End Sub



Sub LoadInfTB19029()
<!-- #Include file="ComLoadInfTB19029.asp" -->
End Sub



Sub InitSpreadSheet()
	frm1.vspdData.ReDraw = False
	frm1.vspddata.allowmultiBlocks=true
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit
    		    
    frm1.vspdData.OperationMode = 3

		
	frm1.vspdData.ReDraw = True
End Sub



Sub SetDefaultVal()
	Self.Returnvalue = ""
End Sub

Sub Form_Load()
		<% ' 이미지 효과 자바스크립트 함수 호출  %>
	Call MM_preloadImages("../image/Query.gif","../image/OK.gif","../image/Cancel.gif")
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
    
	Call InitVariables
				
	Call SetDefaultVal()
	Call InitSpreadSheet()
	
End Sub



sub btnCancel_Onclick()
	self.close
End sub

function GetLastRows()
	dim i
	
	i = 0
	
	do 
		i = i + 1
		frm1.vspdData.Row = i
		frm1.vspdData.col = 1
		
	loop until frm1.vspdData.Text = ""

	GetLastRows = i-1
	
end function




sub importExcel(FileName)
	DIM List()
	DIM intRet
	DIM sheet
	dim Rows ,i,j
	dim strData
	dim listcount,handle
	dim arrParent, arrparam, arrReturn(1)
	dim arrEx()	
	intRet=frm1.vspdData.ScriptGetExcelSheetList(FileName, Null, listcount, "", handle, true) 

	ReDim List(listcount)
	
	intRet=frm1.vspdData.ScriptGetExcelSheetList(FileName, List, listcount, "", handle, true)
	sheet=List(0)
	intRet=frm1.vspdData.ImportExcelSheet(cint(handle), sheet) 

	frm1.vspdData.redraw = true
	
	arrParent = window.dialogArguments
	arrParam = arrParent(0)
	
	Rows = GetLastRows
	
	for i = 2 to rows
		frm1.vspdData.Row = i
		
		for j = 1 to arrparam(0) 

			frm1.vspdData.col = j
			strData = strData & chr(11) & frm1.vspdData.Text 
			Redim arrEX(i-1,j-1)
			
			arrEx(i-1,j-1) = frm1.vspddata.text
		next
		
		strData = strData & gColSep  & Cstr(i) & gColSep  & gRowSep
	next
	
	'intRet=frm1.vspdData.setSelection(0,0,arrparam(0),Rows)
	'intRet = frm1.vspdData.ClipboardCopy 
	
	Self.Returnvalue = strData

	self.close 

End sub

Sub btnOk_Onclick()
	
	dim FileName
	dim intRet
	Dim iRet

	
	
	frm1.vspdData.redraw = false
	
	FileName = frm1.ExcelFile.value 
	
	frm1.vspdData.ScriptEnhanced=true 

	intRet = frm1.vspdData.IsExcelFile(FileName) 
	
	select case intRet
		case 0
		       iRet = MsgBox("EXCEL 파일이 아니므로 파일을 열 수 없습니다.", vbExclamation, gLogoName)
               Exit Sub
				
		case 2
		       iRet = MsgBox("다른 응용프로그램에서 파일을 사용중이므로 파일을 열 수 없습니다.", vbExclamation, gLogoName)
			   Exit Sub
	end select 
	LayerShowHide(1)
	call importExcel(FileName)
	LayerShowHide(0)
End sub


</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<BODY SCROLL=no TABINDEX="-1">
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="POST">

<TABLE CLASS="basicTB" CELLSPACING=0 >
	<TR >
		<TD  <%=HEIGHT_TYPE_00%>></TD>
	</TR>	

	<TR>
		<TD CLASS="Tab11">
		<FIELDSET >
			<TABLE>
				<TR>
					<TD CLASS="TD5">대상파일</TD>
					<TD CLASS="TD6"><INPUT type="file" id=ExcelFile name=Excelfile size=30></TD>
				</TR>
			</TABLE>
		</FIELDSET >
		</TD>
	</TR>
	<TR>
		<TD  <%=HEIGHT_TYPE_00%>></TD>
	</TR>	
	<TR>
		<TD align=center colspan=2><INPUT id=btnOK type=button value="확인" name=btnOK>
		<INPUT id=btnCancel type=button value="취소" name=btnCancel></TD>
	</TR>
	<TR>
		<TD><script language =javascript src='./js/importexcel_vspdData_vspdData.js'></script></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" SRC="about:blank" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
	
</TABLE>
</form>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></iframe>
</DIV>
</BODY></HTML>
