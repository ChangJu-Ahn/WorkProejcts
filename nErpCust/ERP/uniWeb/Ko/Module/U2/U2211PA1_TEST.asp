<% Option Explicit %>
<!--
'**********************************************************************************************
'*  1. Module Name          : SALE
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit																	'��: indicates that All variables must be declared in advance
'project_code = dlvy_no,	strReport_No = strDocument_no
Dim arrParent
Dim popupParent
Dim strMode
Dim strTable, strStatus, project_code, strSQL
Dim strBpCd, strReport_No, strins_person, strReport_Text
Dim arrtemp

arrParent   = window.dialogArguments
Set popupParent = arrParent(0)


</SCRIPT>

<%

Dim project_code,strMode,lngRow

    '---------------------------------------Common-----------------------------------------------------------
    Call LoadBasisGlobalInf()  
    Call HideStatusWnd									'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
  
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim strSystemFolder
	Dim iTempPath
	Dim strBpCd, strReport_No,strins_person,strReport_Text
	Dim strTitle
	Dim arrTemp
	Dim strReport_Abbr
	Dim struse_dt
	Dim strIns_dt
	Dim strReport_Nobar
	Dim strSQL
	
	strBpCd		 = CStr(UCASE(Request("bp_cd")))
	project_code = CStr(UCASE(Request("dlvy_no")))
	strReport_No = CStr(UCASE(Request("Document_no")))
'RESPONSE.WRITE 	strBpCd
'RESPONSE.WRITE 	project_code
'RESPONSE.WRITE 	strReport_No
	If IsNull(strReport_No) Or strReport_No = "" Then
'		strReport_No = "D0000003" '>>AIR
		
		StrSQL = " "
		StrSQL = StrSQL & " SELECT	ISNULL(MAX(DOCUMENT_NO),'D0000001') "
		StrSQL = StrSQL & " FROM	M_SCM_DOCUMENT_HDR_KO441(NOLOCK) "
		StrSQL = StrSQL & " WHERE	BP_CD = '" & strBpCd & "' "
		StrSQL = StrSQL & " 	AND	DLVY_NO = '" & project_code & "' "

'	'Call ServerMesgBox(StrSQL , vbInformation, I_MKSCRIPT)	
		Call SubOpenDB(lgObjConn)
'
'		         
			If 	FncOpenRs("R",lgObjConn,lgObjRs,StrSQL,"X","X") = False Then                    'If data not exists	
				strReport_No = "D0000001"
			Else
				strReport_No = lgObjRs(0)
		    End If
		    	
		Call SubCloseDB(lgObjConn)  		    	
		
	End If
	
'RESPONSE.WRITE StrSQL	
'RESPONSE.WRITE tempDocumentNo

	strMode  = CStr(Request("strMode"))	

	strSystemFolder = GetSpecialFolder(0) '0->WindowsFolder, 1->SystemFolder, 2->TemporaryFolder		
	strSystemFolder = strSystemFolder & "\TEMP"
	
	If right(strSystemFolder,1) <> "\" Then
		iTempPath = strSystemFolder & "\UNIERPTEMP\"
	Else
		iTempPath = strSystemFolder & "UNIERPTEMP\"
	End If

	'TEMP���� ������ ���� 
	Call CreateFolder(iTempPath)
	Response.Cookies("unierp")("gTempDirForFileUpload") = Replace(iTempPath, "\", "/")
	
	Select Case strMode
		Case CStr(UID_M0001) 
			strReport_Nobar="���"   
	    Case CStr(UID_M0002) 
			Call  SubBizQueryMulti()	
			strReport_Nobar="����"                                             
	End Select

'-----------------------------------------------------------------------------------------
Sub SubBizQueryMulti()
'-----------------------------------------------------------------------------------------
    Dim strData
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    Dim arr,arrCnt
    Dim i,j,StrSQL,kk,adoRec
    'BlankchkFlg = False
    
    On Error Resume Next  
	Err.Clear                                                   '��: Clear Error status
 	
	'project_code="2"
	LngRow = 0
	'StrSQL = "SELECT ins_user,report_no,report_abbr,report_text,ins_dt"
	'StrSQL = StrSQL & " FROM S_PRJ_REPORT_HDR_KO412 (NoLock) WHERE project_code = '" & project_code & "' and report_no = '" & strReport_No & "'"
	StrSQL = " "
	'StrSQL = StrSQL & " SELECT	INS_USER, DOCUMENT_NO, DOCUMENT_ABBR, DOCUMENT_TEXT, INS_DT "
	StrSQL = StrSQL & " SELECT	INS_USER, title, DOCUMENT_ABBR, DOCUMENT_TEXT, INS_DT "
	StrSQL = StrSQL & " FROM	M_SCM_DOCUMENT_HDR_KO441 "
	StrSQL = StrSQL & " WHERE	BP_CD = '" & strBpCd & "' "
	StrSQL = StrSQL & " 	AND	DLVY_NO = '" & project_code & "' "
	StrSQL = StrSQL & " 	AND	DOCUMENT_NO = '" & strReport_No & "' "
'Call ServerMesgBox(StrSQL , vbInformation, I_MKSCRIPT)	
	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)    
	         
		If 	FncOpenRs("R",lgObjConn,lgObjRs,StrSQL,"X","X") = False Then                    'If data not exists	

				'Call DisplayMsgBox("210100", vbInformation, "", "", I_MKSCRIPT)      '�� : ����� ���� ������ �ش��ϴ� �ڷᰡ �������� �ʽ��ϴ�.
				'Call SubCloseDB(lgObjConn)
				'Response.End   
		Else
				arrCnt = lgObjRs.RecordCount 
				arr=lgObjRs.GetRows
				Call SubCloseDB(lgObjConn)  
		
				strins_person	= ConvSPChars(arr(0,0))
				'strReport_No	= ConvSPChars(arr(1,0))
				strTitle		= ConvSPChars(arr(1,0))
				strReport_Abbr	= ConvSPChars(arr(2,0))
				strReport_Text	= arr(3,0)	'arr(3,0)
				strIns_dt		= UniConvDate(arr(4,0))
		
	    End If

End Sub    

Function GetSpecialFolder(iDx)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   GetSpecialFolder = pfile.GetSpecialFolder(CInt(iDx))   
   Set pfile = Nothing
End Function
		
Function CreateFolder(iTempPath)
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")
   Call pfile.CreateFolder(iTempPath)   
   Set pfile = Nothing
End Function


%>
<html>
<head>
<title><%=strReport_NoBar%></title>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->


<Script Language="VBScript">
'Option Explicit

Const BIZ_PGM_ID = "U2211PB1_TEST.asp"

Dim arrFileinf
Dim pStrFileInfo
Dim strMode
Dim arrTemp
Dim project_code
strMode  = "<%=strMode %>"
arrTemp  = "<%=arrTemp %>"
project_code = "<%=project_code %>"

<%'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
%>
<%
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
%>

	
<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
%>
function CheckValid()
	'-----------------------
    'Check condition area
    '-----------------------
  
   Call LayerShowHide(1)

'   if UNIConvDateCompanyToDB(frm1.use_dt.text,"") < UNIConvDateCompanyToDB( frm1.insrt_dt.value,"") then
'		Call DisplayMsgBox("971012", "X", "��ȿ��", "X")
'		frm1.use_dt.focus()
'		exit function
'  end if
   
    If Not chkField(Document, "A") Then         
       Exit Function
    End If

    call dbSave()
End function

Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 

Private Sub Form_Load()

    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, popupParent.gDateFormat, popupParent.gComNum1000, popupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N") 
	call SetDefaultVal()

	If CStr(strMode) = "<%=CStr(UID_M0002)%>" Then	'�� ����� �ƴϸ�, �� �ۼ����̸�,
		Call FncQueryFileInfo()
		'frm1.use_dt.text  = UniConvDateAToB("<%=strUse_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=strIns_dt%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	else
	
		'frm1.use_dt.text  = UniConvDateAToB("<%=dateadd("m",3,GetSvrDate)%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
		frm1.insrt_dt.value =UniConvDateAToB("<%=GetSvrDate%>", popupParent.gServerDateFormat, popupParent.gDateFormat)
	End If

	'Call InitComboBox

	frm1.txtTitle.value="<%=strTitle%>"
	frm1.report_abbr.value="<%=strReport_Abbr%>"
	frm1.ins_person.value="<%=strins_person%>"
	Call ggoOper.SetReqAttr(frm1.insrt_dt, "Q")
	
End Sub

'==========================================  2.2.6 InitComboBox()  ========================================
' Name : InitComboBox()
' Desc : Combo Display
'==========================================================================================================
'Sub InitComboBox()
'
'	'// ǰ������
'	Call CommonQueryRs(" UD_MINOR_CD,UD_MINOR_NM "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD = " & FilterVar("SX006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'	Call SetCombo2(frm1.txtTitle , lgF0, lgF1, Chr(11))
'	
'End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.ins_person.value=popupParent.gUsrID
End Sub

Function FncQueryFileInfo()

	''On Error Resume Next
	Dim IntRetCD1
	Dim iLngRow
	Dim strData
	Dim strRet
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim ArrTmpF0,ArrTmpF1,ArrTmpF2,ArrTmpF3,ArrTmpF4,ArrTmpF5,ArrTmpF6
	
	IntRetCD1= CommonQueryRs("DOCUMENT_NM,DOCUMENT_ID, DOCUMENT_SIZE ","M_SCM_DOCUMENT_DTL_KO441", "BP_CD = '<%=strBpCd%>' and DLVY_NO = '" & project_code & "' and DOCUMENT_NO = '<%=strReport_No%>'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	ArrTmpF0 = split(lgF0,popupParent.gColSep)	
	ArrTmpF1 = split(lgF1,popupParent.gColSep)	
	ArrTmpF2 = split(lgF2,popupParent.gColSep)	

	strData = ""
	strRet = ""
	For iLngRow = 0 To UBound(ArrTmpF0, 1) - 1
		strData = "1" & ArrTmpF0(iLngRow) & "" & ArrTmpF2(iLngRow) & "" & ArrTmpF1(iLngRow) & "101344919970601092656N0NNFFY00YI"		
		strRet = strRet & mid(strData,3,len(strData))
		Call SetAttachFile(strData)
	Next
	Call MakeFileInfoArray("1" & strRet)		
End Function


'==========================================  3.1.2 Window_OnUnLoad() ======================================
'	Name : Window_OnUnLoad()
'	Description : Window �� �ݱ��ư(�ּ�,�ִ�ȭ��ư ���� �ִ� �ݱ��ư)�� ������ �� ����Ǵ� �κ� 
'========================================================================================================= 
Private Sub Window_OnUnLoad()
	If  window.ReturnValue <> True then
		window.ReturnValue = False
	End If
End Sub
	

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************
Function DbSave()

	If UCASE(frm1.txtblnFileAttached.value) = "TRUE" Then
		Call MakeFileInfoString(arrFileinf)
	End If
	Call LayerShowHide(1) 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	

End Function

Function DbSaveOk()	

	//parent.frm1.vspdData.MaxRows = 0
	window.ReturnValue = True
	Self.Close()
End Function


Function vbAttachFile()

	Dim strRet

	strRet = AttachFile()

	if len(strRet) > 3  then

		Call MakeFileInfoArray(strRet)

	else
	exit function
		
	end if
	

End Function

Function MakeFileInfoArray(strRet)
	Dim arrTemp,arrTemp2
	Dim i,j
	Dim iFileCount
	Dim iCurrentSize
	
	arrTemp = Split(strRet,Chr(31))
	
	iFileCount = UBound(arrTemp) - 1	'iFileCount	÷�εǴ� ���� �Ѱ���	
	
	If iFileCount = 0 Then
		Exit Function
	End If
	frm1.txtblnFileAttached.value = "TRUE"

	If IsArray(arrFileinf) Then
		iCurrentSize = UBound(arrFileinf,2)
		Redim Preserve arrFileinf(21, Cint(iCurrentSize + iFileCount))			
	Else
		iCurrentSize = 0
		Redim arrFileinf(21, iCurrentSize + iFileCount)
	End If
	For i = 1 To iFileCount				'arrTemp(0)�� ÷�εǴ� ���� �Ѱ��� + 1
		
		arrTemp2 = Split(arrTemp(i),Chr(29))
		
		For j = 1 To UBound(arrTemp2) + 1		
			arrFileinf(j, iCurrentSize + i) = arrTemp2(j - 1)
		Next
	Next	
	
	Call ShowArrayString(arrFileinf)

End Function

Function MakeFileInfoString(prArrTemp)
	Dim i

	Const C_FileName = 2
	Const C_FileId = 4	
	Const C_FileSize = 9
	Const C_FileCDate = 10
	
	For i = 1 To UBound(prArrTemp,2)	
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileName, i)  & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileId, i)    & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileSize, i)  & popupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileCDate, i) & popupParent.gRowSep
	Next
	frm1.txtFileinf.value = pStrFileInfo
MsgBox frm1.txtFileinf.value	
End Function

dim deltemp
deltemp=""

Function vbDeleteFile
	Dim iDx
	Dim i
	Dim j
	Dim temp
	
	If frm1.filelist.length > 0 Then
	
		If frm1.filelist.selectedIndex < 0 Or IsNull(frm1.filelist.selectedIndex) Then 
			MsgBox "������ ������ �����Ͻʽÿ�."
			Exit Function
		End If
		
		For iDx = CInt(frm1.filelist.length) - 1 To 0 Step -1
			If frm1.filelist.options(iDx).Selected Then
				
				temp=frm1.filelist.options(iDx).value
				if len(deltemp)>1 then
					deltemp=deltemp &","& split(temp,"")(3)
				else 
					deltemp= split(temp,"")(3)
				end if
				
				
				frm1.delTemp.value =	deltemp
	
				frm1.filelist.remove(iDx)
				RearrangeArray(iDx + 1)	
			End If
		Next	

	Else		
		Exit Function
	End If
End Function

Sub RearrangeArray(iDx)
	' �迭�� �������ϰ� �ǵ��� �迭�� dispose�Ѵ�.
    Dim i,j

    For i = iDx To UBound(arrFileinf,2) - 1
        For j = 1 To UBound(arrFileinf,1)	'21�� ���� 
			arrFileinf(j, i) = arrFileinf(j, i + 1)
        Next        
    Next

	Redim Preserve arrFileinf(UBound(arrFileinf,1), UBound(arrFileinf,2) - 1)

End Sub

Function ShowArrayString(prArrTemp)
	Dim i,j,iStrTemp
	
	For i = 1 To UBound(prArrTemp,2)	
		For j = 1 To UBound(prArrTemp,1)
			iStrTemp =  iStrTemp & "(" & j & "," & i & ") = " & prArrTemp(j,i) & vbCrlf
		Next
	Next
	
	
End Function

Function FetchWebSvrIp()	

	Dim gHttpWebSvrIPURL
	
	gHttpWebSvrIPURL =  "http://<%= request.servervariables("server_name") %>"	
	FetchWebSvrIp = Split(gHttpWebSvrIPURL, "/")(2)
End Function


</Script>

<script language=javascript>
function AttachFile(){
	var sRet;
	var optStr;
	var newOpt;
	var temp;
	var temp2;
	var strWebSvrIp;
		
	document.FR_ATTWIZ.SetLanguage('');
	document.FR_ATTWIZ.SetServerOption(0,0);
	document.FR_ATTWIZ.SetUploadMode(1);
	document.FR_ATTWIZ.SetFileUIMode(1);
	document.FR_ATTWIZ.SetModUpload();
	document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	document.FR_ATTWIZ.SetTempDir("<%=Request.Cookies("unierp")("gTempDirForFileUpload")%>");
	
	sRet = document.FR_ATTWIZ.AttachFile('NEW');

	var arrTemp = sRet.split(String.fromCharCode(31));

	for(var i = 0; i < arrTemp.length; i++){

		var iFileinf = arrTemp[i].split(String.fromCharCode(29));

		if(iFileinf.length > 2){
			optStr = new Option(iFileinf[1],sRet,true);	
			document.frm1.filelist.options[document.frm1.filelist.length] = optStr;
		}
	}
	//document.frm1.FileInfo.value  = sRet; 
	return(sRet);
}

function SetAttachFile(sRet){
	var arrTemp = sRet.split(String.fromCharCode(31));
	
	for(var i = 0; i < arrTemp.length; i++){		
		var iFileinf = arrTemp[i].split(String.fromCharCode(29));
		
		if(iFileinf.length > 2){
			optStr = new Option(iFileinf[1],sRet,true);				
			document.frm1.filelist.options[document.frm1.filelist.length] = optStr;
		}
	}		
}
</script>

<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</head>
<BODY TABINDEX="-1" SCROLL="no">


<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<input type=hidden name=txtMode value="<%=strMode%>">
<input type=hidden name=txtblnFileAttached value="">
<input type=hidden name=txtFileinf value="">
<!--<input type=hidden name=project_code value="<%=project_code%>">-->	
<input type=hidden name=BpCd value="<%=strBpCd%>">
<input type=hidden name=DlvyNo value="<%=project_code%>">
<input type=hidden name=DocumentNo value="<%=strReport_No%>">
<input type=hidden name=delTemp value="">
<TABLE  CELLSPACING=0 CELLPADDING=0 >
	
	
	<TR HEIGHT=100%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD  HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
					<TABLE  CLASS="BasicTB" CELLSPACING=0>						
						<TR>
							<TD CLASS="TD5" NOWRAP>����</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="����" name="txtTitle" tag="22"  size=70 MAXLENGTH=100  >
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>��༳��</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
							<INPUT TYPE=TEXT ALT="��༳��" name="report_abbr" tag="21"  size=70 MAXLENGTH=100  ></TD>
						</TR>
				
						<TR>
							<TD CLASS="TD5" NOWRAP>�ۼ���</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><input name="ins_person" alt="�ۼ���" tag="22" MAXLENGTH=15   ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>�����</TD>
							<TD CLASS="TD6" NOWRAP><input type=text name="insrt_dt" size=12>
							</TD>
						</TR>
						
						<TR>
							<TD CLASS="TD5" NOWRAP>����</TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3><TEXTAREA class="tb4" alt="����" cols=64 name="txtReportText" tag="22" rows=13 MAXLENGTH=1000  wrap=phsical><%= strReport_Text%></TEXTAREA></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>
								<INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbScript:vbAttachFile()' type=button value="÷��" id=button1 name=button1><BR>
								<INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick='vbScript:vbDeleteFile()' type=button value="��Ͽ��� ����" id=button2 name=button2></TD>
							<TD CLASS="TD6" NOWRAP COLSPAN=3>
								<SELECT  style="WIDTH: 470px" tag="21"  size=5 name=filelist multiple></SELECT></TD>
						</TR>
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD  HEIGHT=3 ></TD>
	</TR>
		
	<TR>
		<TD WIDTH=* ALIGN=RIGHT><input type=button name="OK" onclick="vbscript:CheckValid()" value="Ȯ��">&nbsp;&nbsp<input type=button name="Cancel" onclick="vbscript:FncClose()" value="�ݱ�">&nbsp;</TD>
		<TD WIDTH=10>&nbsp;</TD>					
	</TR>	
	
</TABLE>
<table border=0 cellpadding="0" cellspacing="0" width="575" align=center>

</table>
</form>


<IFRAME  NAME="MyBizASP" STYLE="display: '';" WIDTH=100% HEIGHT=600 scrolling=yes></IFRAME>
<IFRAME NAME="FR_ATTWIZ" SRC="FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 ></IFRAME><BR>


</center>
</body>
</html>