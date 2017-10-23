<!--
'**********************************************************************************************
'*  1. Module Name          : �������� ���/���� ȭ�� ó�� ASP
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/06
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
<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->

<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/Cookie.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../inc/incImage.js"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
'Option Explicit																	'��: indicates that All variables must be declared in advance
Dim arrParent
Dim PopupParent

arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
</SCRIPT>

<%

    '---------------------------------------Common-----------------------------------------------------------
    Call LoadBasisGlobalInf()    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

'	Dim objSystemFolder 	
'	Set objSystemFolder = Server.CreateObject("PSystem.CFolder")	
'	strSystemFolder = objSystemFolder.GetTempDirectory()	
	
	Dim strSystemFolder
	Dim iTempPath
		
	strSystemFolder = GetSpecialFolder(0) '0->WindowsFolder, 1->SystemFolder, 2->TemporaryFolder		
	strSystemFolder = strSystemFolder & "\TEMP"
	
	
	If right(strSystemFolder,1) <> "\" Then
		iTempPath = strSystemFolder & "\UNIERPTEMP\"
	Else
		iTempPath = strSystemFolder & "UNIERPTEMP\"
	End If	

	'TEMP���� ������ ���� 
	Call CreateFolder(iTempPath)

	Response.Cookies("unierp")("gTempDirForFileUpload")         = Replace(iTempPath, "\", "/")
	'Response.Cookies("unierp")("gTempDirForFileUpload")         = Replace(Request.ServerVariables("APPL_PHYSICAL_PATH") & "Ko\Notice\TEMP\", "\", "/")


Dim strTitle , strMode
Dim strTable, strStatus, intKeyNo, strSQL
Dim strSubject, strWriter, strContents, strPasswd
Dim arrtemp

intKeyNo = CLng(Request("intKeyNo"))
strMode  = CStr(Request("strMode"))													'��: Read Operation Mode (CRUD)

Select Case strMode
    Case CStr(UID_M0001)                                                         '��: Query
		strTitle = "�������� ���"	        
			
    Case CStr(UID_M0002)                                                         '��: Save,Update
		''On Error Resume Next
		Err.Clear 
		strTitle = "�������� ����"
	
		Dim objRec, objConn
	
		Set objRec = Server.CreateObject("ADODB.RecordSet")
	
		Call SubOpenDB(objConn)
				
		strSQL = "SELECT * FROM B_NOTICE WHERE NOTICENUM = '" & intKeyNo & "'"		
		
'		Response.Write strSQL
'		Response.Write gADODBConnString
'		Response.End
		
		objRec.Open strSQL, gADODBConnString
				
		If Not objRec.EOF Then
			strSubject = "" & objRec("subject")
			strWriter = "" & objRec("writer")
			strContents = "" & objRec("contents")			
			objRec.Close
		End If
				
		Set objRec = Nothing	
		Call SubCloseDB(objConn)
			        
    Case CStr(UID_M0003)                                                         '��: Delete
        
End Select


Function DisplayMessageBox(temp)
	Response.Write "<Script Language=vbscript>"            & vbCr
	Response.Write " msgbox """       &  temp   & """" & vbCr
	Response.Write "</Script>"                             & vbCr
End Function

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

Function GetGlobalTheOthersInf(ByVal pData)

	On error resume next 
	
    Dim FileNm
	Dim xmlDoc
	Dim NodeNm 
	
	NodeNm ="Login"
	
	set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")		
	xmlDoc.async = false 
    
	FileNm = Request.Cookies("unierp")("gXMLFileNm")
	xmlDoc.Load FileNm    
		
    Select Case pData    
        Case "HttpWebSvrIPURL"        :  GetGlobalTheOthersInf	= xmlDoc.selectSingleNode("/uniERP/" & NodeNm & "/" & "HttpWebSvrIPURL").text       	        
	End Select   
	
	set xmlDoc = nothing

End Function

%>
<html>
<head>
<title><%=strTitle%></title>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">	
<style>

.body {
        text-decoration: none;; line-height: 10pt}

<!-- ���� ���� -->
.topmenu {
        color: #280A72;
        text-decoration: none;
        font-size: 10pt;
        line-height: normal;}
.topmenu a      {
        color: #0000FF;
        text-decoration: underline;
        font-size: 10pt;
        line-height: normal;}
.topmenu a:hover        {
        color: #FF0000;
        font-size: 10pt;
        line-height: normal;}

<!-- �׸𱤰� �޴� -->                 

.boxmenu {
        color: #E1F4FF;
        text-decoration: none;
        font-size: 10pt;}
.boxmenu a      {
        color: #E1F4FF;
        text-decoration: none;
        font-size: 10pt;}
.boxmenu a:hover        {
        font-size: 10pt;; text-decoration: underline}

.tk {font-size: 10pt; color: #000000; line-height: 16pt; }
.tk2 {font-size: 10pt; color: #000000; line-height: 16pt; }
</style>      

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->


<Script Language="VBScript">
'Option Explicit

Const BIZ_PGM_ID = "frwriteBiz.asp"

Dim arrFileinf
Dim pStrFileInfo
Dim strMode
Dim arrTemp
Dim intKeyNo

strMode  = "<%= strMode %>"
arrTemp  = "<%= arrTemp %>"
intKeyNo = "<%= intKeyNo %>"

<%'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
<%'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================%>

	
<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>
function CheckValid()

	If trim(frm1.subject.value) = "" then
		Msgbox "������ �Է��� �ֽʽÿ�!", vbInformation, "üũ"
		frm1.subject.focus 
	ElseIf trim(frm1.Writer.value) = "" then
		Msgbox "������ �Է��� �ֽʽÿ�!", vbInformation, "üũ"
		frm1.Writer.focus 
	Else
		If len(frm1.txtContent.value) > 50000 Then
			Msgbox "���� ������ 50000�ڸ� �ʰ��� �� �����ϴ�.", vbInformation, "üũ"
			frm1.txtContent.focus 	
			Exit Function		
		End If				
		Call DBSave
	End if

End function

Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function

<% '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################%>
<% '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* %>
<% '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= %>
Private Sub Form_Load()		
		' ��Ŀ�� ó�� 
		frm1.subject.focus
		If CStr(strMode) <> CStr(UID_M0001) Then	'�� ����� �ƴϸ�, �� �ۼ����̸�,		
			Call FncQueryFileInfo()
		End If
		
End Sub

Function FncQueryFileInfo()

	''On Error Resume Next
	Dim IntRetCD1
	Dim iLngRow
	Dim strData
	Dim strRet
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim ArrTmpF0,ArrTmpF1,ArrTmpF2,ArrTmpF3,ArrTmpF4,ArrTmpF5,ArrTmpF6

	IntRetCD1= CommonQueryRs("FLE_NAME, FLE_ID, FLE_PATH ","B_NOTICE_FILE", "NOTICENUM = " & intKeyNo,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	ArrTmpF0 = split(lgF0,PopupParent.gColSep)	
	ArrTmpF1 = split(lgF1,PopupParent.gColSep)	
	ArrTmpF2 = split(lgF2,PopupParent.gColSep)	

'	ArrTmpF0 = split(lgF0,Chr(11))	'For HERMES
'	ArrTmpF1 = split(lgF1,Chr(11))	'For HERMES
'	ArrTmpF2 = split(lgF2,Chr(11))	'For HERMES
			
	strData = ""
	strRet = ""

	For iLngRow = 0 To UBound(ArrTmpF0, 1) - 1
		strData = "1" & ArrTmpF0(iLngRow) & "" & ArrTmpF2(iLngRow) & "" & ArrTmpF1(iLngRow) & "101344919970601092656N0NNFFY00YI"		
		strRet = strRet & mid(strData,3,len(strData))
		Call SetAttachFile(strData)
	Next

	Call MakeFileInfoArray("1" & strRet)		
End Function

'Function SetAttachFile(sRet)
'	Dim iArrTemp, iFileinf
'	Dim i
'	iArrTemp = Split(sRet,Chr(31))
'	
'	For i = 0 To UBound(iArrTemp)
'		iFileinf = Split(iArrTemp(i),Chr(29)
'		If UBound(iFileinf) > 2) Then
'			frm1.filelist.options(UBound(iArrTemp))
'		End If
'	Next
'
'End Function

'==========================================  3.1.2 Window_OnUnLoad() ======================================
'	Name : Window_OnUnLoad()
'	Description : Window �� �ݱ��ư(�ּ�,�ִ�ȭ��ư ���� �ִ� �ݱ��ư)�� ������ �� ����Ǵ� �κ� 
'========================================================================================================= 
Private Sub Window_OnUnLoad()
	If  window.ReturnValue <> True then
		window.ReturnValue = False
	End If
End Sub
	
'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************
Function DbSave()
	If UCASE(frm1.txtblnFileAttached.value) = "TRUE" Then
		Call MakeFileInfoString(arrFileinf)
	End If
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
End Function

Function DbSaveOk()	
	window.ReturnValue = True
	Self.Close()
End Function


Function vbAttachFile()
	Dim strRet
		
	strRet = AttachFile()
	Call MakeFileInfoArray(strRet)

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
	
	'Call ShowArrayString(arrFileinf)

End Function

Function MakeFileInfoString(prArrTemp)
	Dim i

	Const C_FileName = 2
	Const C_FileId = 4	
	Const C_FileSize = 9
	Const C_FileCDate = 10
	
	For i = 1 To UBound(prArrTemp,2)	
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileName, i)  & PopupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileId, i)    & PopupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileSize, i)  & PopupParent.gColSep
			pStrFileInfo = pStrFileInfo & prArrTemp(C_FileCDate, i) & PopupParent.gRowSep
	Next

	frm1.txtFileinf.value = pStrFileInfo		
	
End Function

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
	
	Msgbox iStrTemp
	
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
	//document.FR_ATTWIZ.SetExtension('testgultxt');
	
	//������ip Fetch
	//temp = "<%=Request.Cookies("unierp")("gHttpWebSvrIPURL")%>";	
	//strWebSvrIp = temp.split(String.fromCharCode(47))[2];	
	
	document.FR_ATTWIZ.SetServerInfo(FetchWebSvrIp(), '7775');
	
	//document.FR_ATTWIZ.SetTempDir('F:/Program Files/uniERP II/uniWeb/Ko/Notice/TEMP/');	
	document.FR_ATTWIZ.SetTempDir("<%=Request.Cookies("unierp")("gTempDirForFileUpload")%>");	
	//document.FR_ATTWIZ.SetAttachedFileInfo(document.dmtest.FileInfo.value);
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

<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</head>

<BODY BGCOLOR="#FFFFFF" scroll=no leftmargin=2 rightmargin=0 topmargin=0 bottommargin=0>
<center>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<input type=hidden name=txtMode value="<%=strMode%>"><input type=hidden name=txtKeyNo value="<%=intKeyNo%>">
<input type=hidden name=txtblnFileAttached value="">
<input type=hidden name=txtFileinf value="">
<table border=0 cellpadding="0" cellspacing="0" width="575" align=center>
	<TR>
		<TD HEIGHT=1>&nbsp;<% ' ���� ���� %></TD>
	</TR>
      <tr align="middle">
        <td CLASS="Tab11">                                          
         <table border="0" cellpadding="2" cellspacing="1" class="topmenu" width=100%>
            <tr>
              <td width=100 bgcolor="#eeeeee" align="middle" class="tk"><font color="#666666"><b>�� ��</b></font></td>
              <td width=475 bgcolor="#f7f7f7" class="tk"><input name="subject" tag="2" size=50 MAXLENGTH=50 value="<%=strSubject%>" title="������ �ѱ� 50�� ���� �����մϴ�" STYLE="BACKGROUND-COLOR: white; BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid"></td>
            </tr>                                                                   

            <tr>
              <td width=100 bgcolor="#eeeeee" align="middle" class="tk"><font color="#666666" ><b>�� ��</b></font></TD>
              <td width=475 bgcolor="#f7f7f7" class="tk"><input name="Writer" tag="2" MAXLENGTH=25 <% If Cint(strMode) = UID_M0001 Then %> value="<%=GetGlobalInf("gUsrNm")%>" <%ElseIf Cint(strMode) = UID_M0002 Then%> value="<%=strWriter%>"<% End If %> title="[������ 8��(�ѱ� 4��)�Դϴ�]" STYLE ="BACKGROUND-COLOR: white; BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid"></td>
            </tr>                                                                   

            <tr>
              <td width=100 bgcolor="#eeeeee" align="middle" class="tk"><font color="#666666" ><b>�� ��</b></font></td>
              <td width=475 bgcolor="#f7f7f7" class="tk"><TEXTAREA cols=64 name=txtContent tag="2" rows=20 MAXLENGTH=1000 ><%= strContents %></TEXTAREA></td>
            </tr>                                                                   
            <tr>
				<td width=475 bgcolor="#f7f7f7" class="tk">
					<table width="100%" border="1">
					  <tr>
					    <td width=475 bgcolor="#f7f7f7" align="middle" class="tk"><INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick = 'vbScript:vbAttachFile()' type=button value=÷�� id=button1 name=button1></td>
					  </tr>
					  <tr>
					    <td width=475 bgcolor="#f7f7f7" align="middle" class="tk"><INPUT style="FONT-SIZE: 9pt; WIDTH: 90px; COLOR: #000000; PADDING-TOP: 2px; HEIGHT: 22px; BACKGROUND-COLOR: #d4d0c8" onclick='vbScript:vbDeleteFile()' type=button value="��Ͽ��� ����" id=button2 name=button2></td>
					  </tr>
					</table>				
				</td>
				<td width=475 bgcolor="#f7f7f7" class="tk"><SELECT style="WIDTH: 470px" size=5 name=filelist multiple></SELECT></td>
            </tr>
            <!--tr>
              <td width=475 bgcolor="#f7f7f7" class="tk"><INPUT type=hidden id=text1 name=text1>&nbsp;</td>
              <td width=475 bgcolor="#f7f7f7" class="tk"><INPUT type=hidden NAME = FileInfo SIZE = 50></td>
            </tr-->                                                                   
            
		  </table>
		</td>
     </tr>
	<TR>
		<TD HEIGHT=1>&nbsp;<% ' ���� ���� %></TD>
	</TR>
	<TR>
		<TD WIDTH=* ALIGN=RIGHT><input type=button name="OK" onclick="vbscript:CheckValid()" value="Ȯ��">&nbsp;&nbsp<input type=button name="Cancel" onclick="vbscript:FncClose()" value="�ݱ�">&nbsp;</TD>
		<TD WIDTH=10>&nbsp;</TD>					
	</TR>	     
</table>
</form>


<IFRAME NAME="MyBizASP" STYLE="display: '';" WIDTH=100% HEIGHT=1000></IFRAME>
<IFRAME NAME="FR_ATTWIZ" SRC="FrAttwiz.html" MARGINWIDTH=0 MARGINHEIGHT=0 WIDTH=0 HEIGHT=0 FRAMEBORDER=0></IFRAME><BR>
</center>
</body>
</html>