<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim arrParent
Dim PopupParent
'Dim IsAttach

'IsAttach = False
arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
</SCRIPT>

<%
	Response.Expires = -1                               '☜: will expire the response immediately
	Response.Buffer = True                              '☜: The server does not send output to the client until all of the ASP 
														'    scripts on the current page have been processed
	Dim lgObjConn, lgObjRs
	Dim lgOpModeCRUD, iStrData
	Dim lgStrSQL

	Call LoadBasisGlobalInf()
	'---------------------------------------Common-----------------------------------------------------------
%>
<SCRIPT LANGUAGE="JavaScript">
var nJobTotal = 0;
var nRecTotal = new Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);

var nJobDone = 0;
var nJobCur = 0;
var nRecDone = 0;
var nRecCur = 0;
var nOffset = 0;

var nPvarCur1 = 0;
var nPvarGol1 = 0;
var nPvarCur2 = 0;
var nPvarGol2 = 0;

function init() {
	initBar1();
	
	document.images[20].src = "button1.gif";
	document.images[21].src = "button1.gif";
	document.images[22].src = "button1.gif";
	document.images[23].src = "button1.gif";
	document.images[24].src = "button1.gif";
	document.images[25].src = "button1.gif";
	document.images[26].src = "button1.gif";
	document.images[27].src = "button1.gif";
	document.images[28].src = "button1.gif";
	document.images[29].src = "button1.gif";
	document.images[30].src = "button1.gif";
	document.images[31].src = "button1.gif";
	document.images[32].src = "button1.gif";
	document.images[33].src = "button1.gif";
	document.images[34].src = "button1.gif";
	document.images[35].src = "button1.gif";
	document.images[36].src = "button1.gif";
	document.images[37].src = "button1.gif";
	document.images[38].src = "button1.gif";
	document.images[39].src = "button1.gif";
	
	GetData();
}

function initBar1() {
	document.images[0].src = "button1.gif";
	document.images[1].src = "button1.gif";
	document.images[2].src = "button1.gif";
	document.images[3].src = "button1.gif";
	document.images[4].src = "button1.gif";
	document.images[5].src = "button1.gif";
	document.images[6].src = "button1.gif";
	document.images[7].src = "button1.gif";
	document.images[8].src = "button1.gif";
	document.images[9].src = "button1.gif";
	document.images[10].src = "button1.gif";
	document.images[11].src = "button1.gif";
	document.images[12].src = "button1.gif";
	document.images[13].src = "button1.gif";
	document.images[14].src = "button1.gif";
	document.images[15].src = "button1.gif";
	document.images[16].src = "button1.gif";
	document.images[17].src = "button1.gif";
	document.images[18].src = "button1.gif";
	document.images[19].src = "button1.gif";
}

function PBar() {
	nPvarGol1 = Math.round(nRecDone / nRecTotal[nJobCur] * 19);
	nPvarGol2 = Math.round(nJobDone / nJobTotal * 19);
	nRecCur = Math.round(nPvarCur1 * nRecTotal[nJobCur] / 19);

	if (nPvarCur1 <= nPvarGol1) {
		document.images[nPvarCur1].src = "button2.gif";
		document.Frm1.PERCENT.value = nRecCur + " / " + nRecTotal[nJobCur];
		nPvarCur1++;
		setTimeout("PBar();", 100);
		return;
	} else {
		if(nPvarCur1 >= 19)
		{
			Completed();
		}else{
			setTimeout("GetData();", 2000);
			return;
		}
	}
}

function Completed() {
	var i = 0;
	nJobCur++;
	nOffset = 19 / nJobTotal;

	if (nPvarCur2 <= nPvarGol2) {
		for(i = 0; i <= nOffset; i++)
		{	
			if(nPvarCur2 <= 19)
				document.images[nPvarCur2+20].src = "button2.gif";
			nPvarCur2 ++ ;
		}
		
		document.Frm1.PERCENT1.value = nJobCur + " / " + nJobTotal;

		if(nPvarCur2 < 19)
		{
			if(nPvarCur2 <= nPvarGol2)
			{
				initBar1();
				nPvarCur1 = 0;
				setTimeout("PBar();", 100);
			}else{
				if(nPvarGol2 < 19)
				{
					setTimeout("GetData();", 2000);
				}
			}
		}
	}
}

function GetData()
{
	document.Frm2.submit();
}

function QueryOk()
{
	document.Frm1.PERCENT1.value = nJobCur + " / " + nJobTotal;

	PBar();
}
</SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'==========================================================================================================
Private Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call ggoOper.LockField(Document, "N")
End Sub

Function FncClose()
	window.ReturnValue = False
	Self.Close
End Function
</SCRIPT>

</HEAD>
<BODY SCROLL=NO TABINDEX="-1" onLoad="init()">
<FORM NAME="Frm1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD">
			<TABLE WIDTH=100% CELLSPACING=0>	
				<TR>
					<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
					<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>작업목록</td>
					<TD CLASS=TD6 NOWRAP>작업처리상황</td>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
					<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
				</TR>
				<TR>
					<TD HEIGHT=3 NOWRAP></TD>
					<TD HEIGHT=3 NOWRAP></TD>
				</TR>
				<TR>
					<TD ROWSPAN=2 CLASS=TD5 NOWRAP>
						<TABLE>
<%
				Dim strNum
				
				Call SubOpenDB(lgObjConn)                           '☜: Make a DB Connection
				
				lgStrSQL = "SELECT job_id, job_title " & _
						   "FROM   b_bdc_jobs " & _
						   "WHERE  job_id IN ('" & Replace(Request("txtJobs"), " ", "','") & "')"

				If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then           'If data not exists
					Response.Write "<TR><TD>&nbsp;</TD></TR>" & vbCrLf
				Else
					strNum = 1
					Do while Not (lgObjRs.EOF Or lgObjRs.BOF)
			            Response.Write "<TR><TD>" & CStr(strNum) & ". " & lgObjRs("job_title") & "</TD></TR>" & vbCrLf
			            strNum = strNum + 1
						lgObjRs.MoveNext
					Loop
				End If

				Call SubCloseRs(lgObjRs)                                        '☜ : Release RecordSSet
			    Call SubCloseDB(lgObjConn)                          '☜: Close DB Connection
%>
			 			</TABLE>
					</TD>
					<TD CLASS=TD6 NOWRAP>
						<IMG SRC="button1.gif" name="1" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="2" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="3" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="4" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="5" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="6" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="7" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="8" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="9" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="10" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="11" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="12" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="13" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="14" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="15" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="16" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="17" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="18" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="19" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="20" HEIGHT=10 WIDTH=10 HSPACE=0>
						&nbsp;
						<INPUT TYPE="TEXT" SIZE="8" CLASS=protected READONLY=true TABINDEX="-1" NAME="PERCENT" tag="14">&nbsp;건 
					</TD>
				</TR>
				<TR>
					<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
				</TR>
				<TR>
					<TD HEIGHT=3 NOWRAP></TD>
					<TD HEIGHT=3 NOWRAP></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>전체</TD>
					<TD CLASS=TD6 NOWRAP>
						<IMG SRC="button1.gif" name="21" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="22" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="23" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="24" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="25" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="26" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="27" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="28" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="29" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="30" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="31" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="32" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="33" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="34" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="35" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="36" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="37" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="38" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="39" HEIGHT=10 WIDTH=10 HSPACE=0>
						<IMG SRC="button1.gif" name="40" HEIGHT=10 WIDTH=10 HSPACE=0>
						&nbsp;
						<INPUT TYPE="TEXT" SIZE="8" CLASS=protected READONLY=true TABINDEX="-1" NAME="PERCENT1" tag="14">&nbsp;작업 
					</TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
					<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
				</TR>
			</TABLE>	
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=1>&nbsp;</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=* NOWRAP>&nbsp;&nbsp;</TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="vbscript:FncClose()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</FORM>
<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no NORESIZE FRAMESPACING=0></IFRAME>
<FORM  NAME="Frm2" Target="MyBizASP" Method=post Action="./BDC04PB2.asp">
	<input type=hidden name=txtJobs value="<%=Request("txtJobs")%>" >
</FORM>
</HTML>
