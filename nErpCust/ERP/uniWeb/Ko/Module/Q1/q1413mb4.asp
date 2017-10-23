<%@LANGUAGE = VBScript%> 
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MB4
'*  4. Program Name         : 수입검사품질추이(일별)
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX용 상수를 사용하기 위한 Include 지정 -->
<!-- #include file="../../inc/CfxIE.inc" -->
<%													
On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "QB")

Dim strLotsize
Dim strAlpha
Dim strBeta
Dim strP1
Dim strP2

Dim AttSamlpesize
Dim AttAcceptQty

Dim Temp1
Dim Temp2
Dim Temp3
Dim Temp4

strLotsize = Request("txtLotSize")
strAlpha = Request("txtAlpha")
strBeta = Request("txtBeta")
strP1 = Request("txtP1")
strP2 = Request("txtP2")

'Comproxy를 연결하여 값을 받아온다.

Temp1 = 70
Temp2 = 0
Temp3 = 70
Temp4 = 2
%>
<Script Language=vbscript>
With Parent.frm1
	.txtSampleSize1.Text = "<%= Temp1 %>"
	.txtAcceptSize1.Text = "<%= Temp2 %>"
	.txtSampleSize2.Text = "<%= Temp3 %>"
	.txtAcceptSize2.Text = "<%= Temp4 %>"
End with
</Script>	
