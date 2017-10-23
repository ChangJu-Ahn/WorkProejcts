<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : B1B06MB1.asp
'*  4. Program Name         : List B_ITEM_ACCT
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2004/12/01
'*  7. Modified date(Last)  : 2005/07/18
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next

Dim strQryMode

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim oPB3S115
Dim strItemAcct
Dim iReturn 

Const B491_EG1_item_acct = 0
Const B491_EG1_rep_item_acct = 1

Err.Clear

	' Display b_item_acct
	strItemAcct = Trim(Request("txtItemAcct"))
	
	Set oPB3S115 =  Server.CreateObject("PB3S115.cBSetItemAcct")
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If		
	
	iReturn = oPB3S115.B_QUERY_ITEM_ACCT_INF(gStrGlobalCollection , strItemAcct)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set oPB3S115 = Nothing
		Response.End
	End If
	
	If Not (oPB3S115 Is Nothing) Then 
		Set oPB3S115 = Nothing
	End If	
	
	If IsEmpty(iReturn)  Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End													
	End If
	
%>

<Script Language=vbscript>

Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		
<%  
	If Not (IsEmpty(iReturn)) Then
	
%>			
		ReDim TmpBuffer(<%=Ubound(iReturn ,1)%>)
<%
	End If
	
	For i=0 to Ubound(iReturn ,1)
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(iReturn(i,B491_EG1_item_acct))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(iReturn(i,B491_EG1_rep_item_acct))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
	Next
%>
	iTotalStr = Join(TmpBuffer,"")
	.ggoSpread.Source = .frm1.vspdData
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>