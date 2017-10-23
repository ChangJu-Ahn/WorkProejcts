<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : B1B07MB1.asp
'*  4. Program Name         : List B_ITEM_ACCT_TRACKING
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2006/06/27
'*  7. Modified date(Last)  : 2006/06/27
'*  8. Modifier (First)     : LEE SEUNG WOOK
'*  9. Modifier (Last)      : LEE SEUNG WOOK
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'=======================================================================================================
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim strQryMode

Dim StrNextKey
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim oPB3S116
Dim strPlantCd
Dim strItemAcct
Dim iReturn 

Const B494_EG1_plant_cd = 0
Const B494_EG1_item_acct = 1
Const B494_EG1_tracking_flag = 2

Err.Clear
	
	strPlantCd = Trim(Request("txtPlantCd"))
	strItemAcct = Trim(Request("txtItemAcct"))
	
	Set oPB3S116 =  Server.CreateObject("PB3S116.cBSetItemAcctTracking")
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If		
	
	iReturn = oPB3S116.B_QUERY_ITEM_ACCT_TRACKING(gStrGlobalCollection , strPlantCd , strItemAcct)
	
	If CheckSYSTEMError(Err,True) = True Then
		Set oPB3S116 = Nothing
		Response.End
	End If
	
	If Not (oPB3S116 Is Nothing) Then 
		Set oPB3S116 = Nothing
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

With parent
	LngMaxRow = .frm1.vspdData.MaxRows
		
<%  
	If Not (IsEmpty(iReturn)) Then
	
%>			
		ReDim TmpBuffer(<%=Ubound(iReturn ,1)%>)
<%
	End If
	
	For i=0 to Ubound(iReturn ,1)
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(iReturn(i,B494_EG1_plant_cd))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(iReturn(i,B494_EG1_item_acct))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(iReturn(i,B494_EG1_tracking_flag))%>"
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
Set ADF = Nothing												'¢Ð: ActiveX Data Factory Object Nothing
%>