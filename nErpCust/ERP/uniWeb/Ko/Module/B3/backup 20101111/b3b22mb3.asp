<%@ LANGUAGE=VBSCript %>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22mb3.asp 
'*  4. Program Name         : Called By B3B22MA1 (Class Management)
'*  5. Program Desc         : Manage Class Information(Delete)
'*  6. Modified date(First) : 2003/02/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim oPB3S222

Dim StrNextKey
Dim lgStrPrevKey
Dim strClassCd

Dim lgIntFlgMode
Dim iCommandSent
Dim I1_B_Class

ReDim I1_B_Class(4)
Const C_I1_Class_Cd = 0

Call HideStatusWnd

On Error Resume Next
Err.Clear

    strClassCd = Request("txtClassCd1")

    If Trim(strClassCd) = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("122650", vbOKOnly, "", "", I_MKSCRIPT)             
		Response.End 
	End If
	
	iCommandSent = "DELETE"
	I1_B_Class(C_I1_Class_Cd) = UCase(Trim(strClassCd))

	Set oPB3S222 = Server.CreateObject("PB3S222.cBMngClass")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

    Call oPB3S222.B_MANAGE_CLASS(gStrGlobalCollection, _
								iCommandSent, _
								I1_B_Class)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPB3S222 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End		
    End If
	
	Set oPB3S222 = Nothing								'☜: Unload Comproxy DLL

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "	With parent			  " & vbCr													'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "		.DbDeleteOk		  " & vbCr
	Response.Write " 	End With			  " & vbCr
	Response.Write "</Script>				  " & vbCr

	Response.End
%>