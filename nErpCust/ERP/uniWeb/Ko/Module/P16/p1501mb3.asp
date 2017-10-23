<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1501mb2.asp
'*  4. Program Name         : ManageResource 삭제 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim oPP1G606

Dim StrNextKey
Dim lgStrPrevKey
Dim LngRow
Dim GroupCount 

Dim lgIntFlgMode
Dim iCommandSent
Dim I1_P_Resource			'Array
Dim I2_P_Resource_Group_Cd	'String
Dim I3_B_Plant_Cd			'String

ReDim I1_P_Resource(10)
Const C_I1_Resource_Cd = 0

Call HideStatusWnd

On Error Resume Next
Err.Clear

    If Request("txtResourceCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)             
		Response.End 
	End If
	
	iCommandSent = "DELETE"
	I3_B_Plant_Cd = Trim(Request("txtPlantCd"))
	I1_P_Resource(C_I1_Resource_Cd) = Trim(Request("txtResourceCd"))
	
	Set oPP1G606 = Server.CreateObject("PP1G606.cPMngRsrc")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

    Call oPP1G606.P_MANAGE_RESOURCE(gStrGlobalCollection, _
								iCommandSent, _
								I1_P_Resource, _
								I2_P_Resource_Group_Cd, _
								I3_B_Plant_Cd)
    
    If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G606 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End		
    End If

	Set oPP1G606 = Nothing															'☜: Unload Comproxy DLL
    
Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "	With parent			  " & vbCr													'☜: 화면 처리 ASP 를 지칭함 
Response.Write "		.DbDeleteOk		  " & vbCr
Response.Write " 	End With			  " & vbCr
Response.Write "</Script>				  " & vbCr

%>