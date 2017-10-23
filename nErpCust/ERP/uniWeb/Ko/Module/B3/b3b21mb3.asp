<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b21mb3.asp
'*  4. Program Name         : Delete Characteristic 삭제 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim oPB3S211

Dim StrNextKey
Dim lgStrPrevKey
Dim LngRow
Dim GroupCount 

Dim lgIntFlgMode
Dim iCommandSent

Dim I1_B_Char
Dim IG1_B_Char_Value

ReDim I1_B_Char(3)
Const C_CommandSent = 0
Const C_Char_Cd = 1
Const C_Char_Nm = 2
Const C_Char_Value_Digit = 3
ReDim IG1_B_Char_Value(0,0)

Call HideStatusWnd

On Error Resume Next
Err.Clear

    If Request("txtCharCd1") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)             
		Response.End 
	End If
	
	I1_B_Char(C_CommandSent) = "DELETE"
	I1_B_Char(C_Char_Cd) = UCase(Trim(Request("txtCharCd1")))

	Set oPB3S211 = Server.CreateObject("PB3S211.cBMngChar")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

	Call oPB3S211.B_MANAGE_CHAR(gStrGlobalCollection, _
								I1_B_Char, _
								IG1_B_Char_Value)
    
    If CheckSYSTEMError(Err,True) = True Then
       Set oPB3S211 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End		
    End If

	Set oPB3S211 = Nothing															'☜: Unload Comproxy DLL
    
Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "	With parent			  " & vbCr													'☜: 화면 처리 ASP 를 지칭함 
Response.Write "		.DbDeleteOk		  " & vbCr
Response.Write " 	End With			  " & vbCr
Response.Write "</Script>				  " & vbCr

%>