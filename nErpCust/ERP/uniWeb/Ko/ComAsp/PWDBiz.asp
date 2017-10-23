<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../inc/incSvrMain.asp"  -->
<!-- #Include file="../inc/incSvrDate.inc"  -->
<!-- #Include file="../inc/incSvrNumber.inc"  -->

<%
On Error Resume Next
Err.Clear 

Call HideStatusWnd
																			
'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Select Case Request("txtFlag")

	Case "L", "P"
		Dim iErrPosition	
		Dim importArray		
		Dim iZa001
		DIm Enc
		Dim strPwd
		
		Const C_SelectChar = 0
		Const C_UsrId = 1
		Const C_OldPwd = 2
		Const C_NewPwd = 3
		Const C_ConfirmPwd = 4
		Const C_UpdtUsrId = 5		
		
		Redim importArray(C_UpdtUsrId)
		
		importArray(C_SelectChar) = Request("txtFlag")
		importArray(C_UsrId)      = gUsrId
		importArray(C_OldPwd)     = UnEscape(Request("txtOld"))
		importArray(C_NewPwd)     = UnEscape(Request("txtNew"))				
		importArray(C_ConfirmPwd) = UnEscape(Request("txtRe"))
		importArray(C_UpdtUsrId)  = gUsrId
		
		Set iZa001 = Server.CreateObject("PZAG001.cCtrlUsrMastRec")		

		If CheckSYSTEMError(Err,True) = True Then
			Response.End
		End If

		Call iZa001.ZA_Update_Usr_Mast_Rec_Pwd(gStrGlobalCollection,importArray,iErrPosition)
		
   		If Not IsEmpty(importArray) Then
   			Erase ImportArray
   		End If		

        If Err.number <> 0 Then
		   Response.Write "<Script Language=vbscript>"   & vbCr
	       Response.Write "Parent.HandleError(""" & Err.Description  & """)"        & vbCr
		   Response.Write "</Script>"                    & vbCr   
        End If

		If CheckSYSTEMError2(Err, True, iErrPosition & "행:","","","","") = True Then
		   Set iZa001 = Nothing
		   Response.End 
		End If

		Set iZa001 = Nothing	
		
 	    Call DisplayMsgBox("210026", vbOKOnly, "", "", I_MKSCRIPT)					 '비밀번호가 변경되었습니다!						   
    
		Response.Write "<Script Language=vbscript>"   & vbCr
		Response.Write "Parent.SaveOk "            & vbCr
		Response.Write "</Script>"                    & vbCr   
		     			
End Select				
%>