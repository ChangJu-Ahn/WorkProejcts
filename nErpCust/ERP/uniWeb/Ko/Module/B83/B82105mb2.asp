<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../B81/B81COMM.ASP" -->

<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd 

Const Y105_I1_DT     = 0
Const Y105_I1_GRADE  = 1
Const Y105_I1_DESC   = 2
Const Y105_I1_PERSON = 3

Dim lgIntFlgMode
Dim iStrSelectChar
Dim iCisChangeItemReq
Dim iStrReqNo
Dim isUpMod           '검토단계구분(R:접수,T:기술,P:구매,Q:품질)


Dim PY2G105

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
isUpMod = Request("txtUpdMode")                     '☜: 검토단계구분(R:접수,T:기술,P:구매,Q:품질)
	    
iStrSelectChar = "UPDATE"
iStrReqNo = Request("txtReqNo")

Redim iCisChangeItemReq(3)

Select Case isUpMod
       Case "R"
           iCisChangeItemReq(Y105_I1_DT)      = uniconvDate(Request("htxtRDt"))
           iCisChangeItemReq(Y105_I1_GRADE)   = Request("htxtRGrade")
           iCisChangeItemReq(Y105_I1_DESC)    = Request("htxtRDesc")
           iCisChangeItemReq(Y105_I1_PERSON)  = Request("htxtRPerson")
       Case "T"
           iCisChangeItemReq(Y105_I1_DT)      = uniconvDate(Request("htxtTDt"))
           iCisChangeItemReq(Y105_I1_GRADE)   = Request("htxtTGrade")
           iCisChangeItemReq(Y105_I1_DESC)    = Request("htxtTDesc")
           iCisChangeItemReq(Y105_I1_PERSON)  = Request("htxtTPerson")
       Case "P"
           iCisChangeItemReq(Y105_I1_DT)      = uniconvDate(Request("htxtPDt"))
           iCisChangeItemReq(Y105_I1_GRADE)   = Request("htxtPGrade")
           iCisChangeItemReq(Y105_I1_DESC)    = Request("htxtPDesc")
           iCisChangeItemReq(Y105_I1_PERSON)  = Request("htxtPPerson")
       Case "Q"
           iCisChangeItemReq(Y105_I1_DT)      = uniconvDate(Request("htxtQDt"))
           iCisChangeItemReq(Y105_I1_GRADE)   = Request("htxtQGrade")
           iCisChangeItemReq(Y105_I1_DESC)    = Request("htxtQDesc")
           iCisChangeItemReq(Y105_I1_PERSON)  = Request("htxtQPerson")
End Select     

Set PY2G105 = Server.CreateObject("PY2G105.cCisChangeItemReqApp")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G105.Y_MAINT_CHANGE_ITEM_REQ_APP_SVR(gStrGlobalCollection, iStrSelectChar, iStrReqNo, isUpMod, iCisChangeItemReq )
	    

If CheckSYSTEMErrorY(Err,True,"검토자") = True Then
	Set PY2G102 = Nothing
		goFocus(iErrorPosition)
		Response.End

End If	    
		              
Set PY2G105 = Nothing

Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "With parent"				& vbCr	
Response.Write ".DbSaveOk"                  & vbCr
Response.Write "End With"                   & vbCr
Response.Write "</Script>"                  & vbCr
Response.End
	
%>