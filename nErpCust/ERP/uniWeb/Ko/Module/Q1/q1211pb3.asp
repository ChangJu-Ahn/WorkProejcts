<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1211PB3
'*  4. Program Name         : AQL 팝업 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG002
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Park Hyun Soo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call loadInfTB19029B("I", "*", "NOCOOKIE","PB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
																			
Dim PQBG002
Dim lgstr
Dim StrData
Dim i

	
Dim strAQLCd
Dim strMAJOR_CD
Dim EG1_group_export
ReDim EG1_group_export(0)

strAQLCd = UniConvNum(Request("strAQLCd"), 0)
strMAJOR_CD = Request("strMAJOR_CD")

Set PQBG002 = Server.CreateObject("PQBG002.cQListAQLSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
		
Call PQBG002.Q_LIST_AQL_SVR(gStrGlobalCollection, _
			 				strAQLCd, _
							strMAJOR_CD, _
							EG1_group_export)

If CheckSYSTEMError(Err,True) = True Then
	Set PQBG002 = Nothing
	Response.End
End If

Dim TmpBuffer
Dim iTotalStr
ReDim TmpBuffer(UBound(EG1_group_export, 1))

For i = 0 To UBound(EG1_group_export, 1)
	TmpBuffer(i) = Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_group_export(i, 1), 0) & _
				   Chr(11) & i + 1 & _
				   Chr(11) & Chr(12)
Next

iTotalStr = Join(TmpBuffer, "")	

Set PQBG002 = nothing	

%>		    
<Script Language="vbscript">   
With parent
	.ggoSpread.Source = .vspdData
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
	.DbQueryOk
	.vspdData.focus
End With
</Script>