<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","PB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/06/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Youngju Yim
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd 

Const C_SHEETMAXROWS_D = 100

'EXPORT VIEW

Const	E1_QueryID		=	0
Const	E1_QueryNM		=	1
Const	E1_DeptCD		=	2
Const	E1_DeptNM		=	3
Const	E1_SELECT		=	4
Const	E1_FROM			=	5
Const	E1_WHERE		=	6
Const	E1_ETC			=	7
Const	E1_REMARK		=	8
	
    
Dim PB1Z001_KO346

Dim LngMaxRow
Dim StrNextKey1
Dim StrNextKey2
Dim i
Dim StrData

Dim iExportDisposition

Dim strQueryFlag
Dim txtQueryID 
Dim txtQueryNm
Dim txtDept_cd 
Dim cboRole_type 

txtQueryID		= Request("txtQueryID")
txtQueryNm		= Request("txtQueryNm")
txtDept_cd		= Request("txtDept_cd")
cboRole_type	= Request("cboRole_type")
gDepart			= Request("gDepart")


' 해당 Business Object 생성 
Set PB1Z001_KO346 = Server.CreateObject("PB1Z001_KO346.clsQueryCommand")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PB1Z001_KO346.B_QUERY_LIST_SVR(gStrGlobalCollection, _
									C_SHEETMAXROWS_D, _
									txtQueryID, _
									txtQueryNm, _
									txtDept_cd, _
									cboRole_type, _
									gDepart, _
									iExportDisposition)


If CheckSYSTEMError(Err,True) = True Then
	Set PB1Z001_KO346 = Nothing
	Response.End
End If


For i = 0 To UBound(iExportDisposition, 1)
    If i < C_SHEETMAXROWS_D Then
    	strData = strData & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_QueryID))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_QueryNM))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_DeptCD))) _
						  & Chr(11) & ConvSPChars(Trim(iExportDisposition(i, E1_DeptNM))) _
						  & Chr(11) & Replace(Replace(ConvSPChars(Trim(iExportDisposition(i, E1_SELECT))),chr(10),""),chr(13),"") _
						  & Chr(11) & Replace(Replace(ConvSPChars(Trim(iExportDisposition(i, E1_FROM))),chr(10) ,""),chr(13),"") _
						  & Chr(11) & Replace(Replace(ConvSPChars(Trim(iExportDisposition(i, E1_WHERE))),chr(10) ,""),chr(13),"") _
						  & Chr(11) & Replace(Replace(ConvSPChars(Trim(iExportDisposition(i, E1_ETC))),chr(10) ,""),chr(13),"") _
						  & Chr(11) & Replace(Replace(ConvSPChars(Trim(iExportDisposition(i, E1_REMARK))),chr(10) ,""),chr(13),"") _
						  & Chr(11) & Chr(12)
    Else
		StrNextKey1 = ConvSPChars(Trim(iExportDisposition(i, E1_QueryID)))
		'StrNextKey2 = ConvSPChars(Trim(iExportDisposition(i, E1_QueryID)))
    End If
Next  

Set PB1Z001_KO346 = Nothing
%>
<Script Language="vbscript">   
    With parent
		.ggoSpread.Source = .vspdData 
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		.vspdData.focus
		.DbQueryOk
	    
	End with
</Script>

