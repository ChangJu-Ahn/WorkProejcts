<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 배부C/C 그룹등록 
'*  3. Program ID           : c40001mb1
'*  4. Program Name         : 배부C/C 그룹등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2005/08/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : choe0tae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") %>

<%													

On Error Resume Next														

Call HideStatusWnd

Const	C_PlantCd		= 1
Const	C_SalesGrp		= 2
Const	C_ItemCd		= 3
Const	C_SoldToParty	= 4
Const	C_LocExpFlag	= 5

Dim iStrMode
Dim iStrSvrData, iStrSvrData2, iStrNextKey
Dim iObjPC4G001
Dim iArrListOut			' Result of recordset.getrow(), it means iArrListOut is two dimension array (column, row)
Dim iArrListGroupOut	' Result of recordset.getrow(), it means iArrListGroupOut is two dimension array (column, row)
Dim iArrWhereIn, iArrWhereOut
Dim iLngRow
Dim iLngLastRow			' The last row number in the spread
Dim iLngSheetMaxRows	' Row numbers to be displayed in the spread.
Dim iLngErrorPosition

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	' -- MA에서 처리됨 
	
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing

	Dim arrData, sMsg
	
	' -- MA에서 보내온 데이타를 배열로 분리한다.
	arrData = Split(Request("txtSpread"), gColSep)
	
    Set iObjPC4G001 = Server.CreateObject("PC4G001.cCMngCostCenterSvr")  
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iObjPC4G001 = Nothing
		Response.End		
    End If
    
	Call iObjPC4G001.C_MANAGER_COST_CENTER_SVR(gStrGlobalCollection, arrData, sMsg)
	
    If CheckSYSTEMError2(Err,True, UCase(arrData(2)), "", "", "", "") = True Then
       Set iObjPC4G001 = Nothing
	   Response.End 
	End If

    Set iObjPC4G001 = Nothing	

    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> " 													'☜: Row 의 상태 
    Response.End 
End Select

%>

