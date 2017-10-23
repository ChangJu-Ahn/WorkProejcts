<%@LANGUAGE = VBScript%>
<%Option Explicit%>


<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<!-- #Include file="../B81/B81COMM.ASP" -->

<%
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B82110MB1
'*  4. Program Name         : 승인코드ERP전송 
'*  5. Program Desc         : 승인코드ERP전송 
'*  6. Component List       : PY2G110
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
'Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd
 
Const C_SHEETMAXROWS_D = 1000

Dim PY2G110

       
Const Y110_E1_TRANS      = 0
Const Y110_E1_REQ_NO     = 1
Const Y110_E1_REQ_ID     = 2
Const Y110_E1_REQ_NM     = 3
Const Y110_E1_REQ_DT     = 4
Const Y110_E1_REQ_GBN    = 5
Const Y110_E1_REQ_GBN_NM = 6
Const Y110_E1_ITEM_CD    = 7
Const Y110_E1_ITEM_NM    = 8
Const Y110_E1_ITEM_SPEC  = 9
Const Y110_E1_END_DT     = 10
Const Y110_E1_TRANS_DT   = 11
Const Y110_E1_REQ_REASON = 12
Const Y110_E1_REMARK     = 13

Dim StrNextKey		' 다음 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim lgStrPrevKey
Dim strData

Dim strDtFr
Dim strDtTo 
Dim strTransB
Dim strReqId
Dim strItemAcct
Dim strItemKind
Dim strTransA
Dim strReqNo
Dim ArrTransItemERP
                                  
LngMaxRow = Request("txtMaxRows")

strDtFr      = UNIConvDate(Request("txtDtFr"))
strDtTo      = UNIConvDate(Request("txtDtTo"))
strTransB    = Request("htxtTransB")
strReqId     = Request("txtreq_user")
strItemAcct  = Request("cboItemAcct")
strItemKind  = Request("txtItem_Kind")
strTransA    = Request("htxtTransA")	
lgStrPrevKey = Request("lgStrPrevKey")

Call SubOpenDB(lgObjConn)        
call GetNameChk("MINOR_NM","B_MINOR","MINOR_CD="& FilterVar(strReqId,"","S")& " AND MAJOR_CD=" & filterVar("Y1006","''","S") ,	strReqId,"txtreq_user","의뢰자","Y") '의뢰자 
call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&FilterVar(strItemKind,"","S")&"",strItemKind,"txtItem_Kind","품목구분","Y") '품목구분 
 Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

Set PY2G110 = Server.CreateObject("PY2G110.cYTransItemERPQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End if

Call PY2G110.Y_QUERY_TRANS_ITEM_ERP(gStrGlobalCollection, _
                                    C_SHEETMAXROWS_D, _
                                    strDtFr, strDtTo, strTransB, strReqId, strItemAcct, strItemKind, strTransA, lgStrPrevKey, _                                     
                                    ArrTransItemERP )
		   '    800186 900014
err.Description =   replace(err.Description ,"800186","900014") 
If CheckSYSTEMError(Err,True) Then
	Set PY2G110 = Nothing
	Response.End
End If

Dim iTotalStr
Dim TmpBuffer

ReDim TmpBuffer(UBound(ArrTransItemERP))

For LngRow = 0 To UBound(ArrTransItemERP)
	
	If LngRow < C_SHEETMAXROWS_D Then
			
		strData = Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_TRANS)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_NO)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_ID)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_NM)) & _
				  Chr(11) & UniConvDateDbToCompany(ArrTransItemERP(LngRow,Y110_E1_REQ_DT),"") & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_GBN)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_GBN_NM)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_ITEM_CD)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_ITEM_NM)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_ITEM_SPEC)) & _
				  Chr(11) & UniConvDateDbToCompany(ArrTransItemERP(LngRow,Y110_E1_END_DT),"") & _
				  Chr(11) & UniConvDateDbToCompany(ArrTransItemERP(LngRow,Y110_E1_TRANS_DT),"") & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REQ_REASON)) & _
				  Chr(11) & ConvSPChars(ArrTransItemERP(LngRow,Y110_E1_REMARK)) & _
				  Chr(11) & LngMaxRow + LngRow + 1 & Chr(11) & Chr(12)
			
		TmpBuffer(LngRow) = strData
	Else
		StrNextKey = ArrTransItemERP(LngRow,Y110_E1_REQ_NO)
	End If
Next

iTotalStr = Join(TmpBuffer, "")

%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	' Request값을 hidden input으로 넘겨줌 
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.DbQueryOk
    End If
	
End with

</Script>
<%
Set PY2G110 = Nothing
%>