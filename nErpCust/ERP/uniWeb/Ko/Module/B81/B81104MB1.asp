<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : B81104MB1
'*  4. Program Name         : 품목구성코드등록 
'*  5. Program Desc         : 품목구성코드등록()
'*  6. Component List       : PM1G121.cMMntSpplItemPriceS
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="./B81COMM.ASP" -->
<%	
call LoadBasisGlobalInf()
'call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
'call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    'Dim lgOpModeCRUD
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim GroupCount  
    Dim lgPageNo
	Dim iErrorPosition
	Dim arrRsVal(11)
	Dim strSpread
	
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode") 
	 strSpread = Request("txtSpread")										                                              '☜: Read Operation 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim arr  
    ReDim arr(5)
    Dim lgstrdata
  ' on error resume next
   
    arr(0)=Request("txtItem_acct")
    arr(1)=Request("txtItem_kind")
    arr(2)=Request("txtItem_lvl1")
    arr(3)=Request("txtItem_lvl2")
    arr(4)=Request("txtItem_lvl3")
    
     '----- UI 각 항목 체크 ----
    Call SubOpenDB(lgObjConn) 
    call GetNameChk("max(spec_order)","B_CIS_ITEM_CLASS_CATEGORY","item_acct="&filterVar(arr(0),"''","S")&" and item_kind="&filterVar(arr(1),"''","S")&" and item_lvl1_cd="&filterVar(arr(2),"''","S")&" and item_lvl2_cd="&filterVar(arr(3),"''","S")&" and item_lvl3_cd="&filterVar(arr(4),"''","S")&"",				"1111","hMaxSpec_order","","N") 'maxSpec_order 구하기 
    call GetNameChk("minor_nm","b_minor","major_cd='P1001' and minor_cd="&filterVar(Request("txtItem_acct"),"''","S"),	Request("txtItem_acct"),"txtItem_acct","품목계정","Y") '품목계정
	call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&filterVar(Request("txtItem_kind"),"''","S"),	Request("txtItem_kind"),"txtItem_kind","품목구분","Y") '품목구분
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L1' and class_cd="&filterVar(Request("txtItem_lvl1"),"''","S"),	Request("txtItem_lvl1"),"txtItem_lvl1","대분류","Y") '대분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L2' and class_cd="&filterVar(Request("txtItem_lvl2"),"''","S"),	Request("txtItem_lvl2"),"txtItem_lvl2","중분류","Y") '중분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L3' and class_cd="&filterVar(Request("txtItem_lvl3"),"''","S"),	Request("txtItem_lvl2"),"txtItem_lvl3","소분류","Y") '소분류
	
	Call SubCloseDB(lgObjConn)  
	
	
	    
    lgstrdata = PY1G104.B_CIS_LIST(gStrGlobalCollection,arr)
   
	on error goto 0

	if lgstrdata="" then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	else
	
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData "			& vbCr
		Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & ConvSPChars(lgstrdata)	 & """" & ",""F""" & vbCr  
		Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr   
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr   
	end if
	 
End Sub    
	    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================

	
Sub SubBizSaveMulti()
   
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    '----- UI 각 항목 체크 ----
    Call SubOpenDB(lgObjConn) 
	call GetNameChk("minor_nm","b_minor","major_cd='P1001' and minor_cd="&filterVar(Request("txtItem_acct"),"''","S"),	Request("txtItem_acct"),"txtItem_acct","품목계정","Y") '품목계정
	call GetNameChk("minor_nm","b_minor","major_cd='Y1001' and minor_cd="&filterVar(Request("txtItem_kind"),"''","S"),	Request("txtItem_kind"),"txtItem_kind","품목구분","Y") '품목구분
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L1' and class_cd="&filterVar(Request("txtItem_lvl1"),"''","S"),	Request("txtItem_lvl1"),"txtItem_lvl1","대분류","Y") '대분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L2' and class_cd="&filterVar(Request("txtItem_lvl2"),"''","S"),	Request("txtItem_lvl2"),"txtItem_lvl2","중분류","Y") '중분류
	call GetNameChk("class_name","b_cis_item_class","item_acct='"&Request("txtItem_acct")&"' and item_kind ='"&Request("txtItem_kind")&"' and item_lvl='L3' and class_cd="&filterVar(Request("txtItem_lvl3"),"''","S"),	Request("txtItem_lvl2"),"txtItem_lvl3","소분류","Y") '소분류
	Call SubCloseDB(lgObjConn)  
		
  
    Call PY1G104.B_CIS_CTRL(gStrGlobalCollection,strSpread)
    If CheckSYSTEMError(Err,True) = True Then                                              
		Response.End 
    End If
 
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>

<%
End Sub



%>










<OBJECT RUNAT=server PROGID="PY1G104.cBCtrlBiz" id=PY1G104></OBJECT>
