<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81109MB1
'*  4. Program Name         : 품목기본정보조회
'*  5. Program Desc         : 품목기본정보조회
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/30
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
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%	
call LoadBasisGlobalInf()



	Dim istrData
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim GroupCount  
    Dim lgPageNo
	Dim iErrorPosition
	Dim arrRsVal(11)
	Dim strSpread
	
	Dim seq_no
	Dim item_cd

	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode")
    item_cd	      = filterVar(Request("item_cd"),"''","S")
     seq_no	      = Request("seq_no")
     
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call SubOpenDB(lgObjConn) 
	

	Select Case lgOpModeCRUD
        Case CStr(UID_M0001)    
                Call  SubBizQueryView()                                             '☜: Query
                Call  SubBizQueryMulti()
            
        Case CStr(UID_M0002) 
                                                                '☜: Save,Update
        Case CStr(UID_M0003)
        
        Case "VIEW"
           ' Call  SubBizQueryView() 나중에 추가시 사용!
             
    End Select
    
    
	
	
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim getArr
	Dim StrGbn
	dim dataF

	
     '-----------------------------
	  ' dbQuery
	  '-----------------------------
	'*@@ :구분자일뿐 
	StrGbn="L"
	lgStrSQL="exec USP_B81109M_LST *@@"&StrGbn&"*@@,*@@"&item_cd&"*@@,*@@"&seq_no&"*@@"
	
	lgStrSQL=replace(lgStrSQL,"'","''")
	
	lgStrSQL=replace(lgStrSQL,"*@@","'")
	
	
	adoRec.Open lgStrSQL, lgObjConn, 1 ,1 
	
	
	if adoRec.EOF  then
		adoRec.Close 
		Call SubCloseDB(lgObjConn)      
		'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	else
		getArr = adoRec.GetRows()
		adoRec.Close 
		Call SubCloseDB(lgObjConn)  
		Call  ListupDataGrid  (getArr,",3,")
		Response.End
	end if
End Sub   



Sub SubBizQueryMulti()
	Dim getArr
	Dim StrGbn
	dim dataF

	
     '-----------------------------
	  ' dbQuery
	  '-----------------------------
	'*@@ :구분자일뿐 
	StrGbn="L"
	lgStrSQL="exec USP_B81109M_LST *@@"&StrGbn&"*@@,*@@"&item_cd&"*@@,*@@"&seq_no&"*@@"
	
	lgStrSQL=replace(lgStrSQL,"'","''")
	
	lgStrSQL=replace(lgStrSQL,"*@@","'")
	adoRec.Open lgStrSQL, lgObjConn, 1 ,1 
	
	
	if adoRec.EOF  then
		adoRec.Close 
		Call SubCloseDB(lgObjConn)      
		'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	else
		getArr = adoRec.GetRows()
		adoRec.Close 
		Call SubCloseDB(lgObjConn)  
		Call  ListupDataGrid  (getArr,",3,")
		Response.End
	end if
End Sub    
	     
	    


'============================================================================================================
' Name : SubBizQueryView
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryView()
	Dim getArr,strGbn

     '-----------------------------
	  ' dbQuery
	  '-----------------------------
	'*@@ :구분자일뿐 
	Call  ClearData() 
	' call GetNameChk("item_nm","b_cis_item_master","item_cd="&filtervar(request("item_cd"),"''","S")&"","item_cd","txtItem_cd","품목코드","Y") '품목code
	 
	 
	strGbn="V"
	lgStrSQL="exec USP_B81109M_LST *@@"&StrGbn&"*@@,*@@"&item_cd&"*@@,*@@"&seq_no&"*@@"
	lgStrSQL=replace(lgStrSQL,"'","''")
	
	lgStrSQL=replace(lgStrSQL,"*@@","'")
	
	adoRec.Open lgStrSQL, lgObjConn, 1 ,1 

	if adoRec.EOF  then
		adoRec.Close 
		
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	else
		getArr = adoRec.GetRows()
		adoRec.Close 

		Call  setData  (getArr)
	
	end if
End Sub    




'============================================================================================================
'setData
'============================================================================================================

	    
Sub setData(pgetArr)
	Dim strData
	Dim i,j
	
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		
		'Response.Write "	.txtitem_nm2.value			=  """	& pgetArr(0,0)	&""""	& vbCr
		'Response.Write "	.txtitem_spec2.value		=  """	& pgetArr(1,0)	&""""	& vbCr
		Response.Write "	.txtReqdt.text				=  """	& UNIDateClientFormat(pgetArr(2,0))	&""""	& vbCr
		Response.Write "	.txtEndDt.text				=  """	& UNIDateClientFormat(pgetArr(3,0))	&""""	& vbCr
		Response.Write "	.txtitem_lvl1.value			=  """	& pgetArr(4,0)	&""""	& vbCr
		Response.Write "	.txtitem_lvl1_nm.value		=  """	& pgetArr(5,0)	&""""	& vbCr
		Response.Write "	.txtitem_lvl2.value			=  """	& pgetArr(6,0)	&""""	& vbCr
		Response.Write "	.txtitem_lvl2_nm.value		=  """	& pgetArr(7,0)	&""""	& vbCr
		Response.Write "	.txtitem_lvl3.value			=  """	& pgetArr(8,0)	&""""	& vbCr
		Response.Write "	.txtitem_lvl3_nm.value		=  """	& pgetArr(9,0)	&""""	& vbCr
		
		Response.Write "	.txtPur_vendor.value		=  """	& pgetArr(12,0)	&""""	& vbCr
		Response.Write "	.txtPur_vendor_nm.value		=  """	& pgetArr(13,0)	&""""	& vbCr
		Response.Write "	.txtPur_type.value			=  """	& pgetArr(10,0)	&""""	& vbCr
		Response.Write "	.txtItem_unit.value			=  """	& pgetArr(11,0)	&""""	& vbCr
		
		'esponse.Write "	.txtPur_person.value		=  """	& pgetArr(14,0)	&""""	& vbCr
		'esponse.Write "	.txtPur_person_nm.value		=  """	& pgetArr(15,0)	&""""	& vbCr
		
		Response.Write "	.trans_date.value		=  """	& pgetArr(16,0)	&""""	& vbCr
		Response.Write "	.txtReq_id.value		=  """	& pgetArr(17,0)	&""""	& vbCr
		Response.Write "	.txtReq_id_nm.value		=  """	& pgetArr(18,0)	&""""	& vbCr
		Response.Write "	.txtStatus.value		=  """	& pgetArr(19,0)	&""""	& vbCr
		Response.Write "	.txtdoc_no.value		=  """	& pgetArr(20,0)	&""""	& vbCr
		Response.Write "	.txtdoc_no.value		=  """	& pgetArr(20,0)	&""""	& vbCr
		Response.Write "	.txtReq_reason.value	=  """	& pgetArr(21,0)	&""""	& vbCr
		Response.Write "	.txtitem_spec.value		=  """	& pgetArr(22,0)	&""""	& vbCr

		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
End Sub	


'============================================================================================================
'ClearData
'============================================================================================================
   
Sub ClearData
	Dim strData
	Dim i,j
	
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		
		'Response.Write "	.txtitem_nm2.value			=  """""	& vbCr
		'Response.Write "	.txtitem_spec2.value		=  """""	& vbCr
		Response.Write "	.txtReqdt.text				=  """""	& vbCr
		Response.Write "	.txtEndDt.text				=  """""	& vbCr
		Response.Write "	.txtitem_lvl1.value			=  """""	& vbCr
		Response.Write "	.txtitem_lvl1_nm.value		=  """""	& vbCr
		Response.Write "	.txtitem_lvl2.value			=  """""	& vbCr
		Response.Write "	.txtitem_lvl2_nm.value		=  """""	& vbCr
		Response.Write "	.txtitem_lvl3.value			=  """""	& vbCr
		Response.Write "	.txtitem_lvl3_nm.value		=  """""	& vbCr
		Response.Write "	.txtPur_vendor.value		=  """""	& vbCr
		Response.Write "	.txtPur_vendor_nm.value		=  """""	& vbCr
		Response.Write "	.txtPur_type.value			=  """""	& vbCr
		Response.Write "	.txtItem_unit.value			=  """""	& vbCr
		Response.Write "	.trans_date.value		= """""	& vbCr
		Response.Write "	.txtReq_id.value		= """""	& vbCr
		Response.Write "	.txtReq_id_nm.value		= """""	& vbCr
		Response.Write "	.txtStatus.value		= """""	& vbCr
		Response.Write "	.txtdoc_no.value		= """""	& vbCr
		Response.Write "	.txtdoc_no.value		= """""	& vbCr
		Response.Write "	.txtReq_reason.value	= """""	& vbCr
		Response.Write "	.txtitem_spec.value		=  """""	& vbCr
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
End Sub	


%>




















<OBJECT RUNAT=server PROGID=ADODB.Recordset id=adoRec></OBJECT>
