<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: DT
'*  2. Function Name		: 
'*  3. Program ID			: d1211PB1.asp
'*  4. Program Name			: Digital Tax (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2009/12/20
'*  8. Modified date(Last)	: 2009/12/22
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M","NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs1										'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim i
Dim strData

Dim iTotalStr
Dim TmpBuffer1


    Call HideStatusWnd

	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "D1511QA11"
	
    UNIValue(0, 0) = "^"
    UNIValue(0, 1) = Replace(FilterVar(UniConvDate(Request("txtIssuedFromDt")), "''", "S"), "-", "")
    UNIValue(0, 2) = Replace(FilterVar(UniConvDate(Request("txtIssuedToDt")), "''", "S"), "-", "")

    If Request("popInvNo") = "" Then
       UNIValue(0, 3) = "|"
    Else
       UNIValue(0, 3) = FilterVar(UCase(Request("popInvNo")), "''", "S")
    End If
    
    If Request("popKeyNo") = "" Then
       UNIValue(0, 4) = "|"
    Else
       UNIValue(0, 4) = FilterVar(UCase(Request("popKeyNo")), "''", "S")
    End If

    If Request("popLegacyPk") = "" Then
       UNIValue(0, 5) = "|"
    Else
       UNIValue(0, 5) = FilterVar(UCase(Request("popLegacyPk")), "''", "S")
    End If

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If
		
		If Not(rs1.EOF And rs1.BOF) Then
		
			Redim TmpBuffer1(rs1.RecordCount-1)

			For i=0 to rs1.RecordCount-1
				strData = ""
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_no"             ))
                strData = strData & Chr(11) & ConvSPChars(rs1("attr01"             ))
                strData = strData & Chr(11) & ConvSPChars(rs1("legacy_pk"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_type1"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_type1_name"     ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_type2"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_type2_name"     ))
                strData = strData & Chr(11) & ConvSPChars(rs1("proc_flag"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("proc_flag_name"     ))
                strData = strData & Chr(11) & ConvSPChars(rs1("proc_date"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_reg_num"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_reg_id"         ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_cmp_name"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_owner"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_biz_type"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_biz_kind"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sup_address"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_reg_num"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_reg_id"         ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_cmp_name"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_owner"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_biz_type"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_biz_kind"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("dem_address"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_reg_num"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_reg_id"         ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_cmp_name"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_owner"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_biz_type"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_biz_kind"       ))
                strData = strData & Chr(11) & ConvSPChars(rs1("agn_address"        ))
                strData = strData & Chr(11) & ConvSPChars(rs1("amt_input_meth"     ))
                strData = strData & Chr(11) & ConvSPChars(rs1("pub_date"           ))
                strData = strData & Chr(11) & ConvSPChars(rs1("amt_blank"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("deal_type"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("deal_type_name"     ))
                strData = strData & Chr(11) & UniNumClientFormat(rs1("sup_tot_amt"        ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("sur_tax"            ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("sum_amt"            ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("cash_amt"           ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("check_amt"          ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("bill_amt"           ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & UniNumClientFormat(rs1("credit_amt"         ) ,ggAmtOfMoney.DecPoint,0)
                strData = strData & Chr(11) & ConvSPChars(rs1("issue_dtm"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("book_num1"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("book_num2"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("book_num3"          ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_amend_type"     ))
                strData = strData & Chr(11) & ConvSPChars(rs1("inv_amend_type_name"))
                strData = strData & Chr(11) & ConvSPChars(rs1("remark"             ))
                strData = strData & Chr(11) & ConvSPChars(rs1("remark2"            ))
                strData = strData & Chr(11) & ConvSPChars(rs1("remark3"            ))
                strData = strData & Chr(11) & ConvSPChars(rs1("sale_no"            ))

				strData = strData & Chr(11) & i
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(i) = strData
				rs1.MoveNext
				
			Next
			
		iTotalStr = Join(TmpBuffer1,"") 

		End If
		

		rs1.close

		Set rs1 = Nothing

        Set ADF = Nothing
        
%>	
		
    
<Script Language=vbscript>

    
    With parent												'☜: 화면 처리 ASP 를 지칭함 
		
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip  "<%=iTotalStr%>"
    
		.DbQueryOk()
		
    End With
    
</Script>	
