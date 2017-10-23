<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7108mb1
'*  4. Program Name         : ������ ����ȸ 
'*  5. Program Desc         : �����ڻ꺰 �������� ��ȸ 
'*  6. Comproxy List        : +As0069LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2002/03/26
'*  8. Modified date(Last)  : 2002/03/26
'*  9. Modifier (First)     : Ȳ���� 
'* 10. Modifier (Last)      : Ȳ���� 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                       
'********************************************************************************************** -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim txtAsstNo
DIm txtDepryyyymm
Dim DurMnthFg
Dim Asst_NM

Dim txtAcctCd
Dim txtAcctNm
Dim cboDeprMthd
Dim txtRegDt
Dim txtLocAcqAmt
Dim txtAcqQty
Dim txtInvQty
Dim txtDurMnth
'Dim txtDeprRate

Dim txtEndLTermAcqAmt			'���⸻ ���� 
Dim txtEndLTermCptAmt
Dim txtEndLTermDeprAmt
Dim txtEndLTermBalAmt
Dim txtEndLTermInvQty

Dim txtFMnthAcqAmt				'����� 
Dim txtFMnthCptAmt
Dim txtFMnthDeprAmt
Dim txtFMnthBalAmt
Dim txtFMnthInvQty

Dim txtMnthCptAmt				'����߻� 
Dim txtMnthDeprAmt
Dim txtMnthDisUseQty

Dim txtLMnthAcqAmt				'����� 
Dim txtLMnthCptAmt
Dim txtLMnthDeprAmt
Dim txtLMnthBalAmt
Dim txtLMnthInvQty


' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "MB")   'ggQty.DecPoint Setting...

	lgDataExist			= "No"
	txtAsstNo			= Trim(Request("txtAsstNo"))
	txtDepryyyymm		= Trim(Request("txtDepryyyymm"))
	DurMnthFg			= Trim(Request("DurMnthFg"))

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim strWhereUP
	Dim strWhereDown
	
    Redim UNIValue(2,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	If DurMnthFg= "T" then
		UNISqlId(0) = "A7112MA01KO441"	'���� 
	else
		UNISqlId(0) = "A7112MA03KO441"	'���� 
	End If
	
    UNISqlId(1) = "A7112MA02"	'���� 
    UNISqlId(2) = "commonqry"	

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,0) = lgSelectList                                          '��: Select list
    
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	If DurMnthFg= "T" then
		'strWhereUP	= " and e.dur_yrs = a.tax_dur_yrs"
	else
		'strWhereUP	= " and e.dur_yrs = a.cas_dur_yrs"
	End If

	strWhereUP = ""
	strWhereUP = strWhereUP & " and d.major_cd	= " & FilterVar("a2002" , "''", "S") & ""  
	strWhereUP = strWhereUP & " and asst_no		= " & FilterVar(txtAsstNo , "''", "S") 
'Call ServerMesgBox(strWhereUP , vbInformation, I_MKSCRIPT)
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strWhereUP	= strWhereUP	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


	
    strWhereDown = "asst_no = " & FilterVar(txtAsstNo , "''", "S")
    strWhereDown =  strWhereDown & " and depr_yyyymm	= " & FilterVar(txtDepryyyymm , "''", "S")
    strWhereDown =  strWhereDown & " and dur_yrs_fg		= " & FilterVar(DurMnthFg ,"''" ,"S")	'ȸ����ر���

	' �ڻ��ڵ�� �������Ƿ� �������� ���Ѱ��� ���� ���� 

	UNIValue(0,0) = strWhereUP
	UNIValue(1,0) = strWhereDown
	UNIValue(2,0) = "Select Asst_NM from a_asset_master(NOLOCK) where asst_no=" & FilterVar(txtAsstNo , "''", "S")
	    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'--------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF
                                                                      '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)   
   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    'rs2 �ڻ�� �������� 
    If Not(rs2.EOF Or rs2.BOF) Then
		Asst_NM = Trim(rs2("Asst_NM"))		
		rs2.Close
		Set rs2 = Nothing   
	Else
		Asst_NM = ""	
		Call DisplayMsgBox("117400", vbOKOnly, "", "", I_MKSCRIPT)     'No Data Found!!			
 %>
	<Script Language=vbscript>
		parent.frm1.txtAsstNo.focus
	</Script>
 <%
      rs2.Close
      Set rs2 = Nothing
	  Response.End
	  Exit sub 
	End IF
	
	If NOt(rs0.EOF or rs0.BOF) Then
		Call  MakeDataUP()
    Else 		
		Call DisplayMsgBox("117500", vbOKOnly, "", "", I_MKSCRIPT)     'No Data Found!!
 %>
	<Script Language=vbscript>
		parent.frm1.txtAsstNo.focus
	</Script>
 <%
        rs0.Close
        Set rs0 = Nothing
			Response.End 
        Exit Sub
    End If
	
	'rs1	
    If Not (rs1.EOF OR rs1.BOF) Then
		Call MakeDataDown()
	Else
		Call DisplayMsgBox("117500", vbOKOnly, "", "", I_MKSCRIPT)
%>
	<Script Language=vbscript>
		parent.frm1.txtDeprYyyymm.focus
	</Script>
 <%
			rs1.Close
			Set rs1 = Nothing
			Exit Sub
		
	End IF
End Sub
'----------------------------------------------------------------------------------------------------------
' Make Data 
'--------------------------------------------------------------------------------------------------------
Sub MakeDataUP()

	txtAcctCd	= Rs0(0)
	txtAcctNm	= Rs0(1)
	cboDeprMthd	= Rs0(2)
	txtRegDt	= Rs0(3)
	txtLocAcqAmt= Rs0(4)
	txtAcqQty	= Rs0(5)
	txtInvQty	= Rs0(6)
	txtDurMnth	= Rs0(7)
	'txtDeprRate	= Rs0(8)
End Sub

Sub MakeDataDown()

	lgDataExist    = "Yes"
	
	txtEndLTermAcqAmt	= Rs1(0)			       '���⸻ ���� 
	txtEndLTermCptAmt	= Rs1(1)	
	txtEndLTermDeprAmt	= Rs1(2)	
	txtEndLTermBalAmt	= Rs1(3)	
	txtEndLTermInvQty	= Rs1(4)	

	txtFMnthAcqAmt	= Rs1(5)	'����� 
	txtFMnthCptAmt	= Rs1(6)	
	txtFMnthDeprAmt	= Rs1(7)	
	txtFMnthBalAmt	= Rs1(8)	
	txtFMnthInvQty	= Rs1(9)	

	txtMnthCptAmt	= Rs1(10)	'����߻� 
	txtMnthDeprAmt	= Rs1(11)	
	txtMnthDisUseQty= Rs1(12)	

	txtLMnthAcqAmt	= Rs1(13)	'����� 
	txtLMnthCptAmt	= Rs1(14)	
	txtLMnthDeprAmt	= Rs1(15)	
	txtLMnthBalAmt	= Rs1(16)	
	txtLMnthInvQty	= Rs1(17)	

End Sub

%>

<Script Language=vbscript>

With Parent

	If "<%=lgDataExist%>" = "Yes" Then

    	.frm1.txtAcctCd.Value				= "<%=ConvSPChars(txtAcctCd)%>"				'�����ڵ� 
		.frm1.txtAcctNm.value				= "<%=ConvSPChars(txtAcctNm)%>"				'������ 
		.frm1.cboDeprMthd.value				= "<%=ConvSPChars(cboDeprMthd)%>"			'�󰢹�� 
		.frm1.txtRegDt.text					= "<%=UNIDateClientFormat(txtRegDt)%>"				'������� 

		.frm1.txtLocAcqAmt.value			= "<%=UNINumClientFormat(txtLocAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"		'���ݾ�(�ڱ�)
		.frm1.txtAcqQty.value				= "<%=UNINumClientFormat(txtAcqQty, ggQty.DecPoint, 0)%>"				'������ 
		.frm1.txtInvQty.Value				= "<%=UNINumClientFormat(txtInvQty, ggQty.DecPoint, 0)%>"				'������ 
		.frm1.txtDurMnth.value				= "<%=txtDurMnth%>"														'������� 
		'.frm1.txtDeprRate.value				= "<%=UNINumClientFormat(txtDeprRate, ggExchRate.DecPoint, 0)%>"			'���� 

		'''''���⸻ 
		.frm1.txtEndLTermAcqAmt.value		= "<%=UNINumClientFormat(txtEndLTermAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"	'��氡�� 
		.frm1.txtEndLTermCptAmt.value		= "<%=UNINumClientFormat(txtEndLTermCptAmt, ggAmtOfMoney.DecPoint, 0)%>"	'�ں������� 
		.frm1.txtEndLTermDeprAmt.value		= "<%=UNINumClientFormat(txtEndLTermDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"	'�󰢾� 
		.frm1.txtEndLTermBalAmt.value		= "<%=UNINumClientFormat(txtEndLTermBalAmt, ggAmtOfMoney.DecPoint, 0)%>"	'�̻󰢾� 
		.frm1.txtEndLTermInvQty.value		= "<%=UNINumClientFormat(txtEndLTermInvQty, ggQty.DecPoint, 0)%>"	'��� 
		
		'''''�����			
		.frm1.txtFMnthAcqAmt.value			= "<%=UNINumClientFormat(txtFMnthAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"				'��氡�� 
		.frm1.txtFMnthCptAmt.value			= "<%=UNINumClientFormat(txtFMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"			'�ں������� 
		.frm1.txtFMnthDeprAmt.value			= "<%=UNINumClientFormat(txtFMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'�󰢾� 
		.frm1.txtFMnthBalAmt.value			= "<%=UNINumClientFormat(txtFMnthBalAmt, ggAmtOfMoney.DecPoint, 0)%>"			'�̻󰢾� 
		.frm1.txtFMnthInvQty.value			= "<%=UNINumClientFormat(txtFMnthInvQty, ggQty.DecPoint, 0)%>"			'��� 
		'''''����߻�			
		.frm1.txtMnthCptAmt.value			= "<%=UNINumClientFormat(txtMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"				'�ں������� 
		.frm1.txtMnthDeprAmt.value			= "<%=UNINumClientFormat(txtMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'�󰢾� 
		.frm1.txtMnthDisUseQty.value		= "<%=UNINumClientFormat(txtMnthDisUseQty, ggQty.DecPoint, 0)%>"			'�Ű���ⷮ 
		'''''�����	
		.frm1.txtLMnthAcqAmt.value			= "<%=UNINumClientFormat(txtLMnthAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"				'��氡�� 
		.frm1.txtLMnthCptAmt.value			= "<%=UNINumClientFormat(txtLMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"				'�ں������� 
		.frm1.txtLMnthDeprAmt.value			= "<%=UNINumClientFormat(txtLMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'�󰢾� 
		.frm1.txtLMnthBalAmt.value			= "<%=UNINumClientFormat(txtLMnthBalAmt, ggAmtOfMoney.DecPoint, 0)%>"				'�̻󰢾� 
		.frm1.txtLMnthInvQty.value			= "<%=UNINumClientFormat(txtLMnthInvQty, ggQty.DecPoint, 0)%>"				'��� 

		.DbQueryOk
    End If
    .frm1.txtAsstNm.value = "<%=ConvSPChars(Asst_NM)%>"
End With
</Script>
