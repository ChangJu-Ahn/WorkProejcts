<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM112MB1
'*  4. Program Name         : ��Ƽ���۴ϸ���Ȯ��/����-��Ƽ 
'*  5. Program Desc         : ��Ƽ���۴ϸ���Ȯ��/����-��Ƽ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2005/02/28
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : MJG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'Response.Write gStrGlobalCollection
'Response.end
'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim iLngMaxRow		' ���� �׸����� �ִ�Row
Dim iLngRow
Dim GroupCount  
Dim lgCurrency        
Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ����     
Dim lgDataExist
Dim lgPageNo
Dim SoCompanyNm			'�� : ���ֹ��� 
Dim ItemNm			'�� : ǰ��� 
Dim PrTypeNm		'�� : ��û���и� 
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows
intARows=0
intTRows=0
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status


Call HideStatusWnd                                                               '��: Hide Processing message
lgOpModeCRUD  = Request("txtMode") 




Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '��: Query
		 Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '��: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)
        Call SubBizDelete()
End Select

Sub SubBizQueryMulti()

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()		'�� : DB-Agent�� ���� parameter ����Ÿ set
	
	Call QueryData()			'�� : DB-Agent�� ���� ADO query
	
	'-----------------------
	'Result data display area
	'----------------------- 

%>

	<Script Language=vbscript>
		With parent
			.frm1.txtSoCompanyCd.value = "<%=ConvSPChars(Request("txtSoCompanyCd"))%>"			
			.frm1.txtSoCompanyNm.Value	= "<%=SoCompanyNm%>"							
			.frm1.txtSoCompanyCd.focus
			
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then
				
				'Show multi spreadsheet data from this line
				       
				.ggoSpread.Source    = .frm1.vspdData 
				.ggoSpread.SSShowData "<%=istrData%>"                  '��: Display data 
				
				.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
				
				.DbQueryOk <%=intARows%>,<%=intTRows%>
							
			End If
		End with
	</Script>	
<%	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	On Error Resume next
	Err.Clear	

	Dim iRowsData
	Dim iColsData
	Dim iPM8G211
	Dim L_SelectChar
	Dim I3_m_batch_ap_post_wks
	Dim IG1_imp_dtl_group				'��: Protect system from crashing
	Dim pvCB
	Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii
	Dim iErrorPosition
	Dim i
    Dim iCUCount
    
    Const M557_I3_ap_dt_type = 0
    Const M557_I3_ap_dt = 1
    Const M557_I3_import_flg = 2
    
    Const M557_IG1_I1_count = 0
    Const M557_IG1_I2_iv_no = 1
    Const M557_IG1_I3_ap_dt = 2
    
		
	Redim I3_m_batch_ap_post_wks(2)
	             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

             
	Set iPM8G211 = server.CreateObject("PM8G211.cMPostApS")    
 
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM8G211 = Nothing												'��: ComPlus Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	End if
	
	iRowsData = Split(itxtSpread,gRowSep)
	
	I3_m_batch_ap_post_wks(M557_I3_ap_dt_type)		= Trim(Request("hdnApDateFlg"))
	I3_m_batch_ap_post_wks(M557_I3_import_flg)		= Trim(Request("hdnImportFlg"))

	L_SelectChar		= Trim(Request("hdnApFlg"))
	
	pvCB = "F"
	ReDim IG1_imp_dtl_group(ubound(iRowsData) - 1, 2)
	
	For i = 0 To ubound(iRowsData) - 1
		iColsData = Split(iRowsData(i),gColSep)
			
		IG1_imp_dtl_group(i, M557_IG1_I1_count)			=	iColsData(3)	'ROW NO.
		IG1_imp_dtl_group(i, M557_IG1_I2_iv_no)			=	iColsData(1)
		IG1_imp_dtl_group(i, M557_IG1_I3_ap_dt)			=	iColsData(2)
	Next
			
	Call iPM8G211.M_POST_AP_SVR(pvCB,gStrGlobalCollection, L_SelectChar, IG1_imp_dtl_group, I3_m_batch_ap_post_wks, iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "��:","","","","") = True Then
	  	Set iPM8G211 = Nothing
	  	Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'��: ȭ�� ó�� ASP �� ��Ī�� 
		Response.Write "</Script> "
	  	Exit Sub
	End If
		
	Set iPM8G211 = Nothing
                       

    Response.Write "<Script language=vbs> " & vbCr  
    Response.Write " Parent.DbSaveOk "      & vbCr							'��: ȭ�� ó�� ASP �� ��Ī�� 
    Response.Write "</Script> "           
        
End Sub 

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()																			'��: ���� ��û 
	
	On Error Resume Next
    Err.Clear																				'��: Protect system from crashing
	
	Dim iPMMG112,iErrorPosition
	Dim iMaxRow, istrVal
   

	iMaxRow										= Trim(Request("txtMaxRows"))
	istrVal										= Trim(Request("txtSpread"))				   
    
    Set iPMMG112 = Server.CreateObject("PMMG112.cMMaintIvCombi")        
	    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set iPMMG112 = Nothing 		
			Exit Sub
	End If
' 	Call ServerMesgBox(istrVal , vbInformation, I_MKSCRIPT)    	
	    
	
	Call iPMMG112.M_MAINT_IV_COMBI_SVR("F", gStrGlobalCollection, _
									  iMaxRow, _
									  istrVal, _
									  iErrorPosition)								  	    									  

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPMMG112 = Nothing												'��: ComProxy Unload
		Exit Sub															'��: �����Ͻ� ���� ó���� ������ 
	 End If
	 
	 

    Set iPMMG112 = Nothing															'��: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbSaveOk()" & vbCr
	Response.Write "</Script>" & vbCr															'��: Unload Comproxy
												
End Sub
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim PvArr
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt

	Const	M_IV_HDR_CFM_FLG				=	0				'Ȯ������ 
	Const	M_IV_HDR_IV_NO					=	1				'���Թ�ȣ 
	Const	M_IV_HDR_BP_CD					=	2				'���ֹ��� 
	Const	M_IV_HDR_BP_NM					=	3				'���ֹ��θ�                                                  
	Const	M_IV_HDR_SPPL_IV_NO				=	4				'����ó���ݰ�꼭��ȣ                                        
	Const	M_IV_HDR_IV_DT					=	5				'������                                                      
	Const	M_IV_HDR_IV_CUR					=	6				'ȭ��                                                        
	Const	M_IV_HDR_NET_DOC_AMT			=	7				'���ް���                                                    
	Const	M_IV_HDR_TOT_VAT_DOC_AMT		=	8				'�ΰ����ݾ�                                                  
	Const	M_IV_HDR_GROSS_DOC_AMT			=	9				'�հ�ݾ�                                                    
	Const	M_IV_HDR_VAT_TYPE				=	10  			'�ΰ�������                                                  
	Const	M_IV_HDR_VAT_TYPE_NM			=	11  			'�ΰ���������                                                
	Const	M_IV_HDR_VAT_RT					=	12  			'�ΰ�����                                                    
	Const	M_IV_HDR_PAY_METH				=	13  			'�������                                                    
	Const	M_IV_HDR_PAY_METH_NM			=	14  			'��������                                                  
	Const	M_IV_HDR_PAY_TYPE				=	15  			'��������                                                    
	Const	M_IV_HDR_PAY_TYPE_NM			=	16  			'�������Ǹ�                                                  
	Const	M_IV_HDR_PUR_GRP				=	17  			'���ű׷� 
	Const	M_IV_HDR_PUR_GRP_NM				=	18				'���ű׷��                                                    
	Const	M_IV_HDR_TAX_BIZ_AREA			=	19  			'���ݽŰ�����                                              
	Const	M_IV_HDR_GL_NO					=	20  			'��ǥ��ȣ                                                    
	Const	M_IV_HDR_REF_PO_NO				=	21  			'���ֹ�ȣ                                                    
	Const	M_IV_HDR_PO_COMPANY_CD			=	22  			'���ֹ���                                                    
	Const	M_IV_HDR_SO_COMPANY_CD			=	23  			'���ֹ���                     
    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- ���ڵ�� Į�� ���� ----------
	'A.POSTED_FLG, A.IV_NO, A.BP_CD, D.BP_FULL_NM, A.SPPL_IV_NO, 
	'CONVERT(CHAR(10), A.IV_DT, 20) IV_DT, A.IV_CUR,
	'A.NET_DOC_AMT, TOT_VAT_DOC_AMT, GROSS_DOC_AMT, 
	'A.VAT_TYPE, F.MINOR_NM, A.VAT_RT, A.PAY_METH, G.MINOR_NM, A.PAY_TYPE, I.MINOR_NM, 
	'A.PUR_GRP, M.PUR_GRP_NM,
	'A.TAX_BIZ_AREA, A.GL_NO, A.REF_PO_NO, B.PO_COMPANY, B.SO_COMPANY
	'-----------------------------------

	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
	Dim PostedFlg
	
	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		iRowStr = iRowStr & Chr(11) & "0"																			'����    
		If ConvSPChars(rs0(M_IV_HDR_CFM_FLG)) = "Y" Then
			PostedFlg = "Ȯ��"
		Else
			PostedFlg = "��Ȯ��"
		End if
		iRowStr = iRowStr & Chr(11) & PostedFlg																		'Ȯ������ 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_NO))												'���Թ�ȣ 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_BP_CD))												'���ֹ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_BP_NM))												'���ֹ��θ�                                                  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_SPPL_IV_NO))											'����ó���ݰ�꼭��ȣ                                        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_DT))												'������                                                      
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_CUR))												'ȭ��                                                        
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_NET_DOC_AMT), ggAmtOfMoney.DecPoint,0)			'���ް���                                                    
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_TOT_VAT_DOC_AMT), ggAmtOfMoney.DecPoint,0)		'�ΰ����ݾ�                                                  
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_GROSS_DOC_AMT), ggAmtOfMoney.DecPoint,0)		'�հ�ݾ�                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_VAT_TYPE))											'�ΰ�������                                                  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_VAT_TYPE_NM))										'�ΰ���������                                                
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_VAT_RT), ggExchRate.DecPoint,6)					'�ΰ�����                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_METH))											'�������                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_METH_NM))										'��������                                                  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_TYPE))											'��������                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_TYPE_NM))										'�������Ǹ�                                                  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PUR_GRP))											'���ű׷�                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_TAX_BIZ_AREA))										'���ݽŰ�����                                              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_GL_NO))												'��ǥ��ȣ                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_REF_PO_NO))											'���ֹ�ȣ                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PO_COMPANY_CD))										'���ֹ���                                                    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_SO_COMPANY_CD))										'���ֹ���                   
		
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             


		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount-1) = istrData	
		   istrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If
		
		rs0.MoveNext
	Loop
	

	istrData = Join(PvArr, "")

	intARows = iLoopCount
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
	  lgPageNo = ""
	End If
		    
	rs0.Close                                                       '��: Close recordset object
	Set rs0 = Nothing	                                            '��: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    SetConditionData = false
         
    
	If Not(rs1.EOF Or rs1.BOF) Then
		SoCompanyNm = rs1("BP_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSoCompanyCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ֹ���", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    exit function
		End If
	End If   		

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,7)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	strVal = ""
    UNISqlId(0) = "MM112MA101"
    UNISqlId(1) = "MM111MA103"		'���ֹ�����ȸ 

	UNIValue(1,0) = "'zzzzzzzzzz'"            
    
    '���ֹ�����ȸ 
    If Trim(Request("txtSoCompanyCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,0) = "|"
	End If
	
    '��꼭������ 
    If Trim(Request("txtFrBillDt")) <> "" Then
		UNIValue(0,1) =  " '" & Trim(UniConvDate(Request("txtFrBillDt"))) & "' "	
    Else
        UNIValue(0,1) = "|"
	End If
			
    If Trim(Request("txtToBillDt")) <> "" Then
		UNIValue(0,2) =  " '" & Trim(UniConvDate(Request("txtToBillDt"))) & "' "	
    Else
        UNIValue(0,2) = "|"
	End If
	
    '������ 
    If Trim(Request("txtIvFrDt")) <> "" Then
		UNIValue(0,3) =  " '" & Trim(UniConvDate(Request("txtIvFrDt"))) & "' "	
    Else
        UNIValue(0,3) = "|"
	End If
			
    If Trim(Request("txtIvToDt")) <> "" Then
		UNIValue(0,4) =  " '" & Trim(UniConvDate(Request("txtIvToDt"))) & "' "	
    Else
        UNIValue(0,4) = "|"
	End If	

    'Ȯ������ 
    If Trim(Request("rdoCfmflg")) <> "" Then
		UNIValue(0,5) = " '"& FilterVar(Trim(UCase(Request("rdoCfmflg"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,5) = "|"
	End If
		
    '���ֹ�ȣ 
    If Trim(Request("txtCustPoNo")) <> "" Then
		UNIValue(0,6) = " '"& FilterVar(Trim(UCase(Request("txtCustPoNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,6) = "|"
	End If
	
    '����ó���ݰ�꼭��ȣ 
    If Trim(Request("txtBlNo")) <> "" Then
		UNIValue(0,7) = " '"& FilterVar(Trim(UCase(Request("txtBlNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,7) = "|"
	End If			

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)    

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
	If SetConditionData = False Then Exit Sub
	

    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
            Call  MakeSpreadSheetData()

    End If  
End Sub


'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function


%>
