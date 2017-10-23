<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procuremen
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : M51119(Lookup_PO_Hdr)
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2001/10
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Ma Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'* 14. Business Logic of m5141ma1(매입일반등록)
'**********************************************************************************************
	'Dim lgOpModeCRUD
	
	Dim pvCB
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0	
	'Dim lgMaxCount
	Dim lgPageNo
	Dim lgCurrency 
	
	Const C_SHEETMAXROWS_D = 100	

	'Header 부분 
	Const	C_IvNo2			= 0
	Const	C_IvTypeCd		= 1
	Const	C_IvTypeNm		= 2
	Const	C_SpplCd		= 3
	Const	C_SpplNm		= 4
	Const	C_PayeeCd		= 5
	Const	C_PayeeNm		= 6
	Const	C_txtIvAmt		= 7
	Const	C_VatCd			= 8
	Const	C_VatNm			= 9
	Const	C_Vatrt			= 10
	Const	C_GrossVatAmt	= 11
	Const	C_GrpCd			= 12
	Const	C_GrpNm			= 13
	Const	C_Release		= 14
	Const	C_IvDt			= 15
	Const	C_CnfmDt		= 16
	Const	C_BuildCd		= 17
	Const	C_BuildNm		= 18
	Const	C_Cur			= 19
	Const	C_XchRt			= 20
	Const	C_Xchop			= 59
	Const	C_IvLocAmt		= 21
	Const	C_VatFlg1		= 22
	Const	C_GrossVatLocAmt	= 23
	Const	C_BizAreaCd		= 24
	Const	C_BizAreaNm		= 25
	Const	C_SpplRegNo		= 26
	Const	C_PayTermCd		= 27
	Const	C_PayTermNm		= 28
	Const	C_PayTypeCd		= 29
	Const	C_PayTypeNm		= 30
	Const	C_SpplIvNo		= 31
	Const	C_PayDur		= 32
	Const	C_PayTermstxt	= 33
	Const	C_Remark		= 34
	'# 수입여부 추가(2005.04.20)
	Const	C_Importflg		= 69
	
	
	
	'Spread 부분 
    Const   C_IvNo         = 35
    Const   C_IvSeq        = 36
    Const   C_PlantCd      = 37
    Const   C_PlantNm      = 38
    Const   C_ItemCd       = 39
    Const   C_ItemNm       = 40
    Const   C_Spec         = 41
    Const   C_IvQty        = 42
    Const   C_CtlQty       = 43
    Const   C_IvPrc        = 44
    Const   C_CtlPrc       = 45
    Const   C_Amt          = 46
    Const   C_CtlAmt       = 47
    Const   C_VatFlg       = 48
    Const   C_VatAmt       = 49
    Const   C_VatCtlAmt    = 50
    Const   C_LocAmt		= 51
    Const   C_CtlLocAmt		= 52
    Const   C_LocVatAmt		= 53
    Const   C_CtlLocVatAmt	= 54
    Const   C_PoNo			= 55
    Const   C_PoSeq			= 56
    Const   C_IvNohdn		= 57
    Const   C_IvSeqhdn		= 58
    
	Const	C_ItemAcct      = 61
	Const	C_VatType       = 62
	Const	C_MvmtNo        = 63
	Const	C_IvCostCd      = 64
	Const	C_IvBizArea     = 65
	Const	C_MvmtQty       = 66
	Const	C_MvmtFlg       = 67
	Const	C_ItemUnit		= 68

	    	
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	lgOpModeCRUD	=	Request("txtMode")           '☜: Read Operation Mode (CRUD)

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0001)                                                         '☜: Query
	             Call SubBizQuery()
	        Case CStr(UID_M0002)
	             Call SubBizSaveMulti()
	        Case CStr(UID_M0003)                                                         '☜: Delete
	             Call SubBizDelete()
	        Case "Release", "UnRelease"				
				 Call SubReleaseCheck()
	        Case "LookUpSupplier"                                                                 '☜: Check	
	             Call SubLookUpSupplier()    
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
    Err.Clear 

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,2)
    
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    strval = ""
    
    If len(Request("txtIvNo")) Then
		strVal = strVal & " AND A.IV_NO = " & FilterVar(Trim(UCase(Request("txtIvNo"))), " " , "S") & " "		
	End If
	
	If len(Request("lgNextKey")) Then
		strVal = strVal & " AND J.IV_SEQ_NO >= " & FilterVar(Trim(UCase(Request("lgNextKey"))), " " , "S") & " "		
	End If 
		
    UNISqlId(0) = "M5141MA1_1"
	UNIValue(0,1) = strVal 
	UNIVALUE(0,2) = "ORDER BY J.IV_SEQ_NO"
	
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
    iStr = Split(lgstrRetMsg,gColSep)



	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    If  rs0.EOF And rs0.BOF  Then
		      
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else
		Call  MakeHdrData()   	
		Call  MakeSpreadSheetData()
    End If  
    
    Set lgADF   = Nothing
	
	
End Sub

'============================================================================================================
' Name : MakeHdrData
' Desc : Set Data in Header Area
'============================================================================================================
Sub MakeHdrData()

	Dim strDefDate
	
	lgCurrency = ConvSPChars(Trim(rs0(19)))
	strDefDate = UniDateClientFormat("1899-12-31")
	
	Response.Write "<Script Language=vbscript>"															& vbCr
	Response.Write "With parent.frm1"																	& vbCr
	Response.Write "	parent.CurFormatNumericOCX	" 													& vbCr

	'첫번째 탭 
	Response.Write "	.txtIvNo2.value	= """ & ConvSPChars(Trim(rs0(C_IvNo2)))	& """" & vbCr
	Response.Write "	.txtIvTypeCd.value	= """ & ConvSPChars(Trim(rs0(C_IvTypeCd)))	& """" & vbCr
	Response.Write "	.txtIvTypeNm.value	= """ & ConvSPChars(Trim(rs0(C_IvTypeNm)))	& """" & vbCr
	Response.Write "	.txtSpplCd.value	= """ & ConvSPChars(Trim(rs0(C_SpplCd)))	& """" & vbCr
	Response.Write "	.txtSpplNm.value	= """ & ConvSPChars(Trim(rs0(C_SpplNm)))	& """" & vbCr
	Response.Write "	.txtPayeeCd.value	= """ & ConvSPChars(Trim(rs0(C_PayeeCd)))	& """" & vbCr
	Response.Write "	.txtPayeeNm.value	= """ & ConvSPChars(Trim(rs0(C_PayeeNm)))	& """" & vbCr
	Response.Write "	.txtIvAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_txtIvAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.hdnIvAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_txtIvAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtVatCd.value	= """ & ConvSPChars(Trim(rs0(C_VatCd)))	& """" & vbCr
	Response.Write "	.txtVatNm.value	= """ & ConvSPChars(Trim(rs0(C_VatNm)))	& """" & vbCr
	Response.Write "	.txtVatRt.value	= """ & UNINumClientFormat(Trim(rs0(C_Vatrt)),ggExchRate.DecPoint,0) & """"	& vbCr		
	Response.Write "	.txtGrossVatAmt.value	= """ & UNINumClientFormat(Trim(rs0(C_GrossVatAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.hdnGrossVatAmt.value	= """ & UNINumClientFormat(Trim(rs0(C_GrossVatAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtGrpCd.value	= """ & ConvSPChars(Trim(rs0(C_GrpCd)))	& """" & vbCr
	Response.Write "	.txtGrpNm.value	= """ & ConvSPChars(Trim(rs0(C_GrpNm)))	& """" & vbCr
	
	If ConvSPChars(Trim(rs0(C_Release))) = "N" Then '확정여부 POSTED_FLG
		Response.Write "	.rdoRelease(0).Checked= true" 	& vbCr
		Response.Write "	.hdnRelease.value=""N""" 		& vbCr
	Else
		Response.Write "	.rdoRelease(1).Checked= true" 	& vbCr
		Response.Write "	.hdnRelease.value=""Y""" 		& vbCr
	End If		
	
	Response.Write "	.txtIvDt.Text	= """ & UNIDateClientFormat(Trim(rs0(C_IvDt))) 	& """"	& vbCr				
	Response.Write "	.txtCnfmDt.Text	= """ & UNIDateClientFormat(Trim(rs0(C_CnfmDt))) 	& """"	& vbCr				
	Response.Write "	.txtBuildCd.value	= """ & ConvSPChars(Trim(rs0(C_BuildCd)))	& """" & vbCr
	Response.Write "	.txtBuildNm.value	= """ & ConvSPChars(Trim(rs0(C_BuildNm)))	& """" & vbCr
	Response.Write "	.txtCur.value	= """ & ConvSPChars(Trim(rs0(C_Cur)))	& """" & vbCr
	Response.Write "	.txtXchRt.value	= """ & UNINumClientFormat(Trim(rs0(C_XchRt)),ggExchRate.DecPoint,0) & """"	& vbCr		
	Response.Write "	.cboXchop.value				= """ & ConvSPChars(Trim(rs0(C_Xchop)))			& """" & vbCr 'Multi Divide
	Response.Write "	.txtIvLocAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_IvLocAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.hdnIvLocAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_IvLocAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	
	If ConvSPChars(Trim(rs0(C_VatFlg1))) = "2" Then 'VAT 포함여부 VAT_INC_FLAG
		Response.Write "	.rdoVatFlg2.Checked= true" 	& vbCr
		Response.Write "	.hdnVatFlg.value=""Y""" 		& vbCr
	Else
		Response.Write "	.rdoVatFlg1.Checked= true" 	& vbCr
		Response.Write "	.hdnVatFlg.value=""N"""  		& vbCr
	End If	
	
	Response.Write "	.txtGrossVatLocAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_GrossVatLocAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.hdnGrossVatLocAmt.value		= """ & UNINumClientFormat(Trim(rs0(C_GrossVatLocAmt)),ggAmtOfMoney.DecPoint,0) & """"	& vbCr
	Response.Write "	.txtBizAreaCd.value	= """ & ConvSPChars(Trim(rs0(C_BizAreaCd)))	& """" & vbCr
	Response.Write "	.txtBizAreaNm.value	= """ & ConvSPChars(Trim(rs0(C_BizAreaNm)))	& """" & vbCr
	
	
	'두번째 탭 
	Response.Write "	.txtSpplRegNo.value	= """ & ConvSPChars(Trim(rs0(C_SpplRegNo)))	& """" & vbCr
	Response.Write "	.txtPayTermCd.value	= """ & ConvSPChars(Trim(rs0(C_PayTermCd)))	& """" & vbCr
	Response.Write "	.txtPayTermNm.value	= """ & ConvSPChars(Trim(rs0(C_PayTermNm)))	& """" & vbCr
	Response.Write "	.txtPayTypeCd.value	= """ & ConvSPChars(Trim(rs0(C_PayTypeCd)))	& """" & vbCr
	Response.Write "	.txtPayTypeNm.value	= """ & ConvSPChars(Trim(rs0(C_PayTypeNm)))	& """" & vbCr
	Response.Write "	.txtSpplIvNo.value	= """ & ConvSPChars(Trim(rs0(C_SpplIvNo)))	& """" & vbCr
	Response.Write "	.txtPayDur.value	= """ & ConvSPChars(Trim(rs0(C_PayDur)))	& """" & vbCr
	Response.Write "	.txtPayTermsTxt.value	= """ & ConvSPChars(Trim(rs0(C_PayTermstxt)))	& """" & vbCr
	Response.Write "	.txtRemark.value	= """ & ConvSPChars(Trim(rs0(C_Remark)))	& """" & vbCr
	
	Response.Write "	.hdnIvNo.value	= """ & ConvSPChars(Trim(rs0(C_IvNo2)))	& """" & vbCr
	Response.Write "	.hdnImportflg.Value 	= """ & ConvSPChars(Trim(rs0(C_Importflg))) & """" & vbCr
	
		
	Response.Write "End With" 	& vbCr
	Response.Write "</Script>" 	& vbCr

End Sub

'============================================================================================================
' Name : MakeHdrData
' Desc : Set Data in Header Area
'============================================================================================================
Sub MakeSpreadSheetData

	Dim iLngMaxRow
	Dim iMax
	Dim PvArr
	Dim iLngRow
	Dim StrNextKey
	Dim istrData
	
	Dim TmpAmt
	Dim TmpVatAmt
	
	Dim Index
	


   	iLngMaxRow = CLng(Request("txtMaxRows"))
	
	iMax = rs0.RecordCount
	ReDim PvArr(iMax)

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write  "	.vspdData.Redraw = False   "                         & vbCr    	
	
	 

	For iLngRow = 0 To iMax-1
	
        If iLngRow >= C_SHEETMAXROWS_D Then
			StrNextKey = ConvSPChars(Trim(rs0(C_IvSeq)))
			
			Exit For
        End If	

		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvNo))	'1
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvSeq))	'2
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_PlantCd))	'3
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_PlantNm))	'4
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_ItemCd))	'5
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_ItemNm))	'6
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_Spec))	'7
		istrData = istrData & Chr(11) & UNINumClientFormat(rs0(C_IvQty),ggQty.DecPoint,0) '8	       
		istrData = istrData & Chr(11) & UNINumClientFormat(rs0(C_CtlQty),ggQty.DecPoint,0) '9	     
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_IvPrc),lgCurrency,ggUnitCostNo, "X" , "X")		'10	
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlPrc),lgCurrency,ggUnitCostNo, "X" , "X")		'11
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_Amt),lgCurrency,ggUnitCostNo, "X" , "X")		'12		

		If Trim(rs0(C_VatFlg)) = "2" then		'포함 
			TmpAmt = Cdbl(rs0(C_CtlAmt)) + Cdbl(rs0(C_VatCtlAmt))
		Else
			TmpAmt = rs0(C_CtlAmt)
		End If

		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(TmpAmt,lgCurrency,ggUnitCostNo, "X" , "X")		'12		
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'13		
		
		If Trim(rs0(C_VatFlg)) = "2" then								'14
			istrData = istrData & Chr(11) & "포함"
		Else
			istrData = istrData & Chr(11) & "별도"
		End If
		
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_VatFlg))	'15

		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_VatAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'16		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_VatCtlAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'17
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_LocAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'17

		If Trim(rs0(C_VatFlg)) = "2" then		'포함 
			TmpVatAmt = Cdbl(rs0(C_CtlLocAmt)) + Cdbl(rs0(C_CtlLocVatAmt))
		Else
			TmpVatAmt = rs0(C_CtlLocAmt)
		End If		
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(TmpAmt,lgCurrency,ggUnitCostNo, "X" , "X")		'18				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlLocAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'18				
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_LocVatAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'17
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlLocVatAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_PoNo))	'20
		
		If rs0(C_PoNo) = "" Then
			istrData = istrData & Chr(11) & ""
		Else 
			istrData = istrData & Chr(11) & ConvSPChars(rs0(C_PoSeq))	'21
		End If
		
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvNohdn))	'22
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvSeqhdn))	'23
		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlQty),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlPrc),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_VatCtlAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlLocAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlLocVatAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'19				
		

		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'12		
		

		
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(C_CtlLocAmt),lgCurrency,ggUnitCostNo, "X" , "X")		'18	
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_ItemAcct))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_VatType))
		istrData = istrData & Chr(11) & UNINumClientFormat(rs0(C_VatRt), ggExchRate.DecPoint,6)
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_MvmtNo))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvCostCd))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvBizArea))
		istrData = istrData & Chr(11) & UNINumClientFormat(rs0(C_MvmtQty),ggQty.DecPoint,0) 
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_MvmtFlg))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvNohdn))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_IvSeqhdn))
		istrData = istrData & Chr(11) & ConvSPChars(rs0(C_ItemUnit))					
		
		istrData = istrData & Chr(11) & ConvSPChars("Q")	'23
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)  
		PvArr(iLngRow) = istrData
		istrData=""		
		
		if not(rs0.EOF) and not(rs0.BOF) then
			rs0.movenext
		end if 
								
	Next
		
	Response.Write "End With"		& vbCr
	Response.Write "</Script>"		& vbCr 	
	
	istrData = Join(PvArr, "")
	
	Response.Write "<Script Language=vbscript>"															& vbCr
	Response.Write "With parent"		

    For index = 1 to iMax
		Response.Write "			        .ggoSpread.SpreadLock index , -1"		& vbCr
	Next
    	
   Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
   Response.Write "	.lgStrPrevKey           = """ & StrNextKey   & """" & vbCr  


	
   
   Response.Write  "	.frm1.vspdData.Redraw = True   "     & vbCr  
   Response.Write " .DbQueryOk "	& vbCr 
   Response.Write "End With"		& vbCr
   Response.Write "</Script>"		& vbCr    
	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSaveMulti()		'☜: 저장 요청을 받음 
	
	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear 
    
	Dim iPM8G311
    Dim iErrorPosition
    Dim iUpdtUserId, ihdnIvNo, itxtSpread
    '-------------------
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
    
	Dim lgIntFlgMode
	Dim iCommandSent
	
	    
    Dim I1_m_iv_type_iv_type_cd
    Dim I2_b_biz_partner_bp_cd
    Dim I3_b_pur_grp
    Dim I4_b_company
    Dim I5_m_iv_hdr
    Dim I6_m_user_id
    Dim l7_m_iv_post_flg
    Dim l8_m_iv_hdr
    Dim E1_Return

'    I1_m_iv_type_iv_type_cd    				= Trim(UCase(Request("txtIvTypeCd")))
'    I2_b_biz_partner_bp_cd					= Trim(UCase(Request("txtSpplCd")))
'    I3_b_pur_grp							= Trim(UCase(Request("txtGrpCd")))
'    I4_b_company							= gCurrency    
'    I6_m_user_id							= Trim(Request("hdnUsrId"))        

    Redim I5_m_iv_hdr(59)    
    
    Const M528_I4_iv_dt = 0    '  View Name : imp m_iv_hdr
    Const M528_I4_pay_dt = 1
    Const M528_I4_iv_cur = 2
    Const M528_I4_xch_rt = 3
    Const M528_I4_vat_type = 4
    Const M528_I4_pay_meth = 5
    Const M528_I4_pay_dur = 6
    Const M528_I4_vat_rt = 7
    Const M528_I4_remark = 8
    Const M528_I4_gross_doc_amt = 9
    Const M528_I4_tot_vat_doc_amt = 10
    Const M528_I4_sppl_iv_no = 11
    Const M528_I4_ap_post_dt = 12
    Const M528_I4_iv_no = 13
    Const M528_I4_posted_flg = 14
    Const M528_I4_payee_cd = 15
    Const M528_I4_build_cd = 16
    Const M528_I4_pur_org = 17
    Const M528_I4_iv_biz_area = 18
    Const M528_I4_tax_biz_area = 19
    Const M528_I4_iv_cost_cd = 20
    Const M528_I4_pay_terms_txt = 21
    Const M528_I4_pay_type = 22
    Const M528_I4_gross_loc_amt = 23
    Const M528_I4_net_doc_amt = 24
    Const M528_I4_net_loc_amt = 25
    Const M528_I4_cash_doc_amt = 26
    Const M528_I4_cash_loc_amt = 27
    Const M528_I4_tot_vat_loc_amt = 28
    Const M528_I4_tot_diff_doc_amt = 29
    Const M528_I4_tot_diff_loc_amt = 30
    Const M528_I4_pay_bank_cd = 31
    Const M528_I4_pay_acct_cd = 32
    Const M528_I4_pp_no = 33
    Const M528_I4_pp_doc_amt = 34
    Const M528_I4_pp_loc_amt = 35
    Const M528_I4_loan_no = 36
    Const M528_I4_loan_doc_amt = 37
    Const M528_I4_loan_loc_amt = 38
    Const M528_I4_bl_no = 39
    Const M528_I4_bl_doc_no = 40
    Const M528_I4_lc_doc_no = 41
    Const M528_I4_ref_po_no = 42
    Const M528_I4_ext1_cd = 43
    Const M528_I4_ext1_qty = 44
    Const M528_I4_ext1_amt = 45
    Const M528_I4_ext1_rt = 46
    Const M528_I4_ext1_dt = 47
    Const M528_I4_ext2_cd = 48
    Const M528_I4_ext2_qty = 49
    Const M528_I4_ext2_amt = 50
    Const M528_I4_ext2_rt = 51
    Const M528_I4_ext2_dt = 52
    Const M528_I4_ext3_cd = 53
    Const M528_I4_ext3_qty = 54
    Const M528_I4_ext3_amt = 55
    Const M528_I4_ext3_rt = 56
    Const M528_I4_ext3_dt = 57
    Const M528_I4_vat_inc_flag = 58
    Const M528_I4_xch_rate_op = 59        
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별     
    
    l8_m_iv_hdr = Trim(UCase(Request("txtIvNo2")))
    
    I1_m_iv_type_iv_type_cd    				= Trim(UCase(Request("txtIvTypeCd")))	'매입유형 
    I5_m_iv_hdr(M528_I4_iv_dt)				= UniConvDate(Request("txtIvDt"))		'매입일 
    
    '지불예정일을 입력하지 않은 경우 2999/12/31로 셋팅함(2003.09.22)
'    If Trim(Request("txtPayDt")) = "" Then
		I5_m_iv_hdr(M528_I4_pay_dt)			= "2999-12-31"
'    Else
'		I5_m_iv_hdr(M528_I4_pay_dt)			= UniConvDate(Request("txtPayDt"))
'    End If
	I5_m_iv_hdr(M528_I4_ap_post_dt)			= UniConvDate(Request("txtCnfmDt"))		'확정일 
	I5_m_iv_hdr(M528_I4_posted_flg)			= Trim(UCase(Request("hdnRelease")))	'확정여부 
	
	l7_m_iv_post_flg						= Trim(UCase(Request("hdnRelease")))	'확정여부 
	
    I5_m_iv_hdr(M528_I4_vat_inc_flag)		= Trim(UCase(Request("hdvatFlg")))		'vat 포함 구분 
 
    I2_b_biz_partner_bp_cd					= Trim(UCase(Request("txtSpplCd")))		'공급처 
    I5_m_iv_hdr(M528_I4_payee_cd)			= Trim(UCase(Request("txtPayeeCd")))	'지급처 
    I5_m_iv_hdr(M528_I4_build_cd)			= Trim(UCase(Request("txtBuildCd")))	'세금계산서발행처 
    
    I5_m_iv_hdr(M528_I4_sppl_iv_no)			= Trim(UCase(Request("txtSpplIvNo")))	'공급처계산서번호(Invoice No.)
    I3_b_pur_grp							= Trim(UCase(Request("txtGrpCd")))		'구매그룹 
    I5_m_iv_hdr(M528_I4_iv_cur)				= Trim(UCase(Request("txtCur")))		'화폐 
    I5_m_iv_hdr(M528_I4_vat_type)			= Trim(UCase(Request("txtVatCd")))		'VAT Type

    if Trim(Request("txtVatRt")) <> "" then											'VAT RT
		I5_m_iv_hdr(M528_I4_vat_rt)			= UniConvNum(Request("txtVatRt"),0)
	else
		I5_m_iv_hdr(M528_I4_vat_rt)			= "0"
	End if 

    I5_m_iv_hdr(M528_I4_pay_meth)				= Trim(UCase(Request("txtPayTermCd")))	'결제방법 
    if Trim(Request("txtPayDur")) <> "" then											'결제기간 
		I5_m_iv_hdr(M528_I4_pay_dur)			= UniConvNum(Request("txtPayDur"),0)	
	else
		I5_m_iv_hdr(M528_I4_pay_dur)			= "0"
	end if
    I5_m_iv_hdr(M528_I4_pay_type)				= Trim(UCase(Request("txtPayTypeCd")))	'지급유형 
    
    I5_m_iv_hdr(M528_I4_net_doc_amt)			= Request("txtIvAmt")					'매입순금액 
    I5_m_iv_hdr(M528_I4_net_loc_amt)			= Request("txtIvLocAmt")				'매입원화금액 
    
    I5_m_iv_hdr(M528_I4_tot_vat_doc_amt)		= Request("txtGrossVatAmt")				'VAT 총금액 
    I5_m_iv_hdr(M528_I4_tot_vat_loc_amt)		= Request("txtGrossVatLocAmt")			'VAT 총원화금액 
    
    if Trim(Request("txtXchRt")) <> "" then												'환율 
		I5_m_iv_hdr(M528_I4_xch_rt)				= UniConvNum(Request("txtXchRt"),0)
	else
		I5_m_iv_hdr(M528_I4_xch_rt)				= "0"
	end if
    I5_m_iv_hdr(M528_I4_pay_terms_txt)			= Trim(Request("txtPayTermsTxt"))		'대금결제참조 
    I5_m_iv_hdr(M528_I4_tax_biz_area)			= Trim(UCase(Request("txtBizAreaCd")))	'세금신고사업장 
    I5_m_iv_hdr(M528_I4_remark)					= Trim(Request("txtRemark"))			'비고 
    '추가 
    I5_m_iv_hdr(M528_I4_xch_rate_op)			= Trim(Request("hdnxchrateop"))			'환율적용시 Operation
    
    
    I4_b_company						= gCurrency										'화폐 
    I6_m_user_id			= Trim(Request("txtUpdtUserId"))									
    
'    If Request("txtChkPoNo") = "Y" Then 
'		 I5_m_iv_hdr(M528_I4_ref_po_no) = Trim(Request("txtPoNo"))
'	Else
		 I5_m_iv_hdr(M528_I4_ref_po_no) = ""
'	End if
	    
    If lgIntFlgMode = OPMD_CMODE Then
		I5_m_iv_hdr(M528_I4_iv_no)			= Trim(UCase(Request("hdnIvNo")))
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		I5_m_iv_hdr(M528_I4_iv_no)			= Trim(Request("txtIvNo2"))
		iCommandSent = "UPDATE"
    End If             
    
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count

    
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")
   
    
    'Call ServerMesgBox(itxtSpread , vbInformation, I_MKSCRIPT)
   
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
		
		       
	Set iPM8G311 = CreateObject("PM8G311.cMSettleIV")   

	If CheckSYSTEMError(Err,True) = true then 	
		Set iPM8G311 = Nothing
		Exit Sub
	End If
	
	iUpdtUserId = gUsrID
	ihdnIvNo = Trim(Request("txthdnIvNo"))
	
	pvCB = "F"

    If lgIntFlgMode = OPMD_CMODE Then
		E1_Return = iPM8G311.M_MAINT_SETTLE_IV(pvCB, gStrGlobalCollection, Cstr(iCommandSent), _
													Cstr(I1_m_iv_type_iv_type_cd), _
													Cstr(I2_b_biz_partner_bp_cd), _
													Cstr(I3_b_pur_grp), _
													Cstr(I4_b_company), _
													I5_m_iv_hdr, Cstr(I6_m_user_id), _
													Cstr(l7_m_iv_post_flg), _
													Cstr(l8_m_iv_hdr), _													
													itxtSpread)
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		E1_Return = iPM8G311.M_MAINT_SETTLE_IV(pvCB, gStrGlobalCollection, Cstr(iCommandSent), _
													Cstr(I1_m_iv_type_iv_type_cd), _
													Cstr(I2_b_biz_partner_bp_cd), _
													Cstr(I3_b_pur_grp), _
													Cstr(I4_b_company), _
													I5_m_iv_hdr, Cstr(I6_m_user_id), _
													Cstr(l7_m_iv_post_flg), _
													Cstr(l8_m_iv_hdr), _
													itxtSpread)    
    
    End If

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set iPM8G311 = Nothing
		Response.Write "<Script language=vbs> " & vbCr  
		Response.Write " Parent.RemovedivTextArea "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script> "
		Exit Sub
	End If

    Set iPM8G311 = Nothing    

'	Response.Write "<Script Language=VBScript>" & vbCr
'	Response.Write "With parent" & vbCr
'	Response.Write "If """ & lgIntFlgMode & """  = """ & OPMD_CMODE & """  Then " & vbCr
'	Response.Write "  .frm1.txtIvNo.value	= """ & ConvSPChars(E1_Return) & """" & vbCr
'	Response.Write "  .frm1.hdnIvNo.value	= """ & ConvSPChars(E1_Return) & """" & vbCr
'	Response.Write "End If" & vbCr
'	Response.Write " .DbSaveOk" & vbCr
'	Response.Write "End With" & vbCr
'	Response.Write "</Script>" & vbCr
	
	                  
    Response.Write "<Script language=vbs> " & vbCr         
'    Response.Write " Parent.frm1.txtIvNo.Value = """ & ConvSPChars(Request("hdnIvNo")) & """" & vbCr    
	Response.Write " Parent.frm1.txtIvNo.Value  = """ & ConvSPChars(E1_Return) & """" & vbCr 
	Response.Write " Parent.frm1.hdnIvNo.value	= """ & ConvSPChars(E1_Return) & """" & vbCr
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "
End Sub

'============================================================================================================
' Name : SubReleaseCheck
' Desc : 
'============================================================================================================
Sub SubReleaseCheck()																'☜: 회계처리,회계처리취소 요청을 받음	

    On Error Resume Next
    Err.Clear                                                                       '☜: Protect system from crashing

    Dim I2_ief_supplied
	Dim l3_m_iv_hdr
	Dim I4_ap_dt
	Dim l5_import_flg
	Dim l6_m_iv_type

    Dim PM8G312
    Dim lgIntFlgMode

    If Len(Trim(Request("txtCnfmDt"))) Then
        If UNIConvDate(Request("txtCnfmDt")) = "" Then
            Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
            Call LoadTab("parent.frm1.txtCnfmDt", 0, I_MKSCRIPT)
            Exit Sub
        End If
    End If

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                      '☜: 저장시 Create/Update 판별  
    
	
    If Trim(Request("hdnRelease")) = "Y" Then
        I2_ief_supplied = "N"
    Else
        I2_ief_supplied = "Y"
    End If
    
    l3_m_iv_hdr		=	Trim(Request("txtIvNo2"))			'매입번호 
    I4_ap_dt		=	UniConvDate(Request("txtCnfmDt"))	'확정일 
    l5_import_flg	=	Trim(Request("hdnImportflg"))
    l6_m_iv_type	=	Trim(Request("txtIvTypeCd"))

    Set PM8G312 = Server.CreateObject("PM8G312.cMSettlePostAP")

   If CheckSYSTEMError(Err, True) = True Then
        Set PM8G312 = Nothing
		Exit Sub
    End If
    
	pvCB = "F"
	
    Call PM8G312.M_SETTLE_POST_AP(pvCB, gStrGlobalCollection,I2_ief_supplied, l3_m_iv_hdr, I4_ap_dt, l5_import_flg, l6_m_iv_type)

    If CheckSYSTEMError(Err, True) = True Then
        Set PM8G312 = Nothing
        	Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.btnCfm.disabled = False" & vbCr
			Response.Write "</Script>" & vbCr
        Exit Sub
    End If
    '-----------------------
    'Result data display area
    '-----------------------
'    Response.Write "<Script Language=VBScript>" & vbCr
'    Response.Write "With parent" & vbCr
'    Response.Write " .MainQuery()" & vbCr
'    Response.Write "End With" & vbCr
'    Response.Write "</Script>" & vbCr
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.frm1.txtIvNo.Value = """ & ConvSPChars(Request("hdnIvNo")) & """" & vbCr    
    Response.Write " Parent.MainQuery() "      & vbCr   
    Response.Write "</Script> "    

    Set PM8G312 = Nothing                                                   '☜: Unload Comproxy
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()
	Dim iPM8G311	
    Dim	l7_m_iv_post_flg
    Dim l8_m_iv_hdr
    Dim iCommandSent, itxtSpread
    
	On Error Resume Next
	Err.Clear 


	

	Set iPM8G311 = Server.CreateObject("PM8G311.cMSettleIV")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPM8G311 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     

    l7_m_iv_post_flg						= Trim(UCase(Request("hdnRelease")))	'확정여부 
    l8_m_iv_hdr								= Trim(Request("txtIvNo"))				'매입번호 

	pvCB = "F"	
	iCommandSent = "DELETE"		
	itxtSpread = ""
	
	Call iPM8G311.M_MAINT_SETTLE_IV(pvCB, gStrGlobalCollection, Cstr(iCommandSent), _
												, _
												, _
												, _
												, _
												, , _
												Cstr(l7_m_iv_post_flg), _
												Cstr(l8_m_iv_hdr), _													
												itxtSpread)
    
	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iPM8G311 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iPM8G311 = Nothing

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "Call parent.DbDeleteOk()" & vbCr
	Response.Write "</Script>" & vbCr
	        
End Sub
'============================================================================================================
' Name : SubLookUpSupplier
' Desc : 
'============================================================================================================
Sub SubLookUpSupplier()
	Dim BpType, BpCd
	Dim iPB5CS41
	Dim E1_b_biz_partner
	
	Const S074_E1_bp_rgst_no = 2
	Const S074_E1_currency = 17
	Const S074_E1_pay_meth = 29
	Const S074_E1_pay_dur = 30
	Const S074_E1_vat_type = 33
	Const S074_E1_pay_type = 45
	Const S074_E1_pay_terms_txt = 46
	Const S074_E1_vat_type_nm = 124                           '[부가세유형명]
	Const S074_E1_pay_meth_nm = 133       
	Const S074_E1_pay_type_nm = 134 
	'추가(구매용)
	Const S074_E1_pay_meth_pur = 115                          '결재방법(구매)
    Const S074_E1_pay_type_pur = 116                          '입출금유형(구매)
    Const S074_E1_pay_dur_pur = 117                           '결재기간(구매)
    '네임추가 
    Const S074_E1_pay_meth_pur_nm = 141                       '[결재방법명(구매)]
    Const S074_E1_pay_type_pur_nm = 142                       '[입출금유형명(구매)]

    On Error Resume Next
    Err.Clear
    
    Set iPB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If

    BpType = Trim(Request("txtBpType"))
    BpCd = Trim(Request("txtBpCd"))
    
    Call iPB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection,"QUERY",BpCd,E1_b_biz_partner) 
    
    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If
    
    Set iPB5CS41 = Nothing     
    
    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "With parent.frm1" & vbCr
    Response.Write "If  """ & BpType & """ = ""1"" then " & vbCr       '공급처경우 
    Response.Write "    .txtCur.Value             = """   & ConvSPChars(E1_b_biz_partner(S074_E1_currency))    & """" & vbCr
    Response.Write "    .txtCurNm.Value           = """"" & vbCr
    Response.Write " parent.GetPayDt()"   & vbCr
    Response.Write " parent.ChangeCurr()" & vbCr
	'***2002.12월 패치********
    Response.Write "ElseIf """ & BpType & """ = ""2"" then" & vbCr  '지급처인 경우 
    '이성룡 주석 
    'Response.Write "  If .ChkPoNo.checked = False  then " & vbCr
    'Response.Write "    .txtPayMethCd.Value       = """   & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth_pur))    & """" & vbCr
    'Response.Write "    .txtPayMethNm.Value       = """	  & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth_pur_nm))      & """" & vbCr
    'Response.Write "  End If"             & vbCr

    Response.Write "  .txtPayDur.Value            = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_dur_pur))       & """" & vbCr
    Response.Write "  .txtPayTermstxt.Value       = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_terms_txt)) & """" & vbCr
    Response.Write "  .txtPayTypeCd.Value         = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur))      & """" & vbCr
    Response.Write "  .txtPayTypeNm.Value         = """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur_nm))      & """" & vbCr
    Response.Write "ElseIf """ & BpType & """  = ""3""  then" & vbCr  '세금계산서발행처인 경우 
    Response.Write "  .txtVatCd.Value             = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type))      & """" & vbCr
    Response.Write "  .txtVatNm.Value             = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type_nm))   & """" & vbCr
    'Response.Write "  .txtSpplRegNo.Value        = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_rgst_no))    & """" & vbCr
    Response.Write " parent.SetVatType()" & vbCr
    Response.Write "End If"               & vbCr
    Response.Write "End With"             & vbCr
    Response.Write "</Script>"            & vbCr
                                                           '☜: Process End
End Sub

'============================================================================================================
' Name : SubLookUpPo
' Desc :
'============================================================================================================
Sub SubLookUpPo()
	Const M239_E22_release_flg = 0
    Const M239_E22_merg_pur_flg = 1
    Const M239_E22_po_no = 2
    Const M239_E22_po_dt = 3
    Const M239_E22_xch_rt = 4
    Const M239_E22_vat_type = 5
    Const M239_E22_vat_rt = 6
    Const M239_E22_tot_vat_doc_amt = 7
    Const M239_E22_tot_po_doc_amt = 8
    Const M239_E22_tot_po_loc_amt = 9
    Const M239_E22_pay_meth = 10
    Const M239_E22_pay_dur = 11
    Const M239_E22_pay_terms_txt = 12
    Const M239_E22_pay_type = 13
    Const M239_E22_sppl_sales_prsn = 14
    Const M239_E22_sppl_tel_no = 15
    Const M239_E22_remark = 16
    Const M239_E22_vat_inc_flag = 17
    Const M239_E22_offer_dt = 18
    Const M239_E22_fore_dvry_dt = 19
    Const M239_E22_expiry_dt = 20
    Const M239_E22_invoice_no = 21
    Const M239_E22_incoterms = 22
    Const M239_E22_transport = 23
    Const M239_E22_sending_bank = 24
    Const M239_E22_delivery_plce = 25
    Const M239_E22_applicant = 26
    Const M239_E22_manufacturer = 27
    Const M239_E22_agent = 28
    Const M239_E22_origin = 29
    Const M239_E22_packing_cond = 30
    Const M239_E22_inspect_means = 31
    Const M239_E22_dischge_city = 32
    Const M239_E22_dischge_port = 33
    Const M239_E22_loading_port = 34
    Const M239_E22_shipment = 35
    Const M239_E22_import_flg = 36
    Const M239_E22_bl_flg = 37
    Const M239_E22_cc_flg = 38
    Const M239_E22_rcpt_flg = 39
    Const M239_E22_subcontra_flg = 40
    Const M239_E22_ret_flg = 41
    Const M239_E22_iv_flg = 42
    Const M239_E22_rcpt_type = 43
    Const M239_E22_issue_type = 44
    Const M239_E22_iv_type = 45
    Const M239_E22_po_cur = 46
    Const M239_E22_xch_rate_op = 47
    Const M239_E22_pur_org = 48
    Const M239_E22_pur_biz_area = 49
    Const M239_E22_pur_cost_cd = 50
    Const M239_E22_tot_vat_loc_amt = 51
    Const M239_E22_cls_flg = 52
    Const M239_E22_lc_flg = 53
    Const M239_E22_sppl_cd = 54
    Const M239_E22_payee_cd = 55
    Const M239_E22_build_cd = 56
    Const M239_E22_charge_flg = 55
    Const M239_E22_tracking_no = 56
    Const M239_E22_so_no = 57
    Const M239_E22_inspect_method = 58
    Const M239_E22_ref_no = 59
    Const M239_E22_ext1_cd = 60
    Const M239_E22_ext1_qty = 61
    Const M239_E22_ext1_amt = 62
    Const M239_E22_ext1_rt = 63
    Const M239_E22_ext1_dt = 64
    Const M239_E22_ext2_cd = 65
    Const M239_E22_ext2_qty = 66
    Const M239_E22_ext2_amt = 67
    Const M239_E22_ext2_rt = 68
    Const M239_E22_ext2_dt = 69
    Const M239_E22_ext3_cd = 70
    Const M239_E22_ext3_qty = 71
    Const M239_E22_ext3_amt = 72
    Const M239_E22_ext3_rt = 73
    Const M239_E22_ext3_dt = 74

    
    Const M239_E19_po_type_cd = 0
    Const M239_E19_po_type_nm = 1
    
    Const M239_E9_bp_cd = 0
	Const M239_E9_bp_nm = 1
    
    Const M239_E20_pur_grp = 0
    Const M239_E20_pur_grp_nm = 1
    
    
    Dim M31119
    Dim E1_b_bank_bank_nm
    Dim E2_b_minor_vat_type
    Dim E3_b_minor_pay_meth
    Dim E4_b_minor_pay_type
    Dim E5_b_minor_incoterms
    Dim E6_b_minor_transport
    Dim E7_b_minor_delivery_plce
    Dim E8_b_minor_origin
    Dim E9_b_biz_partner
    Dim E10_b_biz_partner_applicant_nm
    Dim E11_b_biz_partner_manufacturer_nm
    Dim E12_b_minor_packing_cond
    Dim E13_b_minor_inspect_means
    Dim E14_b_minor_dischge_city
    Dim E15_b_minor_dischge_port
    Dim E16_b_minor_loading_port
    Dim E17_b_configuration_reference
    Dim E18_b_currency_currency_desc
    Dim E19_m_config_process
    Dim E20_b_pur_grp
    Dim E21_b_biz_partner_agent_nm
    Dim E22_m_pur_ord_hdr
    Dim iCommandSent, iPoNo
    
    On Error Resume Next
    Err.Clear
    
    iPoNo = Trim(Request("txtPoNo"))    
    Set M31119 = Server.CreateObject("PM3G119.cMLookupPurOrdHdrS")

    If CheckSYSTEMError(Err, True) = True Then
        Exit Sub
    End If
    
     Call M31119.M_LOOKUP_PUR_ORD_HDR_SVR(gStrGlobalCollection, _
                                      iPoNo, E1_b_bank_bank_nm, E2_b_minor_vat_type, _
                                      E3_b_minor_pay_meth, E4_b_minor_pay_type, E5_b_minor_incoterms, _
                                      E6_b_minor_transport, E7_b_minor_delivery_plce, E8_b_minor_origin, _
                                      E9_b_biz_partner, E10_b_biz_partner_applicant_nm, _
                                      E11_b_biz_partner_manufacturer_nm, E12_b_minor_packing_cond, _
                                      E13_b_minor_inspect_means, E14_b_minor_dischge_city, _
                                      E15_b_minor_dischge_port, E16_b_minor_loading_port, _
                                      E17_b_configuration_reference, E18_b_currency_currency_desc, _
                                      E19_m_config_process, E20_b_pur_grp, _
                                      E21_b_biz_partner_agent_nm, E22_m_pur_ord_hdr)
                                                  
    If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
        Set M31119 = Nothing                                                '☜: ComProxy Unload
        Exit Sub                                                            '☜: 비지니스 로직 처리를 종료함 
     End If

   'If UCase(Trim(E22_m_pur_ord_hdr(M239_E22_ret_flg))) = "Y" Then
   '   Call DisplayMsgBox("17a014", vbOKOnly, "반품발주건", "조회", I_MKSCRIPT)
   '   Set M31119 = Nothing                                                                 '☜: ComProxy UnLoad
   '   Exit Sub                                                            '☜: 비지니스 로직 처리를 종료함 
   'End If
   
   Set M31119 = Nothing                                                                 '☜: ComProxy UnLoad

    '-----------------------
    'LookUp Iv Name
    '-----------------------
    '===================
    Dim strIvTypeNm
    Dim strPayeeCd
	Dim strPayeeNm
	Dim strBuildCd
	Dim strBuildNm
	
	Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
	If  E22_m_pur_ord_hdr(M239_E22_iv_type) <> "" Or E22_m_pur_ord_hdr(M239_E22_iv_type) <> Null then  			
		lgStrSQL = "select iv_type_nm from m_iv_type " 
		lgStrSQL = lgStrSQL & " WHERE iv_type_cd =  " & FilterVar(UCase(E22_m_pur_ord_hdr(M239_E22_iv_type)), "''", "S") & "" 
		
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			strIvTypeNm = ""
		Else
			strIvTypeNm	= lgObjRs("iv_type_nm")
		End If
	End If
		
	lgStrSQL = "SELECT A.BP_CD, A.BP_NM  FROM B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B  " 
	lgStrSQL = lgStrSQL & " WHERE B.PARTNER_BP_CD = A.BP_CD AND B.DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND B.BP_CD =  " & FilterVar(E9_b_biz_partner(M239_E9_bp_cd), "''", "S") & ""  
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
		strPayeeCd = ""
		strPayeeNm = ""
	Else
		strPayeeCd	= lgObjRs("BP_CD")
		strPayeeNm	= lgObjRs("BP_NM")
	End If
		
	lgStrSQL = "SELECT A.BP_CD, A.BP_NM  FROM B_BIZ_PARTNER A, B_BIZ_PARTNER_FTN B  " 
	lgStrSQL = lgStrSQL & " WHERE B.PARTNER_BP_CD = A.BP_CD AND B.DEFAULT_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND B.PARTNER_FTN = " & FilterVar("MBI", "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND B.BP_CD =  " & FilterVar(E9_b_biz_partner(M239_E9_bp_cd), "''", "S") & ""  
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
		strBuildCd = ""
		strBuildNm = ""
	Else
		strBuildCd	= lgObjRs("BP_CD")
		strBuildNm	= lgObjRs("BP_NM")
	End If
		
	Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)
    
    '========================


    Dim StrtxtBuildCd
    If E22_m_pur_ord_hdr(M239_E22_build_cd) = "" Then
        StrtxtBuildCd = E9_b_biz_partner(M239_E9_bp_cd)
    Else
        StrtxtBuildCd = E22_m_pur_ord_hdr(M239_E22_build_cd)
    End If

    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "With parent" & vbCr
    '##### Rounding Logic #####
        '항상 거래화폐가 우선 
    Response.Write ".frm1.txtCur.value          = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)) & """" & vbCR
    Response.Write "parent.CurFormatNumericOCX" & vbCr
    '##########################
    Response.Write ".frm1.txtIvAmt.text         = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtIvLocAmt.text      = """ & UNIConvNumDBToCompanyByCurrency(0, gCurrency, ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtVatAmt.text        = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, gTaxRndPolicyNo, "X") & """" & vbCr
    
    Response.Write ".frm1.txtnetDocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtnetLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, gCurrency, ggAmtOfMoneyNo, "X", "X") & """" & vbCr
    Response.Write ".frm1.txtVatLocAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(0, ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)), ggAmtOfMoneyNo, gTaxRndPolicyNo, "X") & """" & vbCr
    
    Response.Write ".frm1.txtGrpCd.Value        = """ & ConvSPChars(E20_b_pur_grp(M239_E20_pur_grp)) & """" & vbCr
    Response.Write ".frm1.txtGrpNm.Value        = """ & ConvSPChars(E20_b_pur_grp(M239_E20_pur_grp_nm)) & """" & vbCr
    Response.Write ".frm1.txtSpplCd.Value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_cd)) & """" & vbCr
    Response.Write ".frm1.hdnSpplCd.value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_cd)) & """" & vbCr
    Response.Write ".frm1.txtSpplNm.Value       = """ & ConvSPChars(E9_b_biz_partner(M239_E9_bp_nm)) & """" & vbCr

    'Response.Write ".frm1.txtPayeeCd.Value      = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_cd)) & """" & vbCr
    'Response.Write ".frm1.txtPayeeNm.Value      = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_nm)) & """" & vbCr

    'Response.Write ".frm1.txtBuildCd.Value      = """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_cd)) & """" & vbCr
    'Response.Write ".frm1.txtBuildNm.Value      = """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_nm)) & """" & vbCr
'=========
    Response.Write ".frm1.txtPayeeCd.Value      = """ & ConvSPChars(strPayeeCd) & """" & vbCr
    Response.Write ".frm1.txtPayeeNm.Value      = """ & ConvSPChars(strPayeeNm) & """" & vbCr

    Response.Write ".frm1.txtBuildCd.Value      = """ & ConvSPChars(strBuildCd) & """" & vbCr
    Response.Write ".frm1.txtBuildNm.Value      = """ & ConvSPChars(strBuildNm) & """" & vbCr
'=========
    Response.Write ".frm1.hdnCur.Value          = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_cur)) & """" & vbCr
    
    Response.Write ".frm1.txtCurNm.Value        = """ & ConvSPChars(E18_b_currency_currency_desc) & """" & vbCr
    Response.Write ".frm1.txtVatCd.Value        = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_type)) & """" & vbCr
    Response.Write ".frm1.txtVatNm.Value        = """ & ConvSPChars(E2_b_minor_vat_type) & """" & vbCr
    Response.Write ".frm1.txtVatRt.Text         = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_vat_rt), ggExchRate.DecPoint, 0) & """" & vbCr
    'Response.Write ".frm1.txtPayMethCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_meth)) & """" & vbCr
    'Response.Write ".frm1.hdnPayMethCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_meth)) & """" & vbCr
        
    Response.Write ".frm1.txtPayMethNm.Value    = """ & ConvSPChars(E3_b_minor_pay_meth) & """" & vbCr
    Response.Write ".frm1.txtPayDur.Text        = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_pay_dur), 0, 0) & """" & vbCr
    Response.Write ".frm1.txtPayTypeCd.Value    = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_type)) & """" & vbCr
    Response.Write ".frm1.txtPayTypeNm.Value    = """ & ConvSPChars(E4_b_minor_pay_type) & """" & vbCr
    Response.Write ".frm1.txtPayTermsTxt.Value  = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_pay_terms_txt)) & """" & vbCr
    Response.Write ".frm1.txtIvTypeCd.Value     = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_iv_type)) & """" & vbCr
    Response.Write ".frm1.txtIvTypeNm.value     = """ & strIvTypeNm & """" & vbCr
    Response.Write ".frm1.txtXchRt.text         = """ & UNINumClientFormat(E22_m_pur_ord_hdr(M239_E22_xch_rt), ggExchRate.DecPoint, 0) & """" & vbCr
    Response.Write ".frm1.txtPoNo.Value         = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_no)) & """" & vbCr
    
    Response.Write ".frm1.hdnDiv.value           = """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_xch_rate_op)) & """" & vbcr
    
    Response.Write "If """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_vat_inc_flag)) & """ = ""2""  then " & vbCr   'vat 포함 여부 
    Response.Write "  .frm1.rdoVatFlg2.Checked = true " & vbCr
    Response.Write "  .frm1.hdvatFlg.value  = ""2""" & vbCr
    Response.Write "Else" & vbCr
    Response.Write "  .frm1.rdoVatFlg1.Checked = true " & vbCr
    Response.Write "  .frm1.hdvatFlg.value  = ""1""" & vbCr
    Response.Write "End If  " & vbCr
    
    '이성룡 주석 
    'Response.Write "If """ & ConvSPChars(E22_m_pur_ord_hdr(M239_E22_po_no)) & """  <> """" Then " & vbCr
    'Response.Write "  .frm1.chkPoNo.checked = True" & vbCr
    'Response.Write "  .frm1.txtChkPoNo.value = ""Y"" " & vbCr
    'Response.Write "End If" & vbCr
    
    '===========
    '반품여부 Flag추가(2003.08.01)
    Response.Write "If """ & ConvSPChars(UCase(E22_m_pur_ord_hdr(M239_E22_ret_flg))) & """ = ""Y"" Then " & vbCr
    Response.Write "  .frm1.hdnRetflg.value = ""Y"" " & vbCr
    Response.Write "Else " & vbCr
    Response.Write "  .frm1.hdnRetflg.value = ""N"" " & vbCr
    Response.Write "End If" & vbCr
    '============
    Response.Write "If Trim(.frm1.txtPayeeCd.Value) = """"  then " & vbCr
    Response.Write "  .frm1.txtPayeeCd.Value =  .frm1.txtSpplCd.value" & vbCr
    Response.Write "End If" & vbCr
    Response.Write "If Trim(.frm1.txtPayeeNm.Value) = """" then " & vbCr
    Response.Write "  .frm1.txtPayeeNm.Value =  .frm1.txtSpplNm.value" & vbCr
    Response.Write "End If" & vbCr

    Response.Write "If Trim(.frm1.txtBuildCd.Value) = """" then " & vbCr
    Response.Write "  .frm1.txtBuildCd.Value =  .frm1.txtSpplCd.value " & vbCr
    Response.Write "End If" & vbCr
    Response.Write "If Trim(.frm1.txtBuildNm.Value) = """" then" & vbCr
    Response.Write "  .frm1.txtBuildNm.Value =  .frm1.txtSpplNm.value" & vbCr
    Response.Write "End If" & vbCr
    
    
    Response.Write "parent.GetPayDt()" & vbCr
    Response.Write "parent.GetTaxBizArea(""BP"")" & vbCr
    Response.Write "parent.ChangeTag(False)" & vbCr
    'parent.DbPoQueryOK()
    
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr
'---------------------------
	On Error Resume Next
	Err.Clear

	Dim BpType
	Dim iPB5CS41
	Dim E8_b_biz_partner
	
	Const S074_E1_bp_rgst_no = 2

    On Error Resume Next
    Err.Clear

    Set iPB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing	
        Exit Sub
    End If

    BpType = Request("txtBpType")
    
    Call iPB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection,"QUERY",StrtxtBuildCd,E8_b_biz_partner) 

    If CheckSYSTEMError(Err, True) = True Then
        Set iPB5CS41 = Nothing
        Exit Sub
    End If
    
    Response.Write "<Script Language=VBScript>" & vbCr
    Response.Write "Parent.frm1.txtSpplRegNo.Value	= """ & ConvSPChars(E8_b_biz_partner(S074_E1_bp_rgst_no)) & """" & vbCr
    Response.Write "parent.DbPoQueryOK()" & vbCr
	Response.Write "</Script>" & vbCr

    Set iPB5CS41 = Nothing                                                   '☜: Unload Comproxy
    Set M31119   = Nothing                                                   '☜: Unload Comproxy
    Set iPB5GS45 = Nothing

End Sub


%>
