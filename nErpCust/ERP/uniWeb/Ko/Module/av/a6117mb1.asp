<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>



<%
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : VAT
'*  3. Program ID           : a6117mb1
'*  4. Program Name         : 부가세수정 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004.05.10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Eun Kyung , KANG
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
%>
<%					
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim lgOpModeCRUD
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
    
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
    
             Call SubBizSaveMulti()
    End Select
    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	Dim iPAVG080										' 입력/수정용 ComProxy Dll 사용 변수 


    Dim iStrData
    Dim exportData1
	
	Const C_EG1_bdg_cd_fr = 0
	Const C_EG1_bdg_nm_fr = 1
	Const C_EG1_bdg_cd_to = 2
	Const C_EG1_bdg_nm_to = 3
	Const C_EG1_dept_cd = 4
	Const C_EG1_dept_nm = 5
	Const C_EG1_bdg_yyyymm_fr=6
	Const C_EG1_bdg_yyyymm_to=7	
	
    
    Dim iLngRow,iLngCol
    
    Dim iStrNextKey
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim iIntLoopCount
    
    Dim importArray
	Dim lgStrPrevVatKey
	
	Const C_SHEETMAXROWS_D  = 100
	
	Const C_Issued_from		= 0
	Const C_Issued_to		= 1
	Const C_Report_Biz_Area = 2
	Const C_Io_Fg			= 3
	Const C_Vat_Type	    = 4
	Const C_Bp_Cd	    	= 5
    Const C_Next_Vat_No_Key	= 6	
    Const C_issue_dt_fg	    = 7
    Const C_cboissue_dt_kind= 8
	
    Const C_EG2_VAT_NO               = 0
    Const C_EG2_ISSUED_DT            = 1
    Const C_EG2_IO_FG                = 2
    Const C_EG2_BP_CD                = 3
    Const C_EG2_BP_NM                = 4
    Const C_EG2_MADE_VAT_FG          = 5
    Const C_EG2_VAT_TYPE             = 6
    Const C_EG2_VAT_TYPE_NM          = 7
    Const C_EG2_NET_LOC_AMT          = 8
    Const C_EG2_VAT_LOC_AMT          = 9
    Const C_EG2_C_CREDIT_CD          = 10
    Const C_EG2_REPORT_BIZ_AREA_CD   = 11
    Const C_EG2_REPORT_BIZ_AREA_NM   = 12
    Const C_EG2_BIZ_AREA_CD          = 13
    Const C_EG2_BIZ_AREA_NM          = 14
    Const C_EG2_GL_NO                = 15
    Const C_EG2_REF_NO               = 16	
    Const C_TAXNO                    = 17
    Const C_ISSUE_DT_KIND_CD         = 18
	Const C_ISSUE_DT_KIND_NM         = 19
    Const C_ISSUE_DT_FG_CD           = 20
	Const C_ISSUE_DT_FG_NM           = 21

	'Key 값을 읽어온다	
	
	lgStrPrevVatKey	= Trim(Request("lgStrPrevVatKey"))
	   
	iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")	
	
	 Redim importArray(C_cboissue_dt_kind)	
		    
	    importArray(C_Issued_from)		= UniConvDate(Request("txtIssuedDtFr"))
	    importArray(C_Issued_to)		= UniConvDate(Request("txtIssuedDtTo"))
	    importArray(C_Report_Biz_Area)	= FilterVar(UCase(Trim(Request("txtReportBizArea"))), "", "SNM")
	    importArray(C_Io_Fg)			= FilterVar(UCase(Trim(Request("cboIoFg"))), "", "SNM")
	    importArray(C_Vat_Type)  		= FilterVar(UCase(Trim(Request("txtVatType"))), "", "SNM")
	    importArray(C_Bp_Cd)		    = FilterVar(UCase(Trim(Request("txtBpCd"))), "", "SNM")
	    importArray(C_issue_dt_fg)		    = FilterVar(UCase(Trim(Request("txtissue_dt_fg_cd"))), "", "SNM")
	    importArray(C_cboissue_dt_kind)		    = FilterVar(UCase(Trim(Request("txtissue_dt_kind_cd"))), "", "SNM")
	    

	    importArray(C_Next_Vat_No_Key)	= lgStrPrevVatKey


	 If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If  
    
    Set iPAVG080 = Server.CreateObject("PAVG080.cFListVatSvr")
    
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
     
    Call iPAVG080.F_LIST_VAT_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,importArray, exportData1)
	
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set iPAVG080 = Nothing
       Exit Sub
    End If    
    
   
    
    Set IPAVG080 = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	
 
   If IsEmpty(exportData1) = False Then

       For iLngRow = 0 To UBound(exportData1, 1)
          iIntLoopCount = iIntLoopCount + 1

           If iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_VAT_NO)) 
                iStrData = iStrData & Chr(11) & UNIDateClientFormat(exportData1(iLngRow, C_EG2_ISSUED_DT))
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_IO_FG)) 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_BP_CD)) 
                iStrData = iStrData & Chr(11) & ""
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_BP_NM))  
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_TAXNO))
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_MADE_VAT_FG))
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_VAT_TYPE)) 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_VAT_TYPE_NM)) 
                iStrData = iStrData & Chr(11) & ""
  				iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_ISSUE_DT_FG_CD))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_ISSUE_DT_FG_NM))               
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_ISSUE_DT_KIND_CD))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_ISSUE_DT_KIND_NM))              
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, C_EG2_NET_LOC_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, C_EG2_VAT_LOC_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_C_CREDIT_CD)) 
                iStrData = iStrData & Chr(11) & ""                 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_REPORT_BIZ_AREA_CD)) 
                iStrData = iStrData & Chr(11) & ""                 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_BIZ_AREA_CD)) 
                iStrData = iStrData & Chr(11) & ""                 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_GL_NO)) 
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData1(iLngRow, C_EG2_REF_NO)) 
                                
                iStrData = iStrData & Chr(11) & CStr(iIntMaxRows + iLngRow + 1) & Chr(11) & Chr(12)


          Else
                lgStrPrevVatKey	= exportData1(iLngRow, C_EG2_VAT_NO)
                
                iIntQueryCount = iIntQueryCount + 1
                Exit For
         End If
       Next
    End If


    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
	    lgStrPrevVatKey = ""
	    iIntQueryCount = ""
    End If
    
	
	Response.Write " <Script Language=vbscript>								" & vbCr
	Response.Write " With parent											" & vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData			        " & vbcr
	Response.Write "	.ggoSpread.SSShowData """ & istrData & """	        " & vbcr
    Response.Write "    .lgStrPrevVatKey     = """ & lgStrPrevVatKey & """  " & vbCr	
	Response.Write "	.DbQueryOk()										" & vbcr
    Response.Write " End With												" & vbCr
    Response.Write " </Script>												" & vbCr
	

    
End Sub    	 
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    Dim iPAVG080		
    Dim importString 
    Dim iErrorPosition
 '   importString = Trim(Request("txtAcctCd"))
 
    Set iPAVG080 = Server.CreateObject("PAVG080.cFMngVatSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    
    Call iPAVG080.F_MANAGE_VAT_SVR(gStrGlobalCollection, Trim(Request("txtSpread")), iErrorPosition)			
 
	    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then				
       Set iPAVG080 = Nothing
		Exit Sub
    End If    
    
   
     Set iPAVG080 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
    
End Sub    

%>
