<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4114mb3.asp
'*  4. Program Name			: Operation Management
'*  5. Program Desc			: Save Production Order Detail
'*  6. Comproxy List		: PP4G121.cPMngProdOrdDtl
'*  7. Modified date(First)	: 2001/06/30
'*  8. Modified date(Last) 	: 2002/07/08
'*  9. Modifier (First)		: Park, Bum-Soo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->

<%														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

Call HideStatusWnd

On Error Resume Next

Dim oPP4G121
Dim iErrorPosition																			'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount
Dim lgStrSQL        '2008-05-19 5:59���� :: hanc
Dim lgObjConn , lgObjRs, lgstrData, adCmdText , adExecuteNoRecords

Dim iCUCount
Dim iDCount

Dim ii

Err.Clear																		'��: Protect system from crashing


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

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
itxtSpread = Join(itxtSpreadArr,"")


    lgStrSQL = " EXEC USP_TEMP_20080506001_KO441 " & FilterVar(Trim(Request("txtOpr"))	, "''", "S") & " "

'    lgStrSQL = lgStrSQL & "FROM NEPES..B_CONFIGURATION       "
'    lgStrSQL = lgStrSQL & "where major_cd like 'ZZ002'       "
'    lgStrSQL = lgStrSQL & "and   minor_cd = 'DAYS'           "


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    

    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
'       Response.Write  " parent.frm1.txtPeriod.value  = """ & UCase(Trim(tPeriod)) & """" & vbCr      
'       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
'       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
'       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DbSaveOk   " & vbCr      
'       Response.Write  "    Parent.InitSpreadSheet   " & vbCr           '20080303::hanc
       Response.Write  " </Script>             " & vbCr
    End If

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection
    
%>