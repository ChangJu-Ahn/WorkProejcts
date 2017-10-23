<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb3
'*  4. Program Name         : �����ڻ���泻����� 
'*  5. Program Desc         : �����ڻ꺰 ��泻���� ���� 
'*  6. Comproxy List        : +As0041ManageSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : ������ 
'* 10. Modifier (Last)      : ������ 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%	
On Error Resume Next													
Call HideStatusWnd														

    Call LoadBasisGlobalInf()
    
    '-- Common --
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '��: Set to space
'    lgOpModeCRUD      = Request("txtMode")       
    
'-------------------------
' ����, ��� ���� 
'-------------------------
	Dim iPAAG010																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

	'Import Variant
	Dim I3_a_asset_acq
	Dim I4_ief_supplied
	Dim E1_a_asset_master
	Dim E3_a_asset_acq
	Dim IG2_import_itm_grp

	'Import Const
	'View Name : import a_asset_acq
	Public Const A504_I3_acq_no = 0
	Public Const A504_I3_acq_fg = 2
	Public Const A504_I3_ap_no = 18
	Public Const A504_I3_gl_no = 19

	'View Name : import_mode_fg ief_supplied
	Public Const A504_I4_select_char = 0

	'View Name : export a_asset_master
	Public Const A504_E1_asst_no = 0

	'View Name : export a_asset_acq
	Public Const A504_E3_acq_no = 0    

'-------------------------   
' ���� ó�� 
'-------------------------
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 
    
    strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then											'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
	ElseIf Request("txtAcqNo") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'���� ���ǰ��� ����ֽ��ϴ�!           
		Response.End 
	End If


	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqMngSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If    
	
	Redim I3_a_asset_acq(30)
	Redim I4_ief_supplied(0)
	
    I3_a_asset_acq(A504_I3_acq_no) = Trim(Request("txtAcqNo"))
    I3_a_asset_acq(A504_I3_acq_fg) = Request("cboAcqFg")
    I3_a_asset_acq(A504_I3_gl_no)  = Trim(Request("txtGLNo"))
    I3_a_asset_acq(A504_I3_ap_no)  = Trim(Request("txtApNo"))
    
	I4_ief_supplied(A504_I4_select_char) = "D"
	
	E1_a_asset_master = Request("txtSpread_m")		'Master Data Spread
	IG2_import_itm_grp = Request("txtSpread_i")		'���󼼳��� Spread
		
	call iPAAG010.AS0021_ACQ_MANAGE_SVR(gStrGloBalCollection, _
										, , I3_a_asset_acq, I4_ief_supplied, E1_a_asset_master, IG2_import_itm_grp, , , , , _
										, E3_a_asset_acq)
            
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
    Response.Write " parent.DbDeleteOk()		" & vbCr
    Response.Write " </Script>					" & vbCr
%>