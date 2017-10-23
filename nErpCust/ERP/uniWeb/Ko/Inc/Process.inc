

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "../../ComAsp/ProcessSvr.asp"

Sub InitVariables()
End Sub

Sub SetDefaultVal()
End Sub


Sub Form_Load()

    Call InitVariables

    Call SetToolbar("110000000000001")
    Call SetDefaultVal
    
	Call FncQuery()
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


Function FncQuery()

    Dim IntRetCD 
    
    FncQuery = False                                                    
    
    Err.Clear                                                           
    
    Call DbQuery
       
    FncQuery = True														
            
End Function

Function FncNew() 
End Function

Function FncDelete()
End Function

Function FncSave() 
End Function

Function FncCopy() 
End Function

Function FncCancel() 
End Function


Function FncInsertRow() 
End Function

Function FncDeleteRow() 
End Function


Function FncPrint() 
End Function


Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)
End Function

Function FncExit()
	FncExit = True
End Function


Function DbDelete() 
End Function

Function DbDeleteOk()
End Function

Function DbQuery()
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtPrcsCD=" & PRCS_CD
    
	Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True                                                      
End Function

Function DbQueryOk()
End Function

Function DbSave() 
End Function

Function DbSaveOk()
End Function

</SCRIPT>