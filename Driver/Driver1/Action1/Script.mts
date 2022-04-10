'Datatable.AddSheet "Module"
'Datatable.ImportSheet "C:\Magicbricks\Organizer\organizer.xlsx",1,"Module"
'Datatable.ImportSheet "C:\Magicbricks\Organizer\organizer.xlsx",2,"Testcase"
'Datatable.ImportSheet "C:\Magicbricks\Organizer\organizer.xlsx",3,"TestStep"


'transaction point start here
Services.StartTransaction "Magicbricks"

mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount
For i = 1 To mrowcount Step 1
Datatable.SetCurrentRow(i)
Modexe=Datatable("Moduleexe","Action1")
'msgbox Modexe
If Modexe="Y" Then
        Modid=Datatable("ModuleID","Action1")
        msgbox Modid
        trowcount=datatable.GetSheet("Action2").GetRowCount
        msgbox trowcount
        For j=1 To trowcount Step 1
    Datatable.SetCurrentRow(j)
    If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" then
    testcaseid=Datatable("TestcaseId","Action2")
    msgbox TestcaseId
        
        tsrowcount=Datatable.GetSheet("Action3").GetRowCount
        msgbox tsrowcount
        For k = 1 to tsrowcount Step 1
        datatable.SetCurrentRow(k)
        If testcaseid=Datatable("TestcaseId","Action3") Then
        keyword=Datatable("Keyword","Action3")
        msgbox keyword       
     
      Select case (Keyword)
      
        Case "ur"
        Call OpenUrl()
       
        
        Case "va"
        Call ValidAmount()
        
        Case "bt"
        Call BalanceTransfer()
        
        Case "sb"
        Call SelectBanks()
        

       End  Select
 
       
    End If
       
    Next
   
   
End If
  
Next


 End If
       
    Next
   
  
Services.EndTransaction "Magicbricks"
'transaction point end here










