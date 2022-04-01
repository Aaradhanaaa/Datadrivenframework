'Datatable.AddSheet "Module"
'Datatable.ImportSheet "‪C:\Users\All uft\KeywordDriven\organizer.xlsx",1,"Module"
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
msgbox testcaseid
     trowcount=Datatable.GetSheet("Action3").GetRowCount
     msgbox trowcount
     
     For k  = 1 To trowcount Step 1
     datatable.SetCurrentRow(k)
     If testcaseid =Datatable("TestcaseId","Action3") Then
     keyword=Datatable("Keyword","Action3")
     msgbox keyword
     Select Case (keyword)
     Case	"ln"
     Call Login("john","hp")
     
     Case "ca"
      Call Closeapp()
      
      Case "oo"
      Call openorder()
      Case "ua"
      Call Updateorder()
      Case "lnd"
      
      drowcount=datatable.GetSheet("Action4").GetRowCount
      For l = 1 To drowcount Step 1
      	    datatable.SetCurrentRow(l)
          Call login(datatable("username","Action4"),datatable("password","Action4"))
           Call Closeapp() 
      Next
      Case "ood"
      orrowcount=datatable.GetSheet("Action4").GetRowCount
      For m = 1 To orrowcount Step 1
      	datatable.SetCurrentRow(m)
      	Call openorder(datatable("orderno","Action4"))
      Next
      
    
     End Select
     End If
     Next
		
	End If
	Next
		
End If

Next @@ hightlight id_;_525364_;_script infofile_;_ZIP::ssf35.xml_;_


