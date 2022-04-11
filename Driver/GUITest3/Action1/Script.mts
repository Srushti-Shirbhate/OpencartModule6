mrowcount=datatable.GetSheet("Action1").GetRowcount
msgbox mrowcount
For i = 1 To mrowcount Step 1
	Datatable.SetCurrentRow(i)
	Modexe=Datatable("Moduleexe","Action1")
	msgbox Modexe
	If Modexe="Y" Then
		Modid=Datatable("ModuleID","Action1")
		msgbox Modid
		trowcount=datatable.GetSheet("Action2").GetRowCount
		msgbox trowcount
		For j = 1 To trowcount Step 1
			Datatable.SetCurrentRow(j)
			If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
				testcaseid=Datatable("TestcaseId","Action2")
				msgbox testcaseid
				tsrowcount=datatable.GetSheet("Action3").GetRowCount
		        msgbox tsrowcount
		        For k = 1 To tsrowcount Step 1
			    Datatable.SetCurrentRow(k)
			    If testcaseid=Datatable("TestcaseId","Action3") Then
				keyword=Datatable("Keyword","Action3")
				msgbox keyword
				Select Case (keyword)
					Case "M6_GCh"
					Call GuestCheckout()
				End Select
				
			End If
		Next
	End If

Next
End If
Next @@ script infofile_;_ZIP::ssf21.xml_;_


