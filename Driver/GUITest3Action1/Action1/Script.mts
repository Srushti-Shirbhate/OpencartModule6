Services.StartTransaction "TR_MOD6"

mrowcount=datatable.GetSheet("Action1").GetRowcount
'msgbox mrowcount
For i = 1 To mrowcount Step 1
	Datatable.SetCurrentRow(i)
	Modexe=Datatable("Moduleexe","Action1")
	'msgbox Modexe
	If Modexe="Y" Then
		Modid=Datatable("ModuleID","Action1")
		'msgbox Modid
		trowcount=datatable.GetSheet("Action2").GetRowCount
		'msgbox trowcount
		For j = 1 To trowcount Step 1
			Datatable.SetCurrentRow(j)
			If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
				testcaseid=Datatable("TestcaseId","Action2")
				'msgbox testcaseid
				tsrowcount=datatable.GetSheet("Action3").GetRowCount
		        'msgbox tsrowcount
		        For k = 1 To tsrowcount Step 1
			    Datatable.SetCurrentRow(k)
			    If testcaseid=Datatable("TestcaseId","Action3") Then
				keyword=Datatable("Keyword","Action3")
				'msgbox keyword
				Select Case (keyword)
				        Case "M6_URL"
				        Call OpenURL
				        wait 3
				        
				        Case "M6_Ch"
				        'msgbox "Click on Checkout Option"
				        Call AddtoCart()
					Call Checkout()
					
				        Case "M6_GCh"
				        'msgbox "Click on GuestCheckout Option"
				        wait 2
					Call GuestCheckout()
					
					Case "M6_BD"
					trowcount=datatable.GetSheet("Action4").GetRowCount
		'msgbox trowcount
		For l = 1 To trowcount Step 1
			Datatable.SetCurrentRow(l)
					Call BillingDetails(datatable("First Name","Action4"),datatable("Last Name","Action4"),datatable("Email","Action4"),datatable("Telephone","Action4"),datatable("Address","Action4"),datatable("Company","Action4"),datatable("City","Action4"), datatable("Post Code","Action4"))
	
					
					Next
					
					Case "M6_DM"
					Call DeliveryMethod()
					
					Case "M6_PyM"
					Call PaymentMethod()
					
					Case "M6_CfO"
					'msgbox "Click on ConfirmOrder Option"
					Call ConfirmOrder()
				End Select
				
			End If
		Next
	End If

Next
End If
Next @@ script infofile_;_ZIP::ssf21.xml_;_

Services.EndTransaction "TR_MOD6"


'Browser("HP LP3065").
'Browser("HP LP3065").Page("HP LP3065").Sync

 @@ script infofile_;_ZIP::ssf35.xml_;_
'Browser("HP LP3065").Page("HP LP3065").Link("Laptops & Notebooks").Click
'Browser("HP LP3065").Page("HP LP3065").Link("Show All Laptops & Notebooks").Click
'Browser("HP LP3065").Page("HP LP3065").WebButton("Add to Cart").Click
'Browser("HP LP3065").Page("HP LP3065").WebButton("Add to Cart_2").Click
' @@ script infofile_;_ZIP::ssf28.xml_;_
' @@ script infofile_;_ZIP::ssf29.xml_;_

'Browser("HP LP3065").Page("HP LP3065").Link("Laptops & Notebooks").Click
 @@ script infofile_;_ZIP::ssf38.xml_;_

 @@ script infofile_;_ZIP::ssf40.xml_;_
