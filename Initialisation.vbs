Dim objuft

set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("C:\Users\sfjbs\Desktop\Opencart Data Driven\Driver\GUITest3Action1")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing

