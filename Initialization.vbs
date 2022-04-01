Dim objuft

Set objuft=Createobject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.Open("D:\DataDriven\driver1\Driver")
objuft.Test.Run
objuft.Test.Close
objuft.quit
Set objuft=nothing