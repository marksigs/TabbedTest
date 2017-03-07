Set fso = CreateObject("Scripting.FileSystemObject")
Set textfile = fso.OpenTextFile("Output_FormFill.xml")

strRequest = textfile.ReadAll

textfile.close

Set objOmiga3Download = CreateObject("Om3Manager.CustRegBo")

strResponse = objOmiga3Download.Create(strRequest)

msgbox strresponse