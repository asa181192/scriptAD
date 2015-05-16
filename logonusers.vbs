On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Description
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strIpPhone = objuser.ipPhone 

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.InlineShapes.AddPicture("\\MTYDC1\MTYShare\logo.jpg")
objselection.TypeText Chr(11)
objSelection.Font.Name = "Calibri" 
objSelection.Font.Size = 11 
objselection.Font.Bold = true 
objSelection.Font.Color = RGB (0,0,0) 
objSelection.TypeText strName
objselection.TypeText Chr(11)
objSelection.Font.Size = 10 
objSelection.Font.Color = RGB (0,76,153) 
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)
objselection.Font.Bold = false 
objSelection.TypeText "TransNetwork Corp."
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB (128,128,128)
objSelection.TypeText strPhone+" x:"+strIpPhone
objSelection.TypeText Chr(11)
objSelection.objSelection.TypeText "Web: "
objSelection.Font.Color = RGB (128,128,128)
objSelection.Font.Size = 10 	
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range, "http://transnetwork.com/TransnetworkSite/index.aspx", , , "www.TransNetwork.com") 
objSelection.TypeText Chr(10)
objSelection.Font.Color = RGB (0,0,0)
objSelection.TypeText "Warning : The information contained in this message may be privileged and confidential and protected from disclosure. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution or copying of this communication is strictly prohibited. If you have received this communication in error, please notify us immediately by replying to this message and then delete it from your computer. All e-mail sent to this address will be received by the Transnetwork corporate e-mail system and is subject to archiving and review by someone other than the recipient."
Set objSelection = objDoc.Range()

objSignatureEntries.Add "Firma estandar", objSelection
objSignatureObject.NewMessageSignature = "Firma estandar"

objDoc.Saved = True
objWord.Quit

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName
objSelection.TypeParagraph()
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Firma respuesta", objSelection

objSignatureObject.ReplyMessageSignature = "Firma respuesta"

objDoc.Saved = True
objWord.Quit