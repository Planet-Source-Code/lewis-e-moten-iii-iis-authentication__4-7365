<div align="center">

## IIS Authentication


</div>

### Description

Requests users to login to website with NT Account.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Security](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/security__4-14.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-iis-authentication__4-7365/archive/master.zip)





### Source Code

```
Dim ObjSite
Call Authenticate(ObjSite)
Sub Authenticate(ByRef pObjSite)
	Dim lLngInstanceID
	Dim lStrMetabasePath
	Dim lBlnContinue
	Dim lBlnLoginFailure
	On Error Resume Next
	lLngInstanceID = Request.ServerVariables("INSTANCE_ID")
	' Programmers Notes ...
	'
	' Metabase Path											Key Type
	' /LM/W3SVC												IIsWebService
	' /LM/W3SVC/N											IIsWebServer
	' /LM/W3SVC/N/ROOT										IIsWebVirtualDir
	' /LM/W3SVC/N/ROOT/WebVirtualDir						IIsWebVirtualDir
	' /LM/W3SVC/N/ROOT/WebVirtualDir/WebDirectory			IIsWebDirectory
	' /LM/W3SVC/N/ROOT/WebVirtualDir/WebDirectory/WebFile	IIsWebFile
	'
	' N = lLngInstanceID
	'
	'
	lStrMetabasePath = Request.ServerVariables("APPL_MD_PATH")
	lStrMetabasePath = Replace(lStrMetabasePath, "/LM/", "IIS://LOCALHOST/", 1, vbTextCompare)
	'
	'
	'
	Set pObjSite = GetObject(lStrMetabasePath)
	If Err = &H800401E4 Or Err = 70 Then
		Response.Status = "401 access denied"
		BlnContinue = False
		BlnLoginFailure = True
	Else
		If Err = 0 Then
			lBlnContinue = True
		Else
			lBlnContinue = False
			lBlnLoginFailure = False
		End If
	End If
	If lBlnLoginFailure Then
		Response.Write "Login Failure.<BR>"
		Response.End
	End If
	If Not lBlnContinue Then
		Response.Write "Can not continue.<BR>"
		Response.End
	End If
End Sub
```

