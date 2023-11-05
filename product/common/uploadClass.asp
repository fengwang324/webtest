<%''on error resume next
  Dim jpg(1):jpg(0)=CByte(&HFF):jpg(1)=CByte(&HD8)
  Dim bmp(1):bmp(0)=CByte(&H42):bmp(1)=CByte(&H4D)
  Dim png(3):png(0)=CByte(&H89):png(1)=CByte(&H50):png(2)=CByte(&H4E):png(3)=CByte(&H47)
  Dim gif(5):gif(0)=CByte(&H47):gif(1)=CByte(&H49):gif(2)=CByte(&H46):gif(3)=CByte(&H39):gif(4)=CByte(&H38):gif(5)=CByte(&H61)

  '检查是否正常图片文件
  Function CheckFileType(FNID)
	  CheckFileType=false
	  Dim fstream,fileExt,stamp,i
	  fileExt=LCase(Mid(FNID,InStrRev(FNID,".")+1))
	  Set fstream=Server.CreateObject("ADODB.Stream")
	  fstream.Open
	  fstream.Type=1
	  fstream.LoadFromFile FNID
	  fstream.position=0
	  Select Case fileExt
		  Case "jpg","jpeg"
			  stamp=fstream.read(2)
			  For i=0 to 1
				  if ascB(MidB(stamp,i+1,1))=jpg(i) then CheckFileType=true else CheckFileType=false
			  Next
		  Case "gif"
			  stamp=fstream.read(6)
			  For i=0 to 5
				if ascB(MidB(stamp,i+1,1))=gif(i) then CheckFileType=true else CheckFileType=false
			  Next
		  Case "png"
			  stamp=fstream.read(4)
			  For i=0 to 3
				if ascB(MidB(stamp,i+1,1))=png(i) then CheckFileType=true else CheckFileType=false
			  Next
		  Case "bmp"
			  stamp=fstream.read(2)
			  For i=0 to 1
				  if ascB(MidB(stamp,i+1,1))=bmp(i) then CheckFileType=true else CheckFileType=false
			  Next
		  Case "pdf","xls","doc","rar"
		      CheckFileType = True
		  Case else
		      CheckFileType = False
	  End Select
	  fstream.Close
	  Set fstream=Nothing
	  if Err.Number<>0 then CheckFileType=false
	  
	  If CheckFileType=False then			
			Dim fso,ficn
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set ficn = fso.GetFile(FNID)
				ficn.Delete
			Set ficn=Nothing
			Set fso=Nothing
			Response.Write "Please do not upload Error File"
			Response.End
	  End if
  End Function

  '检查是否本地文件
  Function TrueStr(fileTrue)
	  Dim str_len,pos
	  str_len=Len(fileTrue)
	  pos=Instr(fileTrue,chr(0))
	  if pos=0 or pos=str_len then
		 TrueStr=True
	  else
		 TrueStr=False
	  end if
	  IF TrueStr=False then
		  Response.Write "Please do not upload Error Files"	
		  Response.End
	  End IF
  End Function
%>  
<%Dim Data_5xsoft

  '无组件上传类
  Class upload_5xsoft
  
      Dim objForm,objFile,Version

	  Public Function Form(strForm)
	      strForm=LCase(strForm)
	      if not objForm.exists(strForm) then
		     Form=""
	      else
		     Form=objForm(strForm)
	      end if
	  End Function

      Public Function File(strFile)
          strFile=LCase(strFile)
          if not objFile.exists(strFile) then
              Set File=new FileInfo
          else
              Set File=objFile(strFile)
          end if
      End Function


	  Private Sub Class_Initialize 
		  Dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile
		  Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		  Dim iFindStart,iFindEnd
		  Dim iFormStart,iFormEnd,sFormName
		  Dim nTotalBytes,ReadBytes,nPartBytes
		  ReadBytes = 0
		  Version="Version 2.0"
		  Set objForm=Server.CreateObject("Scripting.Dictionary")
		  Set objFile=Server.CreateObject("Scripting.Dictionary")
		  nTotalBytes = Request.TotalBytes
		  if nTotalBytes<1 then Exit Sub
		  Set tStream = Server.CreateObject("Adodb.Stream")
		  Set Data_5xsoft = Server.CreateObject("Adodb.Stream")
		  Data_5xsoft.Type = 1
		  Data_5xsoft.Mode =3
		  Data_5xsoft.Open
		  Rem 循环分块读取
		  Do While ReadBytes < nTotalBytes
			  Rem 分块读取
			  nPartBytes = 64 * 1024 Rem 分成每块64k
			  If nPartBytes + ReadBytes > nTotalBytes Then
				nPartBytes = nTotalBytes - ReadBytes
			  End If
			  Data_5xsoft.Write Request.BinaryRead(nPartBytes)
			  ReadBytes = ReadBytes + nPartBytes
		  Loop
		  Rem Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes) 旧的只可以上传200K文件
		  Data_5xsoft.Position=0
		  RequestData =Data_5xsoft.Read
		
		  iFormStart = 1
		  iFormEnd = LenB(RequestData)
		  vbCrlf = chrB(13) & chrB(10)
		  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1)
		  iStart = LenB (sStart)
		  iFormStart=iFormStart+iStart+1
		  While (iFormStart + 10) < iFormEnd 
		      iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3
			  tStream.Type = 1
			  tStream.Mode = 3
			  tStream.Open
			  Data_5xsoft.Position = iFormStart
			  Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart
			  tStream.Position = 0
			  tStream.Type = 2
			  tStream.CharSet ="UTF-8"
			  sInfo = tStream.ReadText
			  tStream.Close
			  
			  'GETFORMNAME
			  iFormStart = InStrB(iInfoEnd,RequestData,sStart)
			  iFindStart = InStr(22,sInfo,"name=""",1)+6
			  iFindEnd   = InStr(iFindStart,sInfo,"""",1)
			  sFormName  = LCase(Mid(sinfo,iFindStart,iFindEnd-iFindStart))
			  'IF IS FILE
			  if InStr (45,sInfo,"filename=""",1) > 0 then
				  Set theFile=new FileInfo
				  'GET FILENAME
				  iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
				  iFindEnd = InStr(iFindStart,sInfo,"""",1)
				  sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
				  theFile.FileName=getFileName(sFileName)
				  theFile.FilePath=getFilePath(sFileName)
				  'GET FILETYPE
				  iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
				  iFindEnd = InStr(iFindStart,sInfo,vbCr)
				  theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart)
				  theFile.FileStart =iInfoEnd
				  theFile.FileSize = iFormStart -iInfoEnd -3
				  theFile.FormName=sFormName
				  if not objFile.Exists(sFormName) then
				     objFile.add sFormName,theFile
				  end if
			  else
			     'IF IS FORM
				 tStream.Type =1
				 tStream.Mode =3
				 tStream.Open
				 Data_5xsoft.Position = iInfoEnd 
				 Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3
				 tStream.Position = 0
				 tStream.Type = 2
				 tStream.CharSet ="UTF-8"
				 sFormValue = tStream.ReadText 
				 tStream.Close
				 if objForm.Exists(sFormName) then
				     objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
				 else
				     objForm.Add sFormName,sFormValue
				 end if
			  end if
			  iFormStart=iFormStart+iStart+1
		  Wend
		  RequestData=""
		  Set tStream = Nothing
	  End Sub

	  Private Sub Class_Terminate  
	      if Request.TotalBytes>0 then
			  objForm.RemoveAll
			  objFile.RemoveAll
			  Set objForm=Nothing
			  Set objFile=Nothing
			  Data_5xsoft.Close
			  Set Data_5xsoft =Nothing
	      end if
	  End Sub   
 
	  Private Function GetFilePath(FullPath)
		  If FullPath <> "" Then
		      GetFilePath = Left(FullPath,InStrRev(FullPath, "\"))
		  Else
		      GetFilePath = ""
		  End If
	  End  Function
	 
	  Private Function GetFileName(FullPath)
		  If FullPath <> "" Then
		      GetFileName = Mid(FullPath,InStrRev(FullPath, "\")+1)
		  Else
		      GetFileName = ""
		  End If
	  End Function
  End Class

  '上传文件信息类
  Class FileInfo
	  Dim FormName,FileName,FilePath,FileSize,FileType,FileStart
	  Private Sub Class_Initialize 
		  FileName = ""
		  FilePath = ""
		  FileSize = 0
		  FileStart= 0
		  FormName = ""
		  FileType = ""
	  End Sub
	  
	  Public Function SaveAs(FullPath)
		  Dim dr,ErrorChar,i
		  SaveAs=true
		  if Trim(fullpath)="" or FileStart=0 or FileName="" or Right(fullpath,1)="/" then
		      Exit Function
		  end if
		  Set dr=CreateObject("Adodb.Stream")
		  dr.Mode=3
		  dr.Type=1
		  dr.Open
		  Data_5xsoft.position=FileStart
		  Data_5xsoft.copyto dr,FileSize
		  dr.SaveToFile FullPath,2
		  dr.Close
		  Set dr=Nothing 
		  SaveAs=False
	  End Function
  End Class
  
'===================================================================================================
'方法名：FormatPic                                                               
'作  用：按比例缩放图片大小                                                                    
'参  数：PicPath -- 图片相对路径, w -- 新宽度，h -- 新高度, vs -- 垂直间距，hs -- 水平间距
  Function FormatPic(PicPath,w,h,vs,hs)
  	  Dim Jpeg,newW,newH,oldPp,newPp
	  Set Jpeg = Server.CreateObject("Persits.Jpeg")
	  Jpeg.Open Server.MapPath(PicPath)
	  oldPp = Jpeg.OriginalHeight / Jpeg.OriginalWidth
	  newPp = h / w
	  if oldPp = newPp then
	    newH = h
		newW = w
	  elseif oldPp > newPp then 
		IF Jpeg.Height > h THEN
			newH = h
			newW = Jpeg.OriginalWidth * h / Jpeg.OriginalHeight
		ELSE
		    newW = Jpeg.OriginalWidth
			newH = Jpeg.OriginalHeight
		END IF
	  elseif oldPp < newPp then
		IF Jpeg.Width > w THEN
			newW = w
			newH = Jpeg.OriginalHeight * w / Jpeg.OriginalWidth
		ELSE
		    newW = Jpeg.OriginalWidth
			newH = Jpeg.OriginalHeight
		END IF
	  end if
      Set Jpeg = Nothing
	  FormatPic = "<img border=0  src="""&PicPath&""" width="""&newW&""" height="""&newH&""" "&_
	              " vspace="""&vs&""" hspace="""&hs&""" border=""0"" />"
  End Function
  
  '重新生成规格图片 PicPath -- 路径; w -- 宽度; h - 高度.
  Sub RePic(PicPath,w,h)
  	  Dim Jpeg,oldPp,newPp
	  Set Jpeg = Server.CreateObject("Persits.Jpeg")
	  Jpeg.Open PicPath
	  oldPp = Jpeg.OriginalHeight / Jpeg.OriginalWidth
	  newPp = h / w
	  if oldPp > newPp then 
		  IF Jpeg.Height > h THEN
			  Jpeg.Height = h
			  Jpeg.Width = Jpeg.OriginalWidth * h / Jpeg.OriginalHeight
		  END IF
	  else
	      IF Jpeg.Width > w THEN
			  Jpeg.Width = w
			  Jpeg.Height = Jpeg.OriginalHeight * w / Jpeg.OriginalWidth
		  END IF
	  end if 
	  Jpeg.Save PicPath
      Set Jpeg = Nothing
  End Sub
  
  '图片变动时，删除旧图片 
  Sub DelPic(PathID)
 	  Dim FS
	  Set FS = Server.CreateObject("Scripting.FileSystemObject")
  	  if FS.FileExists(Server.MapPath(PathID)) then
     	  FS.DeleteFile Server.MapPath(PathID),True
  	  end if
  	  Set FS = Nothing
  End Sub
  
  Function PrintPic(Path,Model,Content,LLeft,LTop)
	  Dim Jpeg,L,T,S
	  Set Jpeg = Server.CreateObject("Persits.Jpeg")
	  Jpeg.Open Path 
	  
	  '水印文字
	  Jpeg.Canvas.Font.Color = &HA8A8A8 ' 色
	  Jpeg.Canvas.Font.Family = "宋体"
	  Jpeg.Canvas.Font.Italic = True
	  Jpeg.Canvas.Font.Size = 16
	  Jpeg.Canvas.Font.Quality = 3
	  Jpeg.Canvas.Font.ShadowColor = &HECECEC  '//水印文字的阴影色彩。
	  Jpeg.Canvas.Font.ShadowXoffset = 1  '//水印文字阴影向右偏移的像素值，输入负值则向左偏移。
	  Jpeg.Canvas.Font.ShadowYoffset = 1  '//水印文字阴影向下偏移的像素值，输入负值则向右偏移。
	  Jpeg.Canvas.Font.Bold = True  
	  Jpeg.Canvas.Print LLeft, LTop, Content '打印
		  
	  Jpeg.Save Path 
	  Set Jpeg = Nothing
  End Function
  
  Sub Rename(OFILE,NFILE)
      Dim rn
	  Set rn = Server.CreateObject("Scripting.FileSystemObject")
	  'on error resume next
	  rn.MoveFile OFILE, NFILE
	  If Err.Number = 53 Then
			Response.Write File & "文件不存在！"
			Response.End
	  Elseif Err.Number = 58 Then
			Response.Write File & "文件已存在！"
			Response.End
	  Elseif Err.Number <> 0 Then
			Response.Write "未知错误，错误编码：" & Err.Number
			Response.End
	  End If
	  Set rn = Nothing
  End Sub
  
  '检查上传文件夹是否存在，不存在则创建文件夹
  Private Function CheckFloder(FolderName) 
      Dim objFso,FLDR
	  FLDR = Server.Mappath(FolderName)
      Set objFso = CreateObject("Scripting.FileSystemObject")
	  If Not objFso.FolderExists(FLDR) Then
		  objFso.CreateFolder(FLDR)
	  End If
      Set objFso = Nothing
	  CheckFloder = FolderName
  End Function %>