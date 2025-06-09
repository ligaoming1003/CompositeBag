<%
'=========================================================
 '����: UpLoad(����������ϴ���)
 '����: Anlige
 '�汾: An-Upload������ϴ���8.10.20
 '��������: 2008-4-12
 '�޸�����: 2008-10-20
 '������ҳ: http://www.ii-home.cn
 'Email: zhanghuiguoanlige@126.com
 'QQ: 417833272
'=========================================================
Dim StreamT
Class AnUpLoad
 Private Form, Fils
 Private vCharSet, vMaxSize, vSingleSizeg, vErr, vVersion, vTotalSize, vExe, NewName

 '���úͶ�ȡ���Կ�ʼ
 Public Property Let MaxSize(ByVal value)
   vMaxSize = value
 End Property
 
 Public Property Let SingleSize(ByVal value)
   vSingleSize = value
 End Property
 
 Public Property Let Exe(ByVal value)
   vExe = LCase(value)
 End Property
 
 Public Property Let CharSet(ByVal value)
   vCharSet = value
 End Property
 
 Public Property Get Err()
   Err = vErr
 End Property
 
 Public Property Get Description()
   Description = GetErr(vErr)
 End Property
 
 Public Property Get Version()
   Version = vVersion
 End Property
 
 Public Property Get TotalSize()
   TotalSize = vTotalSize
 End Property
 
 '���úͶ�ȡ���Խ���

 Private Sub Class_Initialize()
   set StreamT=server.createobject("ADODB.STREAM")
   set Form = server.createobject("Scripting.Dictionary")
   set Fils = server.createobject("Scripting.Dictionary")
   vVersion = "Anlige������ϴ�8.12.14"
   vMaxSize = -1
   vSingleSize = -1
   vErr = -1
   vExe = ""
   vTotalSize = 0
   vCharSet = "gb2312"
 End Sub

 Private Sub Class_Terminate()
   Set Form = Nothing
   Set Fils = Nothing
   Set StreamT = Nothing
 End Sub

 '������:GetData
 '����:����ͻ����ύ������������
 '����:�ͻ����ύ�����������ݵĶ�������ʽ
 Public Sub GetData()
    Dim value, str, bcrlf, fpos, sSplit, slen, istartg
    Dim formend, formhead, startpos, endpos, formname, FileName, fileExe, valueend, NewName
    If checkEntryType = True Then
        vTotalSize = 0
        StreamT.Type = 1
        StreamT.Mode = 3
        StreamT.Open
        StreamT.Write Request.binaryread(Request.totalbytes)
        StreamT.Position = 0
        tempdata = StreamT.Read
        bcrlf = ChrB(13) & ChrB(10)
        fpos = InStrB(1, tempdata, bcrlf)
        sSplit = MidB(tempdata, 1, fpos - 1)
        slen = LenB(sSplit)
        istart = slen + 2
        Do
            formend = InStrB(istart, tempdata, bcrlf & bcrlf)
            formhead = MidB(tempdata, istart, formend - istart)
            str = Bytes2Str(formhead)
            startpos = InStr(str, "name=""") + 6
            endpos = InStr(startpos, str, """")
            formname = LCase(Mid(str, startpos, endpos - startpos))
            valueend = InStrB(formend + 3, tempdata, sSplit)
            If InStr(str, "filename=""") > 0 Then
                startpos = InStr(str, "filename=""") + 10
                endpos = InStr(startpos, str, """")
                FileName = Mid(str, startpos, endpos - startpos)
                If Trim(FileName) <> "" Then
                    LocalName = FileName
                    FileName = Replace(FileName, "/", "\")
                    FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
                    fileExe = Split(FileName, ".")(UBound(Split(FileName, ".")))
                    If vExe <> "" Then '�ж���չ��
                        If checkExe(fileExe) = True Then
                            vErr = 3
                            Exit Sub
                        End If
                    End If
                    NewName = Getname()
                    NewName = NewName & "." & fileExe
                    vTotalSize = vTotalSize + valueend - formend - 6
                    If vSingleSize > 0 And (valueend - formend - 6) > vSingleSize Then '�ж��ϴ������ļ���С
                        vErr = 5
                        Exit Sub
                    End If
                    If vMaxSize > 0 And vTotalSize > vMaxSize Then '�ж��ϴ������ܴ�С
                        vErr = 1
                        Exit Sub
                    End If
                    If Fils.Exists(formname) Then
                        vErr = 4
                        Exit Sub
                    Else
                        Dim fileCls
			set fileCls=New fileAction
                        fileCls.Size = (valueend - formend - 6)
                        fileCls.Position = (formend + 3)
                        fileCls.NewName = NewName
                        fileCls.LocalName = FileName
                        Fils.Add formname, fileCls
                        Form.Add formname, LocalName
                        Set fileCls = Nothing
                    End If
                End If
            Else
                value = MidB(tempdata, formend + 4, valueend - formend - 6)
                If Form.Exists(formname) Then
                    Form(formname) = Form(formname) & "," & Bytes2Str(value)
                Else
                    Form.Add formname, Bytes2Str(value)
                End If
            End If
            istart = valueend + 2 + slen
        Loop Until (istart + 2) >= LenB(tempdata)
        vErr = 0
   Else
        vErr = 2
   End If
 End Sub
 
 '�ж���չ��
 Private Function checkExe(ByVal ex)
      Dim notIn: notIn = True
      If InStr(1, vExe, "|") > 0 Then
           Dim tempExe: tempExe = Split(vExe, "|")
           Dim I: I = 0
           For I = 0 To UBound(tempExe)
                 If LCase(ex) = tempExe(I) Then
                       notIn = False
                       Exit For
                 End If
           Next
     Else
           If vExe = LCase(ex) Then
                notIn = False
           End If
     End If
     checkExe = notIn
 End Function
 
 '-------------------������ת��Ϊ�ļ���С��ʾ��ʽ
 Public Function GetSize(ByVal Size)
    If Size < 1024 Then
       GetSize = FormatNumber(Size, 2) & "B"
    ElseIf Size >= 1024 And Size < 1048576 Then
       GetSize = FormatNumber(Size / 1024, 2) & "KB"
    ElseIf Size >= 1048576 Then
       GetSize = FormatNumber((Size / 1024) / 1024, 2) & "MB"
    End If
 End Function

 '-------------------����������ת��Ϊ�ַ�
 Private Function Bytes2Str(ByVal byt)
    If LenB(byt) = 0 Then
    Bytes2Str = ""
    Exit Function
    End If
    Dim mystream, bstr
    Set mystream =server.createobject("ADODB.Stream")
    mystream.Type = 2
    mystream.Mode = 3
    mystream.Open
    mystream.WriteText byt
    mystream.Position = 0
    mystream.CharSet = vCharSet
    mystream.Position = 2
    bstr = mystream.ReadText()
    mystream.Close
    Set mystream = Nothing
    Bytes2Str = bstr
 End Function

 '-------------------��ȡ��������
 Private Function GetErr(ByVal Num)
    Select Case Num
      Case 0
        GetErr = "���ݴ������!"
      Case 1
        GetErr = "�ϴ����ݳ���" & GetSize(vMaxSize) & "����!������MaxSize�������ı�����!"
      Case 2
        GetErr = "δ�����ϴ���enctype����Ϊmultipart/form-data,�ϴ���Ч!"
      Case 3
        GetErr = "���зǷ���չ���ļ�!ֻ���ϴ���չ��Ϊ" & Replace(vExe, "|", ",") & "���ļ�"
      Case 4
        GetErr = "�Բ���,��������ʹ����ͬname���Ե��ļ���!"
      Case 5
        GetErr = "�����ļ���С����" & GetSize(vSingleSize) & "���ϴ�����!"
    End Select
 End Function

 '-------------------����������������ļ���
 Private Function Getname()
    Dim y, m, d, h, mm, S, r
    Randomize
    y = Year(Now)
    m = Month(Now): If m < 10 Then m = "0" & m
    d = Day(Now): If d < 10 Then d = "0" & d
    h = Hour(Now): If h < 10 Then h = "0" & h
    mm = Minute(Now): If mm < 10 Then mm = "0" & mm
    S = Second(Now): If S < 10 Then S = "0" & S
    r = 0
    r = CInt(Rnd() * 1000)
    If r < 10 Then r = "00" & r
    If r < 100 And r >= 10 Then r = "0" & r
    Getname = y & m & d & h & mm & S & r
 End Function
 
 '------------------------����ϴ������Ƿ�Ϊmultipart/form-data
 Private Function checkEntryType()
    Dim ContentType, ctArray, bArray
    ContentType = LCase(Request.ServerVariables("HTTP_CONTENT_TYPE"))
    ctArray = Split(ContentType, ";")
    If Trim(ctArray(0)) = "multipart/form-data" Then
        checkEntryType = True
    Else
        checkEntryType = False
    End If
 End Function
 
 '��ȡ�ϴ���ֵ,������ѡ,���Ϊ-1�򷵻�һ���������б����һ��dictionary����
 Public Function Forms(ByVal formname)
 If trim(formname) = "-1" Then
        Set Forms = Form
 Else
        If Form.Exists(LCase(formname)) Then
            Forms = Form(LCase(formname))
        Else
            Forms = ""
        End If
 End If
 End Function
 
'��ȡ�ϴ����ļ���,������ѡ,���Ϊ-1�򷵻�һ�����������ϴ��ļ����һ��dictionary����
 Public Function Files(ByVal formname)
  If trim(formname) = "-1" Then
        Set Files = Fils
 Else
        If Fils.Exists(LCase(formname)) Then
            Set Files = Fils(LCase(formname))
        Else
            Set Files = Nothing
        End If
 End If
 End Function
End Class


Class fileAction
   Private vSize, vPosition, vName, vNewName, vLocalName, vPath, saveName
   
   '��������
   Public Property Let NewName(ByVal value)
          vNewName = value
          vName = value
   End Property

   Public Property Let LocalName(ByVal value)
          vLocalName = value
   End Property

   Public Property Get FileName()
          FileName = vName
   End Property

   Public Property Let Position(ByVal value)
          vPosition = value
   End Property

   Public Property Let Size(ByVal value)
          vSize = value
   End Property
   Public Property Get Size()
          Size = vSize
   End Property
   
   '������:SaveToFile
   '����:���ݲ��������ļ���������
   '����:����1--�ļ������·��
   '     ����2--�ļ�����ķ�ʽ,��������ѡ��0��ʾ��������(ʱ��+�����)Ϊ�ļ�������,1��ʾ��ԭ�ļ��������ļ�
   Public Function SaveToFile(ByVal path, ByVal saveType)
    'On Error Resume Next
    Err.Clear
    vPath = Replace(path, "/", "\")
    If Right(vPath, 1) <> "\" Then vPath = vPath & "\"
    Dim mystream
    Set mystream =server.createobject("ADODB.Stream")
    mystream.Type = 1
    mystream.Mode = 3
    mystream.Open
    StreamT.Position = vPosition
    StreamT.CopyTo mystream, vSize
    vName = vNewName
    If saveType = 1 Then vName = vLocalName
    mystream.SaveToFile vPath & vName, 2
    mystream.Close
    Set mystream = Nothing
    If Err Then
        SaveToFile = False
    Else
        SaveToFile = True
    End If
   End Function

   '������:GetBytes
   '����:��ȡ�ļ��Ķ�������ʽ
   '����:��
   Public Function GetBytes()
    StreamT.Position = vPosition
    GetBytes = StreamT.Read(vSize)
  End Function

End Class
%>