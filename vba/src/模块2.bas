Attribute VB_Name = "模块2"
Option Explicit

' ====== OK 自建应用 PPT 模块 ======
' 应用：笔记
' 生成时间：2026/7/31 12:20:48

' 获取登录用户 ID（从 PPT 登录插件模块读取，无需手动设置）
Private Function OKApp_GetUserId() As String
    On Error Resume Next
    OKApp_GetUserId = Application.Run("GetOKUserId")
    On Error GoTo 0
End Function

' 确保用户已登录（各函数内部自动调用，无需手动执行）
Private Function OKApp_EnsureLogin() As Boolean
    Dim uid As String
    uid = OKApp_GetUserId()
    uid = Trim(uid)
    If uid = "" Or uid = "0" Then
        MsgBox "请先运行 PPT 登录插件完成登录", vbExclamation
        OKApp_EnsureLogin = False
    Else
        OKApp_EnsureLogin = True
    End If
End Function

' 检查当前登录状态
Public Sub AppCheckLogin()
    Dim uid As String
    uid = OKApp_GetUserId()
    uid = Trim(uid)
    If uid = "" Or uid = "0" Then
        MsgBox "未登录，请先运行 PPT 登录插件", vbExclamation
    Else
        MsgBox "已登录，用户 ID：" & uid, vbInformation
    End If
End Sub

' ====== 形状/控件读写兼容函数 ======

' 查找指定名称的形状/控件（跨所有幻灯片）
Private Function OKApp_FindShape(ByVal shapeName As String) As Object
    On Error Resume Next
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        Dim shp As Object
        Set shp = sld.Shapes(shapeName)
        If Err.Number = 0 Then
            Set OKApp_FindShape = shp
            Exit Function
        End If
        Err.Clear
    Next sld
    Set OKApp_FindShape = Nothing
    On Error GoTo 0
End Function

' 获取形状/控件的文本（兼容 ActiveX 文本框控件和普通形状）
Private Function OKApp_GetShapeText(ByRef shp As Object) As String
    On Error Resume Next
    OKApp_GetShapeText = shp.OLEFormat.Object.Text
    If Err.Number <> 0 Then
        Err.Clear
        OKApp_GetShapeText = shp.TextFrame.TextRange.Text
    End If
    On Error GoTo 0
End Function

' 设置形状/控件的文本（兼容 ActiveX 文本框控件和普通形状）
Private Sub OKApp_SetShapeText(ByRef shp As Object, ByVal txt As String)
    On Error Resume Next
    shp.OLEFormat.Object.Text = txt
    If Err.Number <> 0 Then
        Err.Clear
        shp.TextFrame.TextRange.Text = txt
    End If
    On Error GoTo 0
End Sub

' 读取当前用户的云端数据（静默写入文本框，仅报错弹窗）
Public Sub AppLoadData()
    On Error GoTo ErrHandler
    If Not OKApp_EnsureLogin() Then Exit Sub
    Dim uid As String
    uid = OKApp_GetUserId()
    Dim dataKey As String
    dataKey = "note"
    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://www.okteam.cn/app/api/data?app_key=" & PPTSec_Key("data") & "&user_id=" & uid & "&key=" & dataKey
    xmlhttp.Open "GET", url, False
    Call WaitOn
    xmlhttp.Send
    Call WaitOff

    If xmlhttp.Status = 404 Then
        ' 数据不存在，静默退出（不清空文本框）
        Exit Sub
    End If
    If xmlhttp.Status <> 200 Then
        MsgBox "读取失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Sub
    End If
    json = xmlhttp.responseText
    If InStr(json, """success"":true") = 0 Then
        MsgBox "读取失败：" & OKApp_ExtractJsonStr(json, "error"), vbExclamation
        Exit Sub
    End If
    Dim value As String
    value = OKApp_ExtractJsonStr(json, "value")
    Dim shp As Object
    Set shp = OKApp_FindShape("CloudNote")
    If shp Is Nothing Then
        MsgBox "找不到文本框控件：CloudNote", vbExclamation
        Exit Sub
    End If
    Call OKApp_SetShapeText(shp, value)
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "读取出错：" & Err.Description, vbCritical
End Sub

' 保存当前用户的云端数据（直接从文本框读取，仅报错弹窗）
Public Sub AppSaveData()
    On Error GoTo ErrHandler
    If Not OKApp_EnsureLogin() Then Exit Sub
    Dim uid As String
    uid = OKApp_GetUserId()
    Dim dataKey As String
    dataKey = "note"
    Dim value As String
    value = ""
    Dim shp As Object
    Set shp = OKApp_FindShape("CloudNote")
    If shp Is Nothing Then
        MsgBox "找不到文本框控件：CloudNote", vbExclamation
        Exit Sub
    End If
    value = OKApp_GetShapeText(shp)
    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", "https://www.okteam.cn/app/api/data", False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "X-PPTOS-Client", PPTSec_ClientTag()
    Dim body As String
    body = "{""app_key"":""" & PPTSec_Key("data") & """,""user_id"":""" & uid & """,""key"":""" & dataKey & """,""value"":""" & OKApp_JsonEscape(value) & """}"
    Call WaitOn
    xmlhttp.Send body
    Call WaitOff

    If xmlhttp.Status <> 200 Then
        MsgBox "保存失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Sub
    End If
    json = xmlhttp.responseText
    If InStr(json, """success"":true") = 0 Then
        MsgBox "保存失败：" & OKApp_ExtractJsonStr(json, "error"), vbExclamation
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "保存出错：" & Err.Description, vbCritical
End Sub

' 读取公共数据（核心函数，内部调用，无需登录，静默写入）
Private Sub OKApp_FetchPublic(ByVal dataKey As String, ByVal targetShape As String)
    On Error GoTo ErrHandler
    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", "https://www.okteam.cn/app/api/public?app_key=" & PPTSec_Key("data") & "&key=" & dataKey, False
    Call WaitOn
    xmlhttp.Send
    Call WaitOff

    If xmlhttp.Status = 404 Then
        ' 数据不存在，静默退出
        Exit Sub
    End If
    If xmlhttp.Status <> 200 Then
        MsgBox "读取公共数据失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Sub
    End If
    json = xmlhttp.responseText
    Dim value As String
    value = OKApp_ExtractJsonStr(json, "value")
    If targetShape <> "" Then
        Dim shp As Object
        Set shp = OKApp_FindShape(targetShape)
        If shp Is Nothing Then
            MsgBox "找不到文本框控件：" & targetShape, vbExclamation
            Exit Sub
        End If
        Call OKApp_SetShapeText(shp, value)
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "读取出错：" & Err.Description, vbCritical
End Sub

' 读取所有已配置的公共数据（无需输入键名，直接显示到对应形状）
Public Sub AppLoadPublic()
    MsgBox "尚未配置公共数据键名，请在代码配置卡片中勾选需要的公共数据项", vbExclamation
End Sub

' 检查 PPT 版本更新
' 前置条件：
'   1. 管理后台「公共数据」设置 latest_version = 最新版本号
'   2. 管理后台「公共数据」设置 download_url = 下载链接
'   3. PPT 中创建文本框命名为 CurrentVersion，填入当前版本号
Public Sub AppCheckUpdate()
    On Error GoTo ErrHandler
    ' 1) 读取 PPT 文本框 CurrentVersion 中的当前版本号
    Dim localVer As String
    localVer = ""
    Dim cvShp As Object
    Set cvShp = OKApp_FindShape("CurrentVersion")
    If cvShp Is Nothing Then
        MsgBox "未找到命名为 CurrentVersion 的文本框，请先创建并填入当前版本号", vbExclamation
        Exit Sub
    End If
    localVer = Trim(OKApp_GetShapeText(cvShp))
    If localVer = "" Then
        MsgBox "文本框 CurrentVersion 为空，请先填入当前 PPT 版本号", vbExclamation
        Exit Sub
    End If

    ' 2) 读取公共数据 latest_version（云端最新版本号）
    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", "https://www.okteam.cn/app/api/public?app_key=" & PPTSec_Key("data") & "&key=latest_version", False
    Call WaitOn
    xmlhttp.Send
    Call WaitOff

    If xmlhttp.Status = 404 Then
        MsgBox "未找到云端版本号，请先在管理后台设置公共数据 latest_version", vbExclamation
        Exit Sub
    End If
    If xmlhttp.Status <> 200 Then
        MsgBox "检查更新失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Sub
    End If
    json = xmlhttp.responseText
    Dim cloudVer As String
    cloudVer = OKApp_ExtractJsonStr(json, "value")
    If cloudVer = "" Then
        MsgBox "云端版本号为空，请在管理后台设置 latest_version 的值", vbExclamation
        Exit Sub
    End If

    ' 3) 比较版本号
    If OKApp_CompareVersion(localVer, cloudVer) < 0 Then
        ' 有新版本，读取 download_url 并提示用户一键下载
        Dim dlUrl As String
        dlUrl = ""
        xmlhttp.Open "GET", "https://www.okteam.cn/app/api/public?app_key=" & PPTSec_Key("data") & "&key=download_url", False
        Call WaitOn
        xmlhttp.Send
        Call WaitOff
        If xmlhttp.Status = 200 Then
            dlUrl = OKApp_ExtractJsonStr(xmlhttp.responseText, "value")
        End If
        Dim tip As String
        tip = "发现新版本！" & vbCrLf & "当前版本：" & localVer & vbCrLf & "最新版本：" & cloudVer & vbCrLf
        If dlUrl <> "" Then
            tip = tip & vbCrLf & "是否立即下载更新？"
            Dim ret As Integer
            ret = MsgBox(tip, vbYesNo + vbQuestion, "检查更新")
            If ret = vbYes Then
                ActivePresentation.FollowHyperlink Address:=dlUrl, NewWindow:=True
            End If
        Else
            MsgBox tip & vbCrLf & "（未配置下载链接，请在管理后台设置 download_url）", vbExclamation, "检查更新"
        End If
    Else
        MsgBox "当前已为最新版本" & vbCrLf & "版本号：" & localVer, vbInformation, "检查更新"
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "检查更新出错：" & Err.Description, vbCritical
End Sub

' 版本号比较（支持 x.y.z 格式，返回 -1/0/1）
Private Function OKApp_CompareVersion(v1 As String, v2 As String) As Integer
    Dim a1() As String, a2() As String
    a1 = Split(v1, ".")
    a2 = Split(v2, ".")
    Dim i As Integer, n As Integer
    n = IIf(UBound(a1) > UBound(a2), UBound(a1), UBound(a2))
    For i = 0 To n
        Dim n1 As Long, n2 As Long
        n1 = 0: n2 = 0
        If i <= UBound(a1) Then n1 = Val(a1(i))
        If i <= UBound(a2) Then n2 = Val(a2(i))
        If n1 < n2 Then
            OKApp_CompareVersion = -1
            Exit Function
        ElseIf n1 > n2 Then
            OKApp_CompareVersion = 1
            Exit Function
        End If
    Next i
    OKApp_CompareVersion = 0
End Function

' ====== JSON 解析工具函数 ======

Private Function OKApp_ExtractJsonStr(json As String, key As String) As String
    Dim searchStr As String, startPos As Long, endPos As Long
    searchStr = """" & key & """:"""
    startPos = InStr(json, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            OKApp_ExtractJsonStr = Mid(json, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    searchStr = """" & key & """:"
    startPos = InStr(json, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        Dim commaPos As Long, bracePos As Long
        commaPos = InStr(startPos, json, ",")
        bracePos = InStr(startPos, json, "}")
        If commaPos = 0 Then commaPos = 99999
        If bracePos = 0 Then bracePos = 99999
        endPos = IIf(commaPos < bracePos, commaPos, bracePos)
        If endPos > startPos Then
            OKApp_ExtractJsonStr = Trim(Mid(json, startPos, endPos - startPos))
            Exit Function
        End If
    End If
    OKApp_ExtractJsonStr = ""
End Function

Private Function OKApp_JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    OKApp_JsonEscape = s
End Function

