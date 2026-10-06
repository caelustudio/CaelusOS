Attribute VB_Name = "模块5"
Option Explicit

' ====== OK AI 对话 PPT 模块 ======
' 应用：Star Intelligence
' 生成时间：2026/8/11 14:26:28
' 平台不存储对话内容
' 用户需在个人中心→模型管理中配置模型API
' 支持连续对话：上下文保存在 PPT 内存中，关闭文件后清空

' ====== 对话上下文（模块级，仅保存在内存中） ======
Private Const OKAI_MAX_CONTEXT As Long = 50
Private OKAI_Roles() As String
Private OKAI_Contents() As String
Private OKAI_Count As Long
Private OKAI_InitDone As Boolean

' 初始化上下文数组
Private Sub OKAI_Init()
    If OKAI_InitDone Then Exit Sub
    ReDim OKAI_Roles(1 To OKAI_MAX_CONTEXT)
    ReDim OKAI_Contents(1 To OKAI_MAX_CONTEXT)
    OKAI_Count = 0
    OKAI_InitDone = True
End Sub

' 追加一条消息到上下文（超出上限时滚动丢弃最早的消息）
Private Sub OKAI_Push(ByVal role As String, ByVal content As String)
    OKAI_Init
    If OKAI_Count >= OKAI_MAX_CONTEXT Then
        Dim i As Long
        For i = 2 To OKAI_MAX_CONTEXT
            OKAI_Roles(i - 1) = OKAI_Roles(i)
            OKAI_Contents(i - 1) = OKAI_Contents(i)
        Next i
        OKAI_Count = OKAI_MAX_CONTEXT - 1
    End If
    OKAI_Count = OKAI_Count + 1
    OKAI_Roles(OKAI_Count) = role
    OKAI_Contents(OKAI_Count) = content
End Sub

' 构建请求体中的 messages JSON 片段
Private Function OKAI_BuildMessagesJson() As String
    OKAI_Init
    Dim sb As String, i As Long
    sb = "["
    For i = 1 To OKAI_Count
        If i > 1 Then sb = sb & ","
        sb = sb & "{""role"":""" & OKAI_Roles(i) & """,""content"":""" & OKApp_JsonEscape(OKAI_Contents(i)) & """}"
    Next i
    sb = sb & "]"
    OKAI_BuildMessagesJson = sb
End Function

' 获取登录用户 ID（从 PPT 登录插件模块读取）
Private Function OKApp_GetUserId() As String
    On Error Resume Next
    OKApp_GetUserId = Application.Run("GetOKUserId")
    On Error GoTo 0
End Function

' 确保用户已登录
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

' 查找指定名称的形状/控件
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

' 获取形状/控件文本
Private Function OKApp_GetShapeText(ByRef shp As Object) As String
    On Error Resume Next
    OKApp_GetShapeText = shp.OLEFormat.Object.Text
    If Err.Number <> 0 Then
        Err.Clear
        OKApp_GetShapeText = shp.TextFrame.TextRange.Text
    End If
    On Error GoTo 0
End Function

' 设置形状/控件文本
Private Sub OKApp_SetShapeText(ByRef shp As Object, ByVal txt As String)
    On Error Resume Next
    shp.OLEFormat.Object.Text = txt
    If Err.Number <> 0 Then
        Err.Clear
        shp.TextFrame.TextRange.Text = txt
    End If
    On Error GoTo 0
End Sub

' JSON 字符串转义
Private Function OKApp_JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    OKApp_JsonEscape = s
End Function

' 提取 JSON 字符串值（逐字符扫描，正确处理转义）
Private Function OKApp_ExtractJsonStr(json As String, key As String) As String
    Dim searchStr As String, startPos As Long, endPos As Long
    searchStr = """" & key & """:"""
    startPos = InStr(json, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        Dim pp As Long
        pp = startPos
        Do While pp <= Len(json)
            If Asc(Mid$(json, pp, 1)) = 92 Then
                pp = pp + 2
            ElseIf Asc(Mid$(json, pp, 1)) = 34 Then
                endPos = pp
                Exit Do
            Else
                pp = pp + 1
            End If
        Loop
        If endPos > startPos Then
            Dim raw As String
            raw = Mid(json, startPos, endPos - startPos)
            raw = Replace(raw, Chr$(92) & Chr$(34), Chr$(34))
            raw = Replace(raw, Chr$(92) & Chr$(92), Chr$(92))
            raw = Replace(raw, "\n", vbCrLf)
            raw = Replace(raw, "\t", vbTab)
            OKApp_ExtractJsonStr = raw
            Exit Function
        End If
    End If
    OKApp_ExtractJsonStr = ""
End Function

' ====== AI 对话核心函数 ======

' AI 对话（连续对话）：从输入框读取问题 → 调用平台API → 将回复写入输出框
' 平台根据当前登录用户自动路由到其配置的模型API
' 上下文保存在模块内存中，关闭文件后自动清空；可运行 AppAiChatClear 主动清空
Public Sub AppAiChat()
    On Error GoTo ErrHandler
    If Not OKApp_EnsureLogin() Then Exit Sub
    ' 安全闸门：校验登录态 + 每日调用配额（通过即预占一次，失败会退还）
    If Not PPTSec_CanCall("ai") Then Exit Sub

    Dim inputShp As Object, outputShp As Object
    Set inputShp = OKApp_FindShape("TextBox1")
    Set outputShp = OKApp_FindShape("TextBox")
    If inputShp Is Nothing Then
        MsgBox "未找到输入文本框：TextBox1", vbExclamation
        Exit Sub
    End If
    If outputShp Is Nothing Then
        MsgBox "未找到输出文本框：TextBox", vbExclamation
        Exit Sub
    End If

    Dim userInput As String
    userInput = OKApp_GetShapeText(inputShp)
    userInput = Trim(userInput)
    If Len(userInput) = 0 Then
        MsgBox "请在输入框中输入问题", vbExclamation
        Exit Sub
    End If

    Dim uid As String
    uid = OKApp_GetUserId()

    ' 追加用户消息到上下文（请求前）
    OKAI_Push "user", userInput

    Dim messagesJson As String
    messagesJson = OKAI_BuildMessagesJson()

    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", "https://www.okteam.cn/app/api/ai-chat", False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "X-PPTOS-Client", PPTSec_ClientTag()
    Dim body As String
    body = "{""app_key"":""" & PPTSec_Key("ai") & """,""user_id"":""" & uid & """,""messages"":" & messagesJson & "}"
    Call WaitOn
    xmlhttp.Send body
    Call WaitOff

    If xmlhttp.Status <> 200 Then
        ' 请求失败时回滚上下文
        If OKAI_Count > 0 Then OKAI_Count = OKAI_Count - 1
        Call PPTSec_Commit("ai", False, "http" & xmlhttp.Status)
        MsgBox "请求失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Sub
    End If

    json = xmlhttp.responseText
    If InStr(json, """success"":true") > 0 Or InStr(json, """ok"":true") > 0 Then
        Dim reply As String
        reply = OKApp_ExtractJsonStr(json, "reply")
        If reply = "" Then reply = OKApp_ExtractJsonStr(json, "content")
        If reply = "" Then reply = "(模型返回空内容)"
        ' 追加助手回复到上下文（供下一轮对话使用）
        OKAI_Push "assistant", reply
        OKApp_SetShapeText outputShp, reply
        Call PPTSec_Commit("ai", True, "ok")
    Else
        ' 失败时回滚上下文
        If OKAI_Count > 0 Then OKAI_Count = OKAI_Count - 1
        Dim errMsg As String
        errMsg = OKApp_ExtractJsonStr(json, "error")
        If errMsg = "" Then errMsg = "请求失败"
        If InStr(json, "no_model") > 0 Then
            Call PPTSec_Commit("ai", False, "no_model")
            MsgBox "尚未配置AI模型。" & vbCrLf & vbCrLf & "请登录 OK 平台，在个人中心 → 模型管理中配置您的模型API后重试。", vbInformation
        Else
            Call PPTSec_Commit("ai", False, Left$(errMsg, 80))
            MsgBox errMsg, vbExclamation
        End If
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
    ' 异常时回滚上下文
    If OKAI_Count > 0 Then OKAI_Count = OKAI_Count - 1
    Call PPTSec_Commit("ai", False, "exception")
    MsgBox "AI对话出错：" & Err.Description, vbCritical
End Sub

' 清空 AI 对话上下文，从新会话开始
' 运行此宏后，下一次 AppAiChat 将不再带历史消息
Public Sub AppAiChatClear()
    OKAI_Init
    OKAI_Count = 0
    MsgBox "AI 对话上下文已清空，下次对话将从新会话开始。", vbInformation
End Sub

