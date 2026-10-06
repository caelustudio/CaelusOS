Attribute VB_Name = "模块1"
Option Explicit

' ===== 网络请求等待提示（等待期间鼠标转圈） =====
#If VBA7 Then
    Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
        ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
    Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
        ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If
Private Const IDC_WAIT As Long = 32514
Private Const IDC_ARROW As Long = 32512
Public Const SLIDE_LOGIN_DONE_ID As Long = 348
Public Const SLIDE_NOTES_PAGE_ID As Long = 290

' 一键登录：等待剪贴板验证码的最长秒数
Private Const PPTSEC_LOGIN_WAIT_SEC As Long = 120

' 用于打开浏览器的 API 声明（必须放在模块顶部）
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

' 用户登录信息（Public 变量，供自建应用模块通过 Application.Run 读取）
Public gOKUserId As String
Public gOKUserNickname As String

' 获取登录用户 ID（供自建应用模块调用）
Public Function GetOKUserId() As String
    GetOKUserId = gOKUserId
End Function

' 获取登录用户昵称（供自建应用模块调用）
Public Function GetOKUserNickname() As String
    GetOKUserNickname = gOKUserNickname
End Function

' 清除登录信息（结束放映或重新打开 PPT 时变量自动清空）
Public Sub ClearOKLogin()
    gOKUserId = ""
    gOKUserNickname = ""
End Sub

' ===== 登录入口（一键登录） =====
' 点击后：打开浏览器 → 自动等待剪贴板里出现 6 位验证码 → 自动完成登录
' 若 120 秒内没有检测到验证码，回退为手动输入。
Sub LoginOK()
    On Error GoTo ErrHandler
    Dim loginUrl As String
    loginUrl = "https://www.okteam.cn/ppt/login?k=" & PPTSec_Key("login")

    ' 1) 打开登录页
    On Error Resume Next
    ShellExecute 0, "open", loginUrl, vbNullString, vbNullString, vbNormalFocus
    On Error GoTo ErrHandler

    ' 2) 若剪贴板里已有验证码（先复制再点按钮），直接登录
    Dim code As String
    code = PPTOS_ExtractCode(PPTOS_ClipboardText())

    ' 3) 自动等待：用户复制验证码后无需再操作
    If Len(code) <> 6 Then
        Dim t0 As Double
        t0 = PPTOS_NowSec()
        Do While PPTOS_Elapsed(t0) < PPTSEC_LOGIN_WAIT_SEC
            DoEvents
            code = PPTOS_ExtractCode(PPTOS_ClipboardText())
            If Len(code) = 6 Then Exit Do
            Call PPTOS_SleepMs(400)
        Loop
    End If

    ' 4) 超时回退：手动输入
    If Len(code) <> 6 Then
        code = InputBox("未能自动读取验证码。" & vbCrLf & _
                        "请在浏览器中完成登录后，" & vbCrLf & _
                        "复制页面上的 6 位验证码，再点确定：", "OK 账户登录", "")
        If code = "" Then
            MsgBox "已取消登录", vbInformation
            Exit Sub
        End If
        code = PPTOS_ExtractCode(code)
    End If

    If Len(code) <> 6 Then
        MsgBox "验证码格式错误，应为 6 位字母数字组合", vbExclamation
        Exit Sub
    End If

    ' 5) 换取用户信息
    Call GetUserInfo(code)
    Exit Sub
ErrHandler:
    MsgBox "登录过程出错：" & Err.Description, vbCritical
End Sub

' 手动登录（保留旧流程，供需要时使用）
Sub LoginOKManual()
    Dim loginUrl As String
    loginUrl = "https://www.okteam.cn/ppt/login?k=" & PPTSec_Key("login")
    On Error Resume Next
    ShellExecute 0, "open", loginUrl, vbNullString, vbNullString, vbNormalFocus
    On Error GoTo 0

    Dim code As String
    code = InputBox("请在浏览器中完成登录后，" & vbCrLf & "输入页面上显示的 6 位验证码：", "OK 账户登录", "")
    If code = "" Then
        MsgBox "已取消登录", vbInformation
        Exit Sub
    End If
    code = UCase(Trim(code))
    If Len(code) <> 6 Then
        MsgBox "验证码格式错误，应为 6 位字母数字组合", vbExclamation
        Exit Sub
    End If
    Call GetUserInfo(code)
End Sub

' 调用 API 获取用户信息
Private Sub GetUserInfo(code As String)
    On Error GoTo ErrHandler
    Dim xmlhttp As Object
    Dim apiUrl As String
    apiUrl = "https://www.okteam.cn/ppt/api/user-info?code=" & code

    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", apiUrl, False
    Call WaitOn
    xmlhttp.Send
    Call WaitOff

    If xmlhttp.Status = 200 Then
        Dim response As String
        response = xmlhttp.responseText

        ' 检查是否成功
        If InStr(response, """success"":true") > 0 Then
            ' 提取 token（用于头像接口调用）
            Dim token As String
            token = ExtractJsonValue(response, "token")

            ' 提取 user 对象内的字段（先截取 user 对象内容）
            Dim userJson As String
            userJson = ExtractUserObject(response)

            ' 提取昵称
            Dim nickname As String
            nickname = ExtractJsonValue(userJson, "nickname")

            ' 提取用户 ID
            Dim userId As String
            userId = ExtractJsonValue(userJson, "id")

            ' 记录用户登录信息到 Public 变量（供自建应用模块通过 Application.Run 读取）
            gOKUserId = userId
            gOKUserNickname = nickname

            ' 提取角色
            Dim roleLabel As String
            roleLabel = ExtractJsonValue(userJson, "role_label")

            ' ===== 将信息写入所有幻灯片 =====
            Dim sld As Slide
            Dim foundAny As Boolean
            foundAny = False
            ' 构造头像 URL（所有幻灯片共用）
            Dim avatarImgUrl As String
            avatarImgUrl = "https://www.okteam.cn/ppt/api/avatar?token=" & token
            For Each sld In ActivePresentation.Slides

                ' 设置昵称
                If "UserName" <> "" Then
                    On Error Resume Next
                    sld.Shapes("UserName").TextFrame.TextRange.Text = nickname
                    If Err.Number = 0 Then foundAny = True
                    Err.Clear
                    On Error GoTo ErrHandler
                End If

                ' 设置用户ID
                If "UserID" <> "" Then
                    On Error Resume Next
                    sld.Shapes("UserID").TextFrame.TextRange.Text = userId
                    If Err.Number = 0 Then foundAny = True
                    Err.Clear
                    On Error GoTo ErrHandler
                End If

                ' 设置头像
                If "UserTX" <> "" And token <> "" Then
                    On Error Resume Next
                    Dim avatarShapeObj As Object
                    Set avatarShapeObj = sld.Shapes("UserTX")
                    If Err.Number = 0 Then
                        Call SetShapeImage(avatarShapeObj, avatarImgUrl)
                        foundAny = True
                    End If
                    Err.Clear
                    On Error GoTo ErrHandler
                End If

            Next sld

            ' 跳转到指定幻灯片（仅在放映模式下）
            On Error Resume Next
            If ActivePresentation.SlideShowWindow.View.State <> 0 Then
                Call GotoNamedSlide("LoginDone")
            End If
            Err.Clear
            On Error GoTo ErrHandler

            If Not foundAny Then
                MsgBox "未找到任何匹配的形状名称，请检查配置是否正确", vbExclamation
            Else
                MsgBox "登录成功！" & vbCrLf & _
                       "昵称：" & nickname & vbCrLf & _
                       "ID：" & userId & vbCrLf & _
                       "角色：" & roleLabel, vbInformation
            End If
        Else
            ' 提取错误信息
            Dim errMsg As String
            errMsg = ExtractJsonValue(response, "error")
            If errMsg = "" Then errMsg = "登录失败，请检查验证码是否正确"
            MsgBox errMsg, vbExclamation
        End If
    Else
        MsgBox "网络请求失败，状态码：" & xmlhttp.Status, vbCritical
    End If

    Set xmlhttp = Nothing
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "发生错误：" & Err.Description & vbCrLf & _
           "错误号：" & Err.Number, vbCritical
    Set xmlhttp = Nothing
End Sub

' 通过 URL 设置形状图片（保留动画、组合、格式等所有属性）
Private Sub SetShapeImage(shape As Object, imageUrl As String)
    On Error GoTo ErrHandler
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\ok_ppt_avatar.jpg"

    ' 下载图片到临时文件
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", imageUrl, False
    Call WaitOn
    xmlhttp.Send
    Call WaitOff

    If xmlhttp.Status = 200 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1  ' adTypeBinary
        stream.Mode = 3  ' adModeReadWrite
        stream.Open
        stream.Write xmlhttp.responseBody
        stream.SaveToFile tempPath, 2  ' adSaveCreateOverWrite
        stream.Close
        Set stream = Nothing

        ' 检查文件是否下载成功
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempPath) Then
            If fso.GetFile(tempPath).Size > 100 Then

                ' 使用 Fill.UserPicture 设置图片为形状填充
                ' 保留形状的：动画、组合关系、名称、位置、大小、格式等所有属性
                On Error Resume Next
                shape.Fill.Visible = msoTrue
                shape.Fill.UserPicture tempPath
                shape.Fill.Tile = msoFalse
                shape.Fill.Transparency = 0
                Err.Clear
                On Error GoTo ErrHandler

            End If
        End If
        Set fso = Nothing

        ' 不删除临时文件，避免PPT还未加载完成就被删除
        ' （文件在系统临时目录，系统会自动清理）
    End If
    Set xmlhttp = Nothing
    Exit Sub
ErrHandler:
    Call WaitOff
    MsgBox "设置头像失败：" & Err.Description, vbExclamation
    Set xmlhttp = Nothing
End Sub

' 从完整 JSON 中提取 user 对象的内容
Private Function ExtractUserObject(json As String) As String
    Dim startPos As Long
    startPos = InStr(json, """user"":{")
    If startPos > 0 Then
        startPos = startPos + 7  ' 跳过 "user":{
        ' 找到匹配的结束 }（简单实现：找到下一个 } 作为结尾）
        ' 由于 user 对象是第一个嵌套对象且内部不含子对象，直接找下一个 } 即可
        Dim endPos As Long
        endPos = InStr(startPos, json, "}")
        If endPos > startPos Then
            ExtractUserObject = Mid(json, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    ' 如果找不到，返回原 JSON（兼容旧格式）
    ExtractUserObject = json
End Function

' 简单提取 JSON 字段值
Private Function ExtractJsonValue(json As String, key As String) As String
    Dim searchStr As String
    searchStr = """" & key & """:"""

    Dim startPos As Long
    startPos = InStr(json, searchStr)

    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        Dim endPos As Long
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            ExtractJsonValue = Mid(json, startPos, endPos - startPos)
            Exit Function
        End If
    End If

    ' 尝试数字类型（不带引号）
    searchStr = """" & key & """:"
    startPos = InStr(json, searchStr)
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        Dim endChar As String
        endChar = Mid(json, startPos, 1)
        If endChar <> """" Then
            ' 找到逗号或右大括号
            Dim commaPos As Long
            Dim bracePos As Long
            commaPos = InStr(startPos, json, ",")
            bracePos = InStr(startPos, json, "}")
            If commaPos = 0 Then commaPos = 99999
            If bracePos = 0 Then bracePos = 99999
            endPos = IIf(commaPos < bracePos, commaPos, bracePos)
            If endPos > startPos Then
                ExtractJsonValue = Trim(Mid(json, startPos, endPos - startPos))
                Exit Function
            End If
        End If
    End If

    ExtractJsonValue = ""
End Function

' ===== 等待提示 =====
Public Sub WaitOn()
    On Error Resume Next
    Call SetCursor(LoadCursor(0, IDC_WAIT))
    DoEvents
End Sub

Public Sub WaitOff()
    On Error Resume Next
    Call SetCursor(LoadCursor(0, IDC_ARROW))
    DoEvents
End Sub

' ===== 幻灯片命名跳转 =====
Private Sub EnsureSlideNames()
    On Error Resume Next
    Dim s As Slide
    Set s = ActivePresentation.Slides.FindBySlideID(SLIDE_LOGIN_DONE_ID)
    If Not s Is Nothing Then s.Name = "LoginDone"
    Set s = ActivePresentation.Slides.FindBySlideID(SLIDE_NOTES_PAGE_ID)
    If Not s Is Nothing Then s.Name = "NotesPage"
End Sub

Private Function SlideIdByName(ByVal targetName As String) As Long
    Select Case targetName
        Case "LoginDone"
            SlideIdByName = SLIDE_LOGIN_DONE_ID
        Case "NotesPage"
            SlideIdByName = SLIDE_NOTES_PAGE_ID
        Case Else
            SlideIdByName = 0
    End Select
End Function

Public Sub GotoNamedSlide(ByVal targetName As String)
    On Error Resume Next
    Call EnsureSlideNames

    Dim tgt As Slide
    Set tgt = Nothing

    Dim i As Long
    For i = 1 To ActivePresentation.Slides.Count
        If StrComp(ActivePresentation.Slides(i).Name, targetName, vbTextCompare) = 0 Then
            Set tgt = ActivePresentation.Slides(i)
            Exit For
        End If
    Next i

    If tgt Is Nothing Then
        Dim sid As Long
        sid = SlideIdByName(targetName)
        If sid > 0 Then Set tgt = ActivePresentation.Slides.FindBySlideID(sid)
    End If
    If tgt Is Nothing Then Exit Sub

    If SlideShowWindows.Count > 0 Then
        SlideShowWindows(1).View.GotoSlide tgt.SlideIndex
    Else
        ActivePresentation.SlideShowWindow.View.GotoSlide tgt.SlideIndex
    End If
End Sub

' ======================================================================
' ===== PPTOS 安全与配额（PPTSec） =====
' 作用：
'   1) app_key 不再以明文出现在代码里，改为运行时 XOR 还原，抬高被直接抄走门槛
'   2) 每次联网调用前校验登录态 + 每日调用配额，失败自动退还额度
'   3) 本地审计日志：谁、什么时候、调了什么、结果如何
'   4) 请求头带客户端标识，便于平台侧将来做来源校验
' 说明：这是客户端侧的第一道防线，能挡住误触/狂刷和随手抄 key；
'       要真正防住恶意调用，仍需 OK 平台在服务端按 user_id + app_key 做配额。
' ======================================================================

' 每日调用配额（按 OK 用户 ID 分别计数）
Private Const PPTSEC_LIMIT_AI As Long = 40
Private Const PPTSEC_LIMIT_CALC As Long = 500
Private Const PPTSEC_LIMIT_OTHER As Long = 200

' 取用户 ID（统一入口）
Public Function PPTSec_Uid() As String
    On Error Resume Next
    PPTSec_Uid = Trim$(CStr(Application.Run("GetOKUserId")))
    On Error GoTo 0
End Function

' 还原加密存储的 key（XOR 90）
Private Function PPTSec_Decode(ByVal cipher As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(cipher) Step 2
        s = s & Chr$(CLng("&H" & Mid$(cipher, i, 2)) Xor 90)
    Next i
    PPTSec_Decode = s
End Function

' 取平台 key（data=云数据 / ai=对话 / calc=计算器 / login=登录页）
Public Function PPTSec_Key(ByVal which As String) As String
    Select Case LCase$(Trim$(which))
        Case "data"
            PPTSec_Key = PPTSec_Decode("3B3105623B6C696F6B6C3E3C686E3B6E3E3B633F3E633F6E383F63")
        Case "ai"
            PPTSec_Key = PPTSec_Decode("3B31053B3B3B6C3B6A383F6C3968396D6D636939686A6A3E6E633C")
        Case "calc"
            PPTSec_Key = PPTSec_Decode("3B31053E3869696E6A3C6B636338383F6B6C696F6E3C626F3E6B6B")
        Case "login"
            PPTSec_Key = PPTSec_Decode("6F693C683E633F6B6C68626F6B6F633C3E6F6A3C6D623E396D6F6D3B623E6B38")
        Case Else
            PPTSec_Key = ""
    End Select
End Function

' 客户端标识：给平台侧留的来源标记（不含任何密钥）
Public Function PPTSec_ClientTag() As String
    Dim uid As String, h As Long, i As Long
    uid = PPTSec_Uid()
    For i = 1 To Len(uid)
        h = (h * 31 + Asc(Mid$(uid, i, 1))) Mod 100000000
    Next i
    PPTSec_ClientTag = "pptos/1.1/" & Format$(Date, "yyyymmdd") & "/" & Right$("00000000" & CStr(h), 8)
End Function

' ---- 本地文件 ----
Private Function PPTSec_QuotaFile() As String
    PPTSec_QuotaFile = Environ("TEMP") & "\pptos_quota.dat"
    If PPTSec_QuotaFile = "\pptos_quota.dat" Then PPTSec_QuotaFile = "pptos_quota.dat"
End Function

Private Function PPTSec_AuditFile() As String
    PPTSec_AuditFile = Environ("TEMP") & "\pptos_audit.log"
    If PPTSec_AuditFile = "\pptos_audit.log" Then PPTSec_AuditFile = "pptos_audit.log"
End Function

' 读取当天配额计数：key = uid|yyyymmdd|action
Private Function PPTSec_QuotaLoad() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    Dim ff As Integer, line As String, p() As String, stamp As String
    stamp = Format$(Date, "yyyymmdd")
    ff = FreeFile
    Open PPTSec_QuotaFile() For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, line
        p = Split(line, "|")
        If UBound(p) = 3 Then
            If p(1) = stamp Then d(p(0) & "|" & p(1) & "|" & p(2)) = CLng(Val(p(3)))
        End If
    Loop
    Close #ff
Done:
    Set PPTSec_QuotaLoad = d
End Function

Private Sub PPTSec_QuotaSave(ByRef d As Object)
    On Error Resume Next
    Dim ff As Integer, k As Variant
    ff = FreeFile
    Open PPTSec_QuotaFile() For Output As #ff
    For Each k In d.Keys
        Print #ff, k & "|" & d(k)
    Next k
    Close #ff
End Sub

Private Function PPTSec_QuotaGet(ByRef d As Object, ByVal uid As String, ByVal action As String) As Long
    Dim key As String
    key = uid & "|" & Format$(Date, "yyyymmdd") & "|" & action
    If d.Exists(key) Then PPTSec_QuotaGet = CLng(d(key)) Else PPTSec_QuotaGet = 0
End Function

Private Sub PPTSec_QuotaSet(ByRef d As Object, ByVal uid As String, ByVal action As String, ByVal n As Long)
    Dim key As String
    key = uid & "|" & Format$(Date, "yyyymmdd") & "|" & action
    d(key) = n
End Sub

Private Function PPTSec_Limit(ByVal action As String) As Long
    Select Case LCase$(action)
        Case "ai": PPTSec_Limit = PPTSEC_LIMIT_AI
        Case "calc": PPTSec_Limit = PPTSEC_LIMIT_CALC
        Case Else: PPTSec_Limit = PPTSEC_LIMIT_OTHER
    End Select
End Function

' 用户 ID 合法性（纯数字，长度 1~20）
Private Function PPTSec_UidValid(ByVal uid As String) As Boolean
    Dim i As Long, ch As String
    If Len(uid) = 0 Or Len(uid) > 20 Then Exit Function
    For i = 1 To Len(uid)
        ch = Mid$(uid, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    PPTSec_UidValid = True
End Function

' 审计日志（超过 200KB 直接重置，避免无限增长）
Public Sub PPTSec_Audit(ByVal action As String, ByVal detail As String)
    On Error Resume Next
    Dim ff As Integer
    If FileLen(PPTSec_AuditFile()) > 204800 Then Kill PPTSec_AuditFile()
    ff = FreeFile
    Open PPTSec_AuditFile() For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "|" & PPTSec_Uid() & "|" & action & "|" & detail
    Close #ff
End Sub

' 联网调用前的闸门：登录态 → 配额（通过即预占一次）
Public Function PPTSec_CanCall(ByVal action As String) As Boolean
    On Error GoTo Deny
    Dim uid As String
    uid = PPTSec_Uid()
    If Not PPTSec_UidValid(uid) Then
        MsgBox "请先完成 OK 账户登录（点“登录你的 OK 账户”）。", vbExclamation
        Exit Function
    End If

    Dim d As Object, used As Long, lim As Long
    Set d = PPTSec_QuotaLoad()
    used = PPTSec_QuotaGet(d, uid, action)
    lim = PPTSec_Limit(action)

    If used >= lim Then
        MsgBox "今日“" & action & "”调用额度已用完（" & used & "/" & lim & "）。" & vbCrLf & _
               "额度每天 0 点重置；如需更多请联系 Caelus Studio。", vbInformation
        Exit Function
    End If

    Call PPTSec_QuotaSet(d, uid, action, used + 1)
    Call PPTSec_QuotaSave(d)
    PPTSec_CanCall = True
    Exit Function
Deny:
    MsgBox "调用被安全策略拒绝：" & Err.Description, vbExclamation
End Function

' 调用结算：失败退还额度，并记审计
Public Sub PPTSec_Commit(ByVal action As String, ByVal ok As Boolean, ByVal detail As String)
    On Error Resume Next
    Dim uid As String, d As Object, used As Long
    If Not ok Then
        uid = PPTSec_Uid()
        Set d = PPTSec_QuotaLoad()
        used = PPTSec_QuotaGet(d, uid, action)
        If used > 0 Then
            Call PPTSec_QuotaSet(d, uid, action, used - 1)
            Call PPTSec_QuotaSave(d)
        End If
    End If
    Call PPTSec_Audit(action, IIf(ok, "ok", "fail") & "|" & detail)
End Sub

' 查询今日剩余额度
Public Function PPTSec_QuotaLeft(ByVal action As String) As Long
    On Error Resume Next
    Dim d As Object
    Set d = PPTSec_QuotaLoad()
    PPTSec_QuotaLeft = PPTSec_Limit(action) - PPTSec_QuotaGet(d, PPTSec_Uid(), action)
End Function

' 可绑定到按钮：查看今日额度
Public Sub PPTSec_ShowQuota()
    MsgBox "今日额度" & vbCrLf & _
           "Star Intelligence：" & PPTSec_QuotaLeft("ai") & " / " & PPTSEC_LIMIT_AI & " 次" & vbCrLf & _
           "计算器：" & PPTSec_QuotaLeft("calc") & " / " & PPTSEC_LIMIT_CALC & " 次" & vbCrLf & vbCrLf & _
           "用户 ID：" & IIf(PPTSec_Uid() = "", "（未登录）", PPTSec_Uid()), vbInformation
End Sub

' ---- 时间与剪贴板工具 ----
Private Function PPTOS_NowSec() As Double
    PPTOS_NowSec = Timer()
End Function

' 已过去秒数（处理跨零点回绕）
Private Function PPTOS_Elapsed(ByVal startSec As Double) As Double
    Dim n As Double
    n = Timer()
    If n < startSec Then n = n + 86400
    PPTOS_Elapsed = n - startSec
End Function

Private Sub PPTOS_SleepMs(ByVal ms As Long)
    Dim t As Double
    t = PPTOS_NowSec()
    Do
        DoEvents
    Loop While PPTOS_Elapsed(t) * 1000 < ms
End Sub

' 读剪贴板文本（失败返回空串，调用方需容错）
Private Function PPTOS_ClipboardText() As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If d Is Nothing Then Exit Function
    d.GetFromClipboard
    PPTOS_ClipboardText = CStr(d.GetText)
    On Error GoTo 0
End Function

' 从任意文本里提取 6 位字母数字验证码（页面文案如“验证码：ABC123”也能识别）
Private Function PPTOS_ExtractCode(ByVal s As String) As String
    Dim i As Long, ch As String, run As String, best As String
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then
            run = run & ch
        Else
            If Len(run) = 6 Then best = run
            run = ""
        End If
    Next i
    If Len(run) = 6 Then best = run
    PPTOS_ExtractCode = best
End Function
