Attribute VB_Name = "模块4"
Option Explicit

' ====== OK 云计算器 PPT 模块 ======
' 应用：计算器
' 生成时间：2026/8/3 12:53:22

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

' 提取 JSON 字符串值
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
            OKApp_ExtractJsonStr = raw
            Exit Function
        End If
    End If
    OKApp_ExtractJsonStr = ""
End Function

' ====== 云计算器核心函数 ======

' 通用计算函数：发送表达式到服务器，返回结果
' 失败时返回空字符串
Public Function AppCalcEval(expr As String) As String
    On Error GoTo ErrHandler
    AppCalcEval = ""
    If expr = "" Then Exit Function
    If Not OKApp_EnsureLogin() Then Exit Function
    ' 安全闸门：校验登录态 + 每日调用配额（通过即预占一次，失败会退还）
    If Not PPTSec_CanCall("calc") Then Exit Function
    Dim uid As String
    uid = OKApp_GetUserId()
    Dim xmlhttp As Object, json As String
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", "https://www.okteam.cn/app/api/calc", False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "X-PPTOS-Client", PPTSec_ClientTag()
    Dim body As String
    body = "{""app_key"":""" & PPTSec_Key("calc") & """,""user_id"":""" & uid & """,""expr"":""" & OKApp_JsonEscape(expr) & """}"
    Call WaitOn
    xmlhttp.Send body
    Call WaitOff
    If xmlhttp.Status <> 200 Then
        Call PPTSec_Commit("calc", False, "http" & xmlhttp.Status)
        MsgBox "计算失败，HTTP " & xmlhttp.Status, vbExclamation
        Exit Function
    End If
    json = xmlhttp.responseText
    If InStr(json, """ok"":true") > 0 Then
        AppCalcEval = OKApp_ExtractJsonStr(json, "result")
        Call PPTSec_Commit("calc", True, "ok")
    Else
        Call PPTSec_Commit("calc", False, Left$(OKApp_ExtractJsonStr(json, "error"), 80))
        MsgBox "计算失败：" & OKApp_ExtractJsonStr(json, "error"), vbExclamation
    End If
    Exit Function
ErrHandler:
    Call WaitOff
    Call PPTSec_Commit("calc", False, "exception")
    MsgBox "计算出错：" & Err.Description, vbCritical
End Function

' ====== 计算器核心 UI 函数 ======

' 向显示框追加字符（内部函数）
Private Sub AppCalcInput(ByVal ch As String)
    On Error GoTo ErrHandler
    Dim dispShp As Object
    Set dispShp = OKApp_FindShape("Calc_TextBox")
    If dispShp Is Nothing Then Exit Sub
    Dim currentText As String
    currentText = OKApp_GetShapeText(dispShp)
    OKApp_SetShapeText dispShp, currentText & ch
    Exit Sub
ErrHandler:
    Call WaitOff
End Sub

' 等号：计算表达式并显示结果
Public Sub AppCalcEquals()
    On Error GoTo ErrHandler
    If Not OKApp_EnsureLogin() Then Exit Sub
    Dim dispShp As Object
    Set dispShp = OKApp_FindShape("Calc_TextBox")
    If dispShp Is Nothing Then Exit Sub
    Dim expr As String
    expr = OKApp_GetShapeText(dispShp)
    Dim result As String
    result = AppCalcEval(expr)
    If result <> "" Then
        OKApp_SetShapeText dispShp, result
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
End Sub

' 清空显示框
Public Sub AppCalcClear()
    On Error GoTo ErrHandler
    Dim dispShp As Object
    Set dispShp = OKApp_FindShape("Calc_TextBox")
    If dispShp Is Nothing Then Exit Sub
    OKApp_SetShapeText dispShp, ""
    Exit Sub
ErrHandler:
    Call WaitOff
End Sub

' 退格：删除最后一个字符
Public Sub AppCalcBackspace()
    On Error GoTo ErrHandler
    Dim dispShp As Object
    Set dispShp = OKApp_FindShape("Calc_TextBox")
    If dispShp Is Nothing Then Exit Sub
    Dim currentText As String
    currentText = OKApp_GetShapeText(dispShp)
    If Len(currentText) > 0 Then
        OKApp_SetShapeText dispShp, Left$(currentText, Len(currentText) - 1)
    End If
    Exit Sub
ErrHandler:
    Call WaitOff
End Sub

' ====== 按键包装宏（每个按键一个宏，绑定到形状动作） ======

Public Sub CalcBtn_0()
    Call AppCalcInput("0")
End Sub

Public Sub CalcBtn_1()
    Call AppCalcInput("1")
End Sub

Public Sub CalcBtn_2()
    Call AppCalcInput("2")
End Sub

Public Sub CalcBtn_3()
    Call AppCalcInput("3")
End Sub

Public Sub CalcBtn_4()
    Call AppCalcInput("4")
End Sub

Public Sub CalcBtn_5()
    Call AppCalcInput("5")
End Sub

Public Sub CalcBtn_6()
    Call AppCalcInput("6")
End Sub

Public Sub CalcBtn_7()
    Call AppCalcInput("7")
End Sub

Public Sub CalcBtn_8()
    Call AppCalcInput("8")
End Sub

Public Sub CalcBtn_9()
    Call AppCalcInput("9")
End Sub

Public Sub CalcBtn_Add()
    Call AppCalcInput("+")
End Sub

Public Sub CalcBtn_Sub()
    Call AppCalcInput("-")
End Sub

Public Sub CalcBtn_Mul()
    Call AppCalcInput("*")
End Sub

Public Sub CalcBtn_Div()
    Call AppCalcInput("/")
End Sub

Public Sub CalcBtn_Mod()
    Call AppCalcInput("%")
End Sub

Public Sub CalcBtn_Pow()
    Call AppCalcInput("^")
End Sub

Public Sub CalcBtn_Dot()
    Call AppCalcInput(".")
End Sub

Public Sub CalcBtn_LP()
    Call AppCalcInput("(")
End Sub

Public Sub CalcBtn_RP()
    Call AppCalcInput(")")
End Sub

Public Sub CalcBtn_Pi()
    Call AppCalcInput("pi")
End Sub

Public Sub CalcBtn_E()
    Call AppCalcInput("e")
End Sub

Public Sub CalcBtn_Sin()
    Call AppCalcInput("sin(")
End Sub

Public Sub CalcBtn_Cos()
    Call AppCalcInput("cos(")
End Sub

Public Sub CalcBtn_Tan()
    Call AppCalcInput("tan(")
End Sub

Public Sub CalcBtn_SinH()
    Call AppCalcInput("sinh(")
End Sub

Public Sub CalcBtn_CosH()
    Call AppCalcInput("cosh(")
End Sub

Public Sub CalcBtn_TanH()
    Call AppCalcInput("tanh(")
End Sub

Public Sub CalcBtn_ASin()
    Call AppCalcInput("asin(")
End Sub

Public Sub CalcBtn_ACos()
    Call AppCalcInput("acos(")
End Sub

Public Sub CalcBtn_ATan()
    Call AppCalcInput("atan(")
End Sub

Public Sub CalcBtn_Ln()
    Call AppCalcInput("ln(")
End Sub

Public Sub CalcBtn_Log()
    Call AppCalcInput("log(")
End Sub

Public Sub CalcBtn_Log2()
    Call AppCalcInput("log2(")
End Sub

Public Sub CalcBtn_Log10()
    Call AppCalcInput("log10(")
End Sub

Public Sub CalcBtn_Sqrt()
    Call AppCalcInput("sqrt(")
End Sub

Public Sub CalcBtn_Abs()
    Call AppCalcInput("abs(")
End Sub

Public Sub CalcBtn_Exp()
    Call AppCalcInput("exp(")
End Sub

Public Sub CalcBtn_Floor()
    Call AppCalcInput("floor(")
End Sub

Public Sub CalcBtn_Ceil()
    Call AppCalcInput("ceil(")
End Sub

Public Sub CalcBtn_Round()
    Call AppCalcInput("round(")
End Sub

Public Sub CalcBtn_Max()
    Call AppCalcInput("max(")
End Sub

Public Sub CalcBtn_Min()
    Call AppCalcInput("min(")
End Sub

' 等号 / 清空 / 退格 直接使用 AppCalcEquals / AppCalcClear / AppCalcBackspace

