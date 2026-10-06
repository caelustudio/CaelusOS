Attribute VB_Name = "模块3"
Sub Demo1_Click()
    '1. 执行已有数据加载宏
    Call AppLoadData
    '2. 跳转至第70页幻灯片
    Call GotoNamedSlide("NotesPage")
End Sub
     