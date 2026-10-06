# CaelusOS（PPTOS）VBA 源码与构建工具

CaelusOS 的界面是 94 张幻灯片，真正干活的是文件里的 VBA。VBA 在二进制里没法做版本管理，
所以这里把宏抽成文本源码维护，并提供**重建 pptm** 的工具链。

```
vba/
├── src/                  VBA 源码（UTF-8；构建时自动转成工程代码页 936/GBK）
│   ├── 模块1.bas         登录 + 幻灯片跳转 + 等待光标 + 安全配额（PPTSec）
│   ├── 模块2.bas         笔记：云数据读写、公共数据、检查更新
│   ├── 模块3.bas         示例宏 Demo1_Click
│   ├── 模块4.bas         计算器（按键宏 + 云计算）
│   ├── 模块5.bas         Star Intelligence 对话
│   └── Slide49/91/94/100.cls   幻灯片事件模块
├── tools/pptos_vba.py    抽取 / 重建 / 校验 / 体检 / 生成混淆 key
└── APPS.md               还能再做什么 app 的候选清单
```

## 工具用法

```bash
PY=/Users/orange/.workbuddy/binaries/python/envs/default/bin/python

# 从 pptm 抽出源码
$PY tools/pptos_vba.py extract "CaelusOS.pptm" src

# 用源码重建一份新的 pptm（不动原文件）
$PY tools/pptos_vba.py build src "CaelusOS.pptm" "CaelusOS-out.pptm"

# 校验：把新 pptm 里的宏解回来跟源码逐行比对
$PY tools/pptos_vba.py verify "CaelusOS-out.pptm" src

# 粗检源码（引号成对、块平衡）
$PY tools/pptos_vba.py lint src

# 重新生成 app_key 的混淆密文（改了 key 时用）
$PY tools/pptos_vba.py keys
```

## 实现要点（踩过的坑都在这里）

1. **模块流 = 源码容器 + 已编译 p-code**。重建时只写源码、丢掉 p-code，宿主必须重新编译，
   改动才会真正生效；同时把 `__SRP_*` 性能缓存清空（size=0），避免宿主读到旧缓存。
2. **CFB 迷你流阈值**：小于 4096 字节的流必须放在迷你流里。原本是普通流的模块**不能**被压到
   4096 以下（读取方会按迷你流解析 → 乱码），工具会自动改用压缩率更差的字面量编码把流撑住。
3. **压缩容器格式**：`0x01` + 块序列；块头 12 位字段 = 数据长度 − 1，块长 = 字段 + 3；
   CopyToken 的 offset/length 位分配随**块内位置**变化（MS-OVBA 2.4.1.3.19.1）。
   工具的压缩输出已用 oletools 的解压器交叉验证过，字节级一致。
4. **工程代码页是 936**：源文件按 UTF-8 存仓库，构建时转 GBK，别用 UTF-8 直接写进流。
5. **只能原地改，不能新增模块**：工具目前不写 `dir` 流，所以新功能请先加进现有模块；
   若确实需要新增模块（如模块6），要先扩展工具的目录流写入能力。
6. **每次只改源码、重新构建**，不要在 PowerPoint 里直接改宏又存回仓库，否则两边会分叉。

## 运行前提

Windows 版 PowerPoint + 启用宏。宏里用到 `shell32`(ShellExecute)、`user32`(SetCursor)、
`Environ("TEMP")`、`MSXML2.XMLHTTP`、`ADODB.Stream`、`Scripting.Dictionary`，都是 Windows 专有；
Mac 版 PowerPoint 不支持 VBA，跑不了。

## 本版改动（2026-10-06）

### 1. 登录改成一键自动（模块1 `LoginOK`）
点「进入浏览器登录」后：打开登录页 → **自动轮询剪贴板**，检测到 6 位字母数字验证码就自动登录、
写入昵称/ID/头像并跳转 LoginDone，全程不用再手打。120 秒没有检测到才回退为手动输入框。
旧流程保留为 `LoginOKManual`。

### 2. app_key 不再明文（模块1 `PPTSec_Key`）
`ak_8a63…`（云数据）、`ak_aaa6…`（AI）、`ak_db33…`（计算器）、登录页 `k=` 参数
全部改成运行时 XOR 还原，源码里搜不到明文。

> 注意：这只是抬高门槛，**不等于安全**。客户端持有的密钥一定可被逆向出来；
> 真正的防护需要平台侧按 `user_id` + `app_key` 做配额与来源校验。

### 3. 调用配额与审计（模块1 PPTSec 区）
- 每次 AI / 计算调用前过 `PPTSec_CanCall`：校验登录态（用户 ID 必须为纯数字）+ 每日配额
- 默认额度：AI 40 次/天、计算器 500 次/天（按 OK 用户 ID 分别计数，可在 `PPTSEC_LIMIT_*` 调整）
- 调用失败自动退还额度；每次调用写本地审计日志
- 请求头新增 `X-PPTOS-Client`，便于平台侧将来做来源识别
- 可绑按钮：`PPTSec_ShowQuota` 查看今日剩余额度
- 配额与审计文件在 `%TEMP%\pptos_quota.dat`、`%TEMP%\pptos_audit.log`
- 模块2（笔记云数据）、模块4（计算器）、模块5（AI）已全部接入

## 联调与回滚

- 新 pptm 建议先在 Windows 上做三件事：开宏 → 点登录（看剪贴板自动登录）→ 调一次 AI 与计算器。
- 万一 PowerPoint 提示需要修复或宏报错，直接退回原始 pptm（本工具从不修改输入文件）。
- 出错信息格式：`错误号 + 描述`；配额/审计文件删掉即恢复初始额度。
