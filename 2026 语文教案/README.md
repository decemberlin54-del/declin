# 2026 语文教案

个人语文教研与研究工作区。云端仓库：[decemberlin54-del/declin](https://github.com/decemberlin54-del/declin)

## 本机路径（Windows）

```
C:\Users\Bechodan1209\Documents\2026 语文教案\
```

## 子项目

| 文件夹 | 说明 |
|--------|------|
| `初中语文母题研究/` | 生命母题课程群研究（读博方向） |
| `学历案/` | **暑假任务①**：单元学历案（七上第六单元初稿已完成） |
| `暑假任务/` | 任务清单、试题命制 |
| `模板/` | 学历案通用模板与参考样例 |

## 换电脑同步（三步）

### 第一次在这台电脑设置

1. 安装 [Git](https://git-scm.com/download/win) 和 [Cursor](https://cursor.com)
2. 在 PowerShell 中运行：

```powershell
cd "C:\Users\Bechodan1209\Documents"
git clone https://github.com/decemberlin54-del/declin.git "2026 语文教案"
```

3. Cursor → **File → Open Folder** → 选择 `C:\Users\Bechodan1209\Documents\2026 语文教案`

也可以直接双击运行：`scripts\首次安装到本机.ps1`

### 每天写完保存到云端

```powershell
cd "C:\Users\Bechodan1209\Documents\2026 语文教案"
git add .
git commit -m "更新教案"
git push
```

或双击：`scripts\保存并上传云端.ps1`

### 换新电脑后恢复

```powershell
cd "C:\Users\Bechodan1209\Documents"
git clone https://github.com/decemberlin54-del/declin.git "2026 语文教案"
```

然后用 Cursor 打开同一文件夹即可，所有文件与旧电脑一致。

## 注意事项

- **云端 = GitHub 仓库**，不依赖 U 盘或微信传文件
- 修改后记得 `git push`，换电脑前务必执行一次上传
- 仓库为私有项目，需用 GitHub 账号登录授权
