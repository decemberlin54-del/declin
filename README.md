# declin · 初中语文教学资源库

GitHub 账号 [decemberlin54-del](https://github.com/decemberlin54-del) 下目前只有本仓库。原先内容分散在多个 `cursor/*` 分支，现已按主题收拢到 **`projects/`** 目录，便于查找与下载。

## 目录结构

| 目录 | 说明 | 主要内容 |
|------|------|----------|
| [`projects/gushi-lessons/`](projects/gushi-lessons/) | 古诗三课课件 | 《次北固山下》《闻王昌龄…》《天净沙·秋思》15 页 PPT + 生成脚本 |
| [`projects/unit6-grade7/`](projects/unit6-grade7/) | 七年级第六单元 | 教案、研究汇报 PPT、期末试卷与细目表、批量生成脚本 |
| [`projects/motif-research/`](projects/motif-research/) | 生命母题研究 | 工作库说明、Cursor MCP 配置备份 |
| [`projects/teacher-cert-config/`](projects/teacher-cert-config/) | 教资辅导 | `.cursorrules` 备份（根目录有一份生效副本） |

## 常用下载（稳定链接格式）

在 GitHub 上打开对应文件后点 **Download**，或使用 Raw（路径需含 `refs/heads/main`）：

```text
https://raw.githubusercontent.com/decemberlin54-del/declin/refs/heads/main/projects/gushi-lessons/gushi-lessons.pptx
```

## 重新生成课件

**古诗三课：**

```bash
pip install -r projects/gushi-lessons/requirements-ppt.txt
python3 projects/gushi-lessons/scripts/generate_gushi_ppt.py
```

**第六单元：** 见 [`projects/unit6-grade7/README.md`](projects/unit6-grade7/README.md) 与各 `scripts/generate_*.py`。

## 历史分支与 PR

整理前的工作在以下分支（合并后可择机删除分支）：

| 分支 | 草稿 PR |
|------|---------|
| `cursor/gushi-lesson-ppt-0f9a` | [#4](https://github.com/decemberlin54-del/declin/pull/4) |
| `cursor/unit6-teaching-plan-e91b` | [#3](https://github.com/decemberlin54-del/declin/pull/3) |
| `cursor/junior-chinese-motif-research-49ed` | [#2](https://github.com/decemberlin54-del/declin/pull/2) |
| `cursor/teacher-exam-cursorrules-737c` | [#1](https://github.com/decemberlin54-del/declin/pull/1) |

## 仓库维护建议

1. **新资料**：在 `projects/<主题>/` 下新增，避免再在仓库根目录堆文件。
2. **合并本整理 PR 后**：关闭或合并上述 4 个草稿 PR，删除已合并的 `cursor/*` 分支。
3. **无数据库**：本仓库仅为文档与脚本，不含业务数据库。
