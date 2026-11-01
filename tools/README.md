<!-- 文件用途：本目录文件的用途与使用说明；用于协作查阅，不属于参赛报告正文。 -->

# 工具脚本用途说明

脚本顶部已写明输入、输出、依赖与执行影响。当前报告以根目录 Markdown 为准；保留的生成和迁移脚本记录了早期制作过程。

| 脚本 | 用途 | 执行影响与状态 |
| --- | --- | --- |
| [build_tech_note.py](build_tech_note.py) | 用脚本内的固定正文生成旧版 DOCX 和流程图。 | 会覆盖根目录同名 Word；与当前八章 Markdown 不同步，作为历史排版实现参考。 |
| [docx_to_md.py](docx_to_md.py) | 从旧 Word 提取标题、正文、表格和配图，生成最初的 Markdown 主稿。 | 会覆盖根目录同名 Markdown；当前主稿已持续修改，不应使用旧 Word 重新覆盖。 |
| [restructure_eight_chapters.py](restructure_eight_chapters.py) | 对指定六章初稿创建备份后，重排为八章。 | 一次性迁移已完成；包含源稿与备份检查，当前八章稿无需重复运行。 |
| [word_qa.ps1](word_qa.ps1) | 通过本机 Word 统计字数、页数，并导出检查 PDF。 | 从项目根目录执行；只读打开 Word，输出到 .tmp_pdf/docx_render/，不负责 Markdown 转 Word。 |
