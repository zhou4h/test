
import './App.css'
import { convertMarkdownReportToExcel } from "./excel.ts";
import { exportMarkdownToExcel, exportMarkdownToWord} from "./utils.ts";

function App() {

  const markdownContent = `
这是一个**电商平台**开发项目，主要功能包括：
这是一个**电商平台**开发项目，主要功能包括：
这是一个**电商平台**开发项目，主要功能包括：
 
 
# 项目报告

## 项目概述
这是一个**电商平台**开发项目，主要功能包括：


### 核心数据表
| 字段名       | 类型     | 描述                |
|--------------|----------|---------------------|
| user_id      | string   | 用户唯一标识 (UUID)用户唯一标识 (UUID)用户<br> 唯一标识 (UUID)用户唯一标识 (UUID)用户唯一标识 (UUID)用户唯一标识 (UUID)用户唯一标识 (UUID) |
| product_name | varchar  | 商品名称(最长128字符) |

## 代码示例
\`\`\`javascript
function calculatePrice(items) {
  return items.reduce((sum, item) => sum + item.price, 0);
}
\`\`\`

项目预计于2024年Q1上线。
  `

  const markdownContent2 = `
# 项目报告
## sdfsd
## 2023年Q4总结
这是一个**测试**sdjkflsdkjl
这是一个*测试*sdjkflsdkjl


项目整体进展顺利，以下是关键数据：

| 指标         | 目标值 | 实际值 | 完成率 |
|--------------|--------|--------|--------|
| 开发进度     | 100%100%100%100%100<br>%100%100%   | 95%    | 95%    |
| 测试覆盖率   | 80%    | 85%    | 106%   |
| 用户满意度   | 90%    | 92%    | 102%   |

关键成果：
- 提前完成核心模块开发
- 测试覆盖率超过预期
- 获得客户高度评价

下一步计划：
1. Q1完成剩余模块
2. Q2进行用户培训
3. Q3正式上线
  `

  return (
    <>
      <button onClick={() => convertMarkdownReportToExcel(markdownContent)}>
        导出完整文档
      </button>
      <button onClick={() => exportMarkdownToExcel(markdownContent2)}>
        Excel
      </button>
      <button onClick={() => exportMarkdownToWord(markdownContent2)}>
        Word
      </button>
    </>
  )
}

export default App
