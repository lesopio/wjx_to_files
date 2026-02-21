# wjx_to_files 问卷星问卷导出工具
把问卷星的在线问卷转换为docx文档或者markdown以及其他格式
将公开问卷星链接解析为结构化内容，并同时导出三种文件：

- `.docx`（表格化题目清单）
- `.json`（适合程序和大模型读取）
- `.md`（适合人工阅读和大模型上下文输入）

## 环境要求

- Python 3.9+

安装依赖：

```bash
pip install requests beautifulsoup4 python-docx
```

## 使用方式

```bash
python wjx_to_docx.py <问卷链接>
```

示例：

```bash
python wjx_to_docx.py https://v.wjx.cn/vm/aaaaa.aspx
```

运行成功后，会在当前目录生成同名的：

- `xxx.docx`
- `xxx.json`
- `xxx.md`

若同名文件已存在，会自动追加时间戳避免覆盖。

## 输出内容

`docx` 表格列固定为：

- 题号
- 题型
- 必填
- 题干
- 选项
- 逻辑

`json` 主要字段包括：

- `title`
- `description`
- `source_url`
- `crawl_time`
- `sections`
- `questions`

`md` 包含元数据、章节、逐题结构（题型/题干/选项/逻辑）。

## 限制说明

- 仅支持公开、可直接访问的问卷页面
- 不支持验证码、密码、登录态问卷
- 页面结构发生较大变化时可能解析失败

## 退出码

- `0` 成功
- `1` 参数或链接不合法
- `2` 网络请求失败
- `3` 页面解析失败或受限
- `4` 导出文件写入失败
