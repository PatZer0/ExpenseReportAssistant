# ExpenseReportAssistant
一个能帮助你自动将发票+订单截图+支付记录整理在A4纸上的小程序。

## 使用教程
### 运行环境
本程序需要 Python 3.10 以上环境（推荐3.12+），并需要安装`PyMuPDF`、`tqdm`、`pillow`。
```
pip install pymupdf tqdm pillow
```

### 文件整理
首先，你需要将文件整理成这样的格式：

.

├── 子文件夹1

│   ├── 图片1.jpg

│   ├── 图片2.jpg

│   ├── 图片3.jpg

│   ├── 图片4.jpg

│   └── 发票.pdf

├── 子文件夹2

│   ├── 发票.pdf

│   ├── 图片1.jpg

│   └── 图片2.jpg

……
每个子文件夹中都需要至少包含**1个PDF**与**2张图片**，图片数量理论上无上限。

### 运行脚本
默认情况下，只需要将这些包含有PDF与图片的子文件夹放到与`assistant.py`同级别的文件夹内，运行脚本即可。
你也可以修改`main`函数中的`base_folder`，以让脚本在本目录以外的路径中运行。
