# ExpenseReportAssistant
一个能帮助你自动将发票+订单截图+支付记录整理在A4纸上的小程序。
## 效果展示
![image](https://github.com/user-attachments/assets/5790c226-b308-4c25-8708-a29887b4e742)
![image](https://github.com/user-attachments/assets/8313107d-f964-444b-84d8-7332a20a291f)

## 使用教程
### 运行环境
本程序需要 Python 3.10 以上环境（推荐3.12+），并需要安装`PyMuPDF`、`tqdm`、`pillow`。
```
pip install pymupdf tqdm pillow
```

### 运行程序
在程序所在目录打开终端即可：
```
python assistant.py
# 这将以本文件夹作为输入和输出路径
```
你也可以手动指定输入和输出文件夹：
```
python assistant.py --input-folder ./example --output example.pdf
python assistant.py --input-folder D:/path/to/your/folder --output D:/path/to/your
```
程序可以接受以下几个参数输入：
1. `--input-folder`或`-I`（不指定则使用当前文件夹）
   指定输入的**文件夹**。
   此文件夹需要遵守**文件整理**部分的规则，包含若干个子文件夹。
3. `--output`或`-O`（不指定则使用当前文件夹与自动生成的文件名）
   指定输出的**文件**或**文件夹**。
   如果输入一个文件名（或以文件名结尾的路径），那么程序会尝试输出到此文件。
   如果输入一个文件夹，将使用自动生成的文件名，文件保存到此文件夹内。
4. `--debug`或`-D`
   此参数会在执行过程中额外打印一些调试信息，一般不需要使用。

### 文件整理
首先，你需要将文件整理成这样的格式：
```
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
```
每个子文件夹中都需要至少包含**1个PDF**与**2张图片**，推荐将图片数量控制在4张以内，不建议超过8张。
若超出8张图片建议手动拼图，并在文件名的前缀加上`NEWLINE`（独占一行）或`NEWPAGE`（独占一页）
内，运行脚本即可。

### 输出结果不符合预期?
你可以手动更改程序逻辑，或者，你可以更改文件名，在文件名的前缀加上`NEWLINE`或`NEWPAGE`，可以让该图片**单独占一行**或**单独占一页**，通过手动拼图并将文件改为这样的名称，在大多数情况下能够实现你想要的效果。
