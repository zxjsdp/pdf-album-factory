# pdf-album-factory

根据图片及内容，批量生成漂亮的 PDF 格式相册文件。

# 使用方法

## 第一步：安装 Python

若您使用的是 Windows，则可能并没有安装过 Python。请前往 <https://www.python.org/downloads/windows/> 下载并安装，建议使用 3.6 以上的版本。

若您使用的是 Linux / macOS，则应该已有预置的 Python 环境。

## 第二步：安装 wkhtmltopdf

Wkhtmltopdf 是可用于将 HTML 转换为 PDF 的开源工具，GitHub 链接为：[GitHub - wkhtmltopdf/wkhtmltopdf](https://github.com/wkhtmltopdf/wkhtmltopdf/)。

可前往 <https://wkhtmltopdf.org/downloads.html> 下载您所在平台对应的安装包，并进行安装。

## 第三步：安装所依赖的 Python 包

1. 在命令行窗口中，切换路径至 `pdf-album-factory` 目录下。
2. 执行 pip 命令，安装 requirements.txt 中的所有 Python 包（建议使用 virtualenv）

        pip install -r requirements.txt

其中需要特别关注的两个包分别为：

1. openpyxl

    读取及写入 Excel 2010 xlsx/xlsm/xltx/xltm 文件的包。文档可参考：[OpenPyXL](https://openpyxl.readthedocs.io/en/stable/);

2. pdfkit

    wkhtmltopdf 的 Python 封装包，用户将 HTML 转换为 PDF。文档可参考：[GitHub - JazzCore/python-pdfkit](https://github.com/JazzCore/python-pdfkit)。
    

## 第四步：修改模板文件及准备图片

默认的模板文件名为 `相册数据模板.xlsx`。请按照需求修改模板文件，并将需要加入到相册中的图片信息准确的填写在模板文件中。

## 第五步：

执行生成逻辑（建议在 virtualenv 中执行）：

    python generate_pdf_album.pdf

注意：

在使用此工具的流程中，请不要随意修改任何文件的相对路径，避免找不到对应文件的情况。
