# smartppt
首先安装必要的库

pip install python-pptx

编辑config.ini 来选择图片目录

python smartppt.py 

生成ppt

套用自己的标题蒙版

<img width="1329" height="745" alt="image" src="https://github.com/user-attachments/assets/8be88fc9-b798-43c5-9b76-98d38b5cff0e" />
<img width="1329" height="744" alt="image" src="https://github.com/user-attachments/assets/fd41fc63-a878-4f4a-85d7-cafc3af2f3eb" />

以上为两种排版方式

排版逻辑：

二图或三图均居中，占据整体页面的70%，

竖版图片自动3拼，横板图片会结合一张竖版图片做成2拼。

组合逻辑：

多文件夹时，每个文件夹内各自组合，不会混拼。

用途：

在项目初期，把收集到的大量参考图片无脑的排列到ppt中，造成做了很多工作量的假象，用来卷死同时，并瓜分竞争对手的产值。

