%%style-caption:bold,size=12,font=Hei,center
%%style-image:center

%header{13210110059}
%footer{%right %pagenum  / %numpages}


论文标题
===========

论文副标题
-----------


# 第一级编号标题
## 第二级编号标题

正文内容

### 第三级编号标题

正文内容

jishuqi = %C{jishuqi}

jishuqi = %C{jishuqi}

![imagekey](/Users/ZTH/Projects/Md2Docx/test.jpg 图 杯子)
\\ slash

%concat{我 觉得 是 这样的} = 我觉得是这样的

@{这样呢？} 表示一个footnote

- 一个项目
- 另一个项目

1. 还有编号项
1. 还有编号的！

又是正文

## 又一个二级标题

> 缩进的引用

> 继续缩进的引用

希望你段落之间有空行。


# 交叉引用

怎么引用这里面的题注项目呢？
请参见图 %ref{imagekey}
就引用了前面这张图。

# 表格
用tab缩进和分割的就可以了。
如果tab之后有一个|，我们就根据|来画线。比如说

![表格索引1]( 表)
col1	col2	col3
item	1		2
item	3		4

多个连续的tab被看做是一个tab。这样就有一个最基本的表格.表格应该在一个“段落”中解决。如果我们要确定哪里要画线、哪里不用，用下面的格式。

![表格索引3]( 表 我不知道应该怎么说啊)
name	|	col2		col3
item1	|	1			2
item2	|	3			4

很直观。不过目前只考虑整行都要画线的情况。如果有的地方要划有的地方不要。。。我还没想好怎么办。
如你所见这些都是连贯的。^{脚注}

# 定义替换符号
打中文的时候很开心结果突然要打英文的标点符号一定有点不爽。
我还是没想好怎么办。也许是这样的：

%alias{[ 【}
%alias{! ！}

%style{caption}
诸如此类的……
%style

然后似乎还要定义引用的格式。比方说

> \%crosscite 就是说内部交叉引用的时候显示的是标题和编号。arg的编号和{}中空格分隔的标号一致。
> \%=_name_:_format_ 就是说将命令 \%_name_ 展开为 \%_format_ 来处理。arg的编号从0开始。
> \%=newcommand:\%cite({0})

