html转word
<h1>html转换成word文档</h1>

## 项目简介

最近一直在做关于前端富文本生成的html代码转换成word的需求，对各种工具进行了尝试，对效果都不是很满意，最后从茫茫大海中找到一个非常好的工具，顾分享出来

## 主要转换方式
>经过了一周的艰苦测试，主要通过以下三种方式尝试导出，发现各有利弊,当然除了这三种方式还有很多网络大神尝试自己手动去组装word文件流，并且修改样式等操作，但是代码实在是太长了，并且也极具个性化，不适合不同的人使用，也不是很推荐

### poi方式：
>在网上调研发现，主流的一种方式就是用过poi这个比较流行的第三方工具包进行转换
但是在使用过程中发现了许多痛点，虽然可以满足转换的需求，但是转换撑的word文档存在以下问题
  - 如果文档中有图片，发现图片是以url的形式存在，如果电脑断网的话图片加载不出来
  - 这种导出方式导出的word本质上其实还是html代码，是靠office或者wps自己解析样式的，所以因为软件的不同导致样式也会发生变化，体验不好
>poi优点
 - 样式兼容相对比较高
 - 导出速度比较快（因为里面的图片不会转换成真正的图片流，只是一个url）
 - 可以直接修改html的样式，就修改了word的样式（例如给图片加个 hight属性 导出来的word里面的图片高度也会变化）

### docx4j方式:
>这种方式底层应该就是对html进行解码转换成实际word底层格式的xml,但是也是有美中不足的地方
 - 资源消耗高，因为会对大量的字符串进行解析转换，图片也需要下载到本地然后写入文件流，所以在导一些大的文件会发生oom（慎用）
 - 样式兼容不好，有一些复杂的html标签无法解析
 - 很难对word文档中的一些样式进行自定义
 - api文档少的可怜，资料比较少
 - 图片虽然会变成文件流的方式储存到文档中，但是样式就是图片的原本大小，经常出现图片太大了，页面只能展示一部分，需要自己手动调整样式
>docx4j优点
 - 会替换文件中的图片，没网也可以展示

###e-iceblue 方式:(爆赞)
>不知道为什么这么好的方式，没什么人用，网上几乎搜不到
> 这种方式类似docx4j方式一样，都是解析html标签，生成指定格式的文档 推荐原因：
 - 但是相比于docx4j这种方式资源消耗小，速度更快
 - 可以自己定义word的各种样式（例如页面大小，表格样式等）
 - api巨全，不知道这个是不是国人开发的，但是有中文网站，对应的api接受也是非常全面
 - 对图片样式html的修改会直接映射到word文档中，自动适配图片大小
 - 出了做html导出word方式外，几乎所有的文档格式转换都支持

>美中不足吧
 - 耗时较长

官网地址：
`https://www.e-iceblue.cn/`
iceblue除了html转换word之外，几乎所有的office产品之间个的格式转换都是支持的。有兴趣的读者可以去官网自行研究
源码地址
`https://github.com/eiceblue/Spire.PDF-for-Java`

##单元测试
>读者可以试着运行一下

`TestExport`

>的main方法,只需要修改自己的问文件输出路径即可，看看那种适合你把
