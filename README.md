# Java POI 操作渲染PPT

[![HOME](https://img.shields.io/badge/HOME-PlumLi-brightgreen.svg.svg)](https://github.com/licunzhi/dream_on_sakura_rain)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/licunzhi/dream_on_sakura_rain/pulls)
[![GitHub stars](https://img.shields.io/github/stars/licunzhi/poi.svg?style=social)](https://github.com/licunzhi/poi/blob/master/README.md)
[![GitHub forks](https://img.shields.io/github/forks/licunzhi/poi.svg?style=social)](https://github.com/licunzhi/poi/blob/master/README.md)

# 运行环境
- java8 or later
- Eclipse IDE for Java EE Developers
- IntelliJ IDEA
- Text file encoding: UTF-8

# 运行方法
- clone or [download]() project到本地
- 开发工具运行主函数[Main](./src/main/java/com/sakura/rain/Main.java)
- Main读取的ppt模板要根据项目需求自定义存放和读取位置
- PptUtis不直接存储文件，你可以用XMLSlideShow接收方法返回值，使用方法write(outputStream)写入文件流中
- web项目中应用只需拷贝[PptModel](./src/main/java/com/sakura/rain/model/PptModel.java) [PptUtils](./src/main/java/com/sakura/rain/utils/PptUtils.java)

# 依赖信息
```xml
<dependencies>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>3.8</version>
    </dependency>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml-schemas</artifactId>
        <version>3.8</version>
    </dependency>
</dependencies>
```

# 备注
- [ ] 创建表格
- [X] 创建文本
- [X] 图片路径方式
- [X] 图片文件方式
- [X] 图片Base64方式
