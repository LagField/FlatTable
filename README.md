# FlatTable
简单、快速、超轻量级、面向策划友好的导表工具。

## 快速开始

![title](/img/interface.png?raw=true)

[Publish](/Publish/)文件夹中，其中有预先编译好的exe执行文件，下载打开exe即可看到上图。
工具需要你配置excel文件所在路径和生成的C#文件，表格内容编码的二进制文件路径。
配置过后，左边选择想要导出的表格，点击开始生成即可。如果表格格式有错误，工具会给出相应提示。

除了自动生成的cs文件，只另外提供两个简单的载入文件(TableLoader)和表格的基类文件(TableBase)。

如果是Unity项目
[TableBase](/UnityTest/UnityTest/Assets/Scripts/TableBase.cs)   [TableLoader](/UnityTest/UnityTest/Assets/Scripts/TableLoader.cs)   

如果是.net项目
[TableBase](/FlatTable/FlatTable/RuntimeCode/TableBase.cs)   [TableLoader](/FlatTable/FlatTable/RuntimeCode/TableLoader.cs)   

## Excel表格格式要求
