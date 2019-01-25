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

每个表格数据都会对应一个静态类，Unity项目建议游戏启动时就把所有配置表加载了

```csharp
        //例子：从StreamingAssets路径加载
        //设置加载模式
        TableLoader.loadType = LoadType.FilePath;
        //设置加载路径
        TableLoader.fileLoadPath = Application.streamingAssetsPath + "/Table";
        //加载
        TableLoader.Load<AnotherTestTable>();
        
        //接下来就可以读取数据了
        if (AnotherTestTable.ins == null)
        {
            return;
        }
        List<AnotherTestTable.Value> valueList = AnotherTestTable.ins.list;
        for (int i = 0; i < valueList.Count; i++)
        {
            Debug.Log(valueList[i].id);
            Debug.Log(valueList[i].hero_name);
            Debug.Log(valueList[i].speed);
            Debug.Log(valueList[i].damage);
            Debug.Log(valueList[i].is_lock);
        }
        
        //-----------------------------------------
        
        //同样提供自定义加载方法，只需要设置加载回调函数，程序在Load时则会调用该函数进行加载
        TableLoader.customLoader += MyCustomLoader;
        TableLoader.Load<AnotherTestTable>();
        
        //加载函数会传入当前类型对应的文件名称，自定义函数里根据名称加载，并返回一个byte[]。因为Decode只会从头读取数据，读到需要的长度后就会停止，可以方便byte[]的复用。
```


## Excel表格格式要求
表格只会读取第一个sheet。

表格第一行为字段名，也就是导出后代码中的变量名称。第二行表示该列是否需要导出，如果不需要导出，则填入exclude即可(第一列不能exclude)。第三行表示该字段类型，目前支持类型有int,short,float,bool,string及他们的数组类型。

表格有效区域由第一行和第一列来决定，工具会扫描第一行直到遇到第一个空格，并以此作为一行的长度(也就是说字段名是必填)。同样，表格会扫描第一列，直到碰到第一个空格(除了第二行可以不填),并以此作为一列的长度。

表格的第一列必须是"id"，类型必须是int，第二行不能是exclude，而且id该列所有int值不能有重复的。因为生成的代码会自动组合一个字典，它的key就是存储的该id，可用于快速查表。

如果要使用数组，则在字段名后面加上[]并在里面写上该元素是数组的第几个，从0开始，必须连续。例如map_id[0] map_id[1] map_id[2]。数组类型在表格中可以以任意顺序放置，程序会自动帮你把数组排序存放，只需要保证序号是从0开始且连续的，否则工具会弹出错误提示。

如果读取范围内有格子没有内容，则会记录该类型的默认值。

excel格式例子可以参考[文件夹](/Test/ExcelFile/)
