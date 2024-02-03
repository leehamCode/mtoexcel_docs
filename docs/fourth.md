## Class Attribute列表

<small>只有类可以使用</small>

#### <font color=licyan>1.FreezeArea表头文本标签</font>

> 规定Excel的冻结区域。

| 属性名         | 类型 | 说明                      |
| -------------- | ---- | ------------------------- |
| FreezeStartRow | int  | 冻结持续的最大行[从0开始] |
| FreezeStartCol | int  | 冻结持续的最大列[从0开始] |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
//冻结范围3行3列
[FreezeArea(FreezeStartCol =2,FreezeStartRow = 2)]
public class TestClass
{
    [HeaderName("姓名")]
    public string thename{get;set;}

    [HeaderName("年龄")]
    public int age{get;set;}

    [HeaderName("地址")]
    [StructTestAttriubte(new String[]{ "西游记","水浒传"})]
    public string address{get;set;}

    [HeaderName("电话")]
    [Horizon(Models.Enums.Horizon.Center,Models.Enums.VerticalHorizon.Up)]
    [BackForeColor(true,new byte[3]{ 50,187,176})]
    [DynaRowColumnLen(123.45,123.45)]
    public string phone{get;set;}
}
```

```c#
public void TestThree()
{
    List<TestClass> listOne = new List<TestClass>(){
        new TestClass(){ thename = "弗里斯兰", age = 800, address = "荷兰低地", phone = "shitU" },
        new TestClass(){ thename = "布列塔尼", age = 1200, address = "布列塔尼", phone = "franc" },
        new TestClass(){ thename = "伊利里亚", age = 2300, address = "亚得里亚", phone = "ita" },
        new TestClass(){ thename = "东色雷斯", age = 2500, address = "黑海", phone = "asdqa" },
        new TestClass(){ thename = "卡帕多西亚", age = 2500, address = "东地中海", phone = "asdqa" },
        new TestClass(){ thename = "", age = 2500, address = "黑海", phone = "asdqa" }
    };

    WrapperConverter wrapper = new WrapperConverter();

    wrapper.basic = new BasicConverter();

    IWorkbook workbook = wrapper.ConvertToExcel<TestClass>(listOne);

    FileStream file = new FileStream("DEMO.xlsx", FileMode.Create);

    workbook.Write(file);

    file.Close();
}
```

![图例](http://img.leepichome.top/open2.png)

#### <font color=licyan>2.Hide_On_Condition条件隐藏标签</font>

> 指定一个表达式,当List元素的值满足表达式条件(为真),该元素代表的行将被设置隐藏。
>
> -------$1为第一个属性的值
>
> -------$2为第二个属性的值
>
> .......

| 属性名       | 类型   | 说明         |
| ------------ | ------ | ------------ |
| rowCondition | string | 行隐藏的条件 |
| colCondition | string | 列隐藏的条件 |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
[TitleAttribute(Context ="测试打印标题行$1",Font_Name ="新細明體",Font_Size =16,Font_color =new byte[]{ 80,235,227 })]
[Hide_On_Condition(rowCondition ="$1==\"九江\"")]
public class TestClass2
{
    [HeaderName("你的姓名:")]
    [MergeNearEqualBox]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Left)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Bottom)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Right)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Upper)]

    public string Name { get; set; }
    
    /* 其他省略 */
}
```

```c#
static void Main(string[] args)
{
    Console.WriteLine("Hello World!");

    WrapperConverter wrap = new WrapperConverter();
    wrap.basic = new BasicConverter();

    List<TestClass2> ts = new List<TestClass2> {

        new TestClass2(){ Name = "南昌", Address = "长江中下游平原", Birth = "678-12-12", Phone = "123456789", Email = "1537004059@qq.com" },
        new TestClass2(){ Name = "九江", Address = "长江中下游平原", Birth = "678-01-01", Phone = "657712345", Email = "6666677778@qq.com" },
        new TestClass2(){ Name = "宜春", Address = "东南丘陵", Birth = "1056-07-12", Phone = "778812345", Email = "yiyandingzhen@qq.com" },
        new TestClass2(){ Name = "上饶", Address = "罗霄山北面", Birth = "1234-12-12", Phone = "666875652", Email = "5712351231@qq.com" },
        new TestClass2(){ Name = "赣州", Address = "南岭", Birth = "956-12-12", Phone = "98237818923", Email = "6154231@qq.com" },
        new TestClass2(){ Name = "萍乡", Address = "靠近湖南", Birth = "8293-12-12", Phone = "231231", Email = null },
        new TestClass2(){ Name = "萍乡", Address = "赣西地区", Birth = "8293-12-12", Phone = "114514", Email = "wobudaoa!" },
    };

    IWorkbook workbook = wrap.ConvertToExcel<TestClass2>(ts);

    MemoryStream ms = new MemoryStream();

    workbook.Write(ms);

    FileStream fs = new FileStream("wdnmd.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);

    fs.Write(ms.ToArray());

    fs.Close();


}
```

![示例图](http://img.leepichome.top/open1.png)

#### <font color=licyan>3.TitleAttribute标题头行标签</font>

> TitleAttribute是标题行的标签,它用来提供一个简单的标题设置,
>
> 和上一标签一样,你也可以用$1,...来设置一些动态的属性值(<font color=red>只读取list的第一个元素的属性替代这些通配符</font>)
>
> [同时也提供了一些<font color=darkcyan>标题字体的样式属性</font>,(其他等边框,对齐...等标签需单独使用在类上)]

| 属性名        | 类型   | 说明           |
| ------------- | ------ | -------------- |
| Context       | string | 标题的内容     |
| Col_Merge_num | int    | 需要合并的列数 |
| Row_Merge_num | int    | 需要合并的行数 |
| Single_Height | double | 单行的高度     |
| Font_Name     | string | 字体名称       |
| Font_Size     | int    | 字体大小       |
| Font_color    | byte[] | 字体颜色       |
| IsBold        | bool   | 是否加粗       |
| Back_color    | byte[] | 背景颜色       |
| Fore_color    | byte[] | 前景颜色       |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

同上