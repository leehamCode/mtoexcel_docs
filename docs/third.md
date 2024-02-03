## Property Attribute列表

<small>只有属性可以使用</small>

#### <font color=licyan>1.HeaderName 表头文本标签</font>

> 此标签用于作为Excel<b>表头上的文本</b>内容

| 属性名 | 类型   | 说明     |
| ------ | ------ | -------- |
| name   | string | 表头文本 |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
public class TestClass2
{
    [HeaderName("你的姓名:")]
    public string Name { get; set; }
    
    /* 省略大部分 */
    
}
```

![例图](http://img.leepichome.top/mtoexcel.png)

#### <font color=licyan>2.IgnoreType 属性忽略标签</font>

> 此标签用于将<b>不想打印的属性忽略</b>

| 属性名       | 类型 | 说明         |
| ------------ | ---- | ------------ |
| isTrueIgnore | bool | 是否真实忽略 |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
public class Person
{
    [HeaderName("学生学号")]
    public string id { get; set; }

    public string name { get; set; }

    [IgnoreType]
    public float tall { get; set; }

    [ReferenceType(true)]
    public Animal pet { get; set; }
}
```

![例图2](http://img.leepichome.top/Mytest.png)

#### <font color=licyan>3.MergeNearEqualBox 同列临接相同属性合并标签[只实现了同列]</font>

> 此标签用于将<b>相同属性合并</b>b为大单元格(前提为已完成集合的排序工作)

| 属性名  | 类型 | 说明                             |
| ------- | ---- | -------------------------------- |
| OnlyCol | bool | 同列临近相同值单元格合并         |
| OnlyRow | bool | 同行临近相同值单元格合并[未实现] |
| Both    | bool | [未实现]                         |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
public class TestClass2
{
    [HeaderName("你的姓名:")]
    [MergeNearEqualBox]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Left)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Bottom)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Right)]
    [BorderStyle(Models.Enums.BorderWid.MiddBorder,new byte[]{44,56,179},Models.Enums.BorderDirect.Upper)]

    public string Name { get; set; }
    
    /* 省略其他部分 */
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

#### <font color=licyan>4. BackForeColor背景颜色标签</font>

> 用于设置单元格的背景/前景颜色

| 属性名   | 类型   | 说明          |
| -------- | ------ | ------------- |
| back_rgb | byte[] | 背景颜色的RGB |
| fore_rgb | byte[] | 前景颜色的RGB |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

```c#
[TitleAttribute(Context ="测试打印标题行$1",Font_Name ="新細明體",Font_Size =16,Font_color =new byte[]{ 80,235,227 })]
[Hide_On_Condition(rowCondition ="$1==\"九江\"")]
public class TestClass2
{
    [HeaderName("你的姓名:")]
    public string Name { get; set; }

    [HeaderName("地址:")]
    [BackForeColor(true,new byte[] {94,89,244})]
    [BorderStyle(MToExcel.Models.Enums.BorderWid.ThinBorder,new byte[] { 252,28,3 },MToExcel.Models.Enums.BorderDirect.Upper)]
    [DynaRowColumnLen(false,400.00)]
    public string Address { get; set; }

    [HeaderName("手机号")]
    [FontSets("Brush Script MT",14,true,true,true,true,new byte[] { 119,251,232 })]
    public string Phone { get; set; }

    [HeaderName("生日")]
    [Horizon(MToExcel.Models.Enums.Horizon.Left,MToExcel.Models.Enums.VerticalHorizon.Up)]
    public string Birth { get; set; }

    [HeaderName("邮箱")]
    public string Email { get; set; }
}
```

```c#
static void Main(string[] args)
{
    Console.WriteLine("Hello World!");

    WrapperConverter wrap = new WrapperConverter();
    wrap.basic = new BasicConverter();

    List<TestClass> ts = new List<TestClass> {

        new TestClass(){ Name = "南昌", Address = "长江中下游平原", Birth = "678-12-12", Phone = "123456789", Email = "1537004059@qq.com" },
        new TestClass(){ Name = "九江", Address = "长江中下游平原", Birth = "678-01-01", Phone = "657712345", Email = "6666677778@qq.com" },
        new TestClass(){ Name = "宜春", Address = "东南丘陵", Birth = "1056-07-12", Phone = "778812345", Email = "yiyandingzhen@qq.com" },
        new TestClass(){ Name = "上饶", Address = "罗霄山北面", Birth = "1234-12-12", Phone = "666875652", Email = "5712351231@qq.com" },
        new TestClass(){ Name = "赣州", Address = "南岭", Birth = "956-12-12", Phone = "98237818923", Email = "6154231@qq.com" },
        new TestClass(){ Name = "萍乡", Address = "靠近湖南", Birth = "8293-12-12", Phone = "231231", Email = "leehan51240@qq.com" },
    };

    IWorkbook workbook = wrap.ConvertToExcel<TestClass>(ts);

    MemoryStream ms = new MemoryStream();

    workbook.Write(ms);

    FileStream fs = new FileStream("wdnmd.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
    fs.Write(ms.ToArray());
    fs.Close();
}
```

![图例](http://img.leepichome.top/mtoexcel.png)

#### <font color=licyan>5.DynaRowColumnLen单元格长宽标签</font>

> 用于设置单元格的长度(Column Length)和宽度(Row Height)
>
> [如果在一个类中设置了<font color=red>多个Row Height</font>,那么将以最后那个为准]

| 属性名    | 类型   | 说明 |
| --------- | ------ | ---- |
| RowHeight | double | 行高 |
| ColLength | double | 列宽 |

<font color=orang><b><u>:green_apple:使用例:</u></b></font>

同上

![图例](http://img.leepichome.top/mtoexcel.png)

#### <font color=licyan>6.FontSets字体设置标签</font>

> 此标签用于设置单元格使用的字体的样式

| 属性名      | 类型   | 说明                        |
| ----------- | ------ | --------------------------- |
| Name        | string | 字体名称                    |
| Size        | double | 字体大小                    |
| IsBold      | bool   | 是否加粗                    |
| IsItalic    | bool   | 是否倾斜                    |
| IsUnderline | bool   | 是否下划线                  |
| IsStrikeout | bool   | 是否中间线                  |
| FontColor   | byte[] | 字体的颜色                  |
| Dataformat  | string | 设置字体格式(设置千分位...) |
|             |        |                             |

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

    [HeaderName("地址:")]
    [BackForeColor(true,new byte[] {94,89,244})]
    [BorderStyle(MToExcel.Models.Enums.BorderWid.ThinBorder,new byte[] { 252,28,3 },MToExcel.Models.Enums.BorderDirect.Upper)]
    [DynaRowColumnLen(false,400.00)]
    public string Address { get; set; }

    [HeaderName("手机号")]
    [FontSets("Brush Script MT",14,true,true,true,true,new byte[] { 119,251,232 })]
    public string Phone { get; set; }

    [HeaderName("生日")]
    [Horizon(MToExcel.Models.Enums.Horizon.Left,MToExcel.Models.Enums.VerticalHorizon.Up)]
    public string Birth { get; set; }

    [HeaderName("邮箱")]
    public string Email { get; set; }
}
```

![图例](http://img.leepichome.top/open1.png)

