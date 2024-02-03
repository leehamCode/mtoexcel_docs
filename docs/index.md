# MToExcel Document

<b><font color=red size=4>这里是MToExcel的说明文档,你可以在这里找到所有API的说明信息。</font></b>

<b><font color=darkcyan>MToExcel是一个C#的List<>转Excel直接导出的工具。</font></b>

<b><font color=darkcyan>它主要用来减少NPOI样式代码,同时也保留了自定义的部分</font></b>

如果你需要简化Excel的导出(并存在较多的样式要求),你可以选择使用该工具

## :cyclone:使用方法

直接安装NuGet包使用,请查看如下链接
https://www.nuget.org/packages/MToExcel

:green_apple:示例代码:

<b>(测试类代码)</b>

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

<b>(导出类代码)</b>

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

![测试效果图](http://img.leepichome.top/mtoexcel.png)


#### 如果您想贡献代码，请查看以下仓库连接

> https://gitee.com/godenSpirit/mto-excel

> https://github.com/leehamCode/MToExcel

或者可以直接联系邮箱leeham51240@163.com