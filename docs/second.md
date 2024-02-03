## Getting Started开始使用

#### :star:1.添加NuGet依赖

<small><font color=licyan>如果你使用.NetCore3.1,请使用1.0.3版本</font></small>

> dotnet add package MToExcel --version 1.0.3

<small><font color=licyan>如果你使用.Net6,请使用1.0.2版本</font></small>

> dotnet add package MToExcel --version 1.0.2

####  :sailboat:2.定义你的Model

```csharp
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

#### :taco:3.填充数据并转化

```c#
static void Main(string[] args)
{
    Console.WriteLine("Hello World!");

    WrapperConverter wrap = new WrapperConverter();
    wrap.basic = new BasicConverter();
	
    //一般来说可以将ORM框架从数据库中查到的对象List<>直接返回
    List<TestClass2> ts = new List<TestClass2> {

        new TestClass2(){ Name = "南昌", Address = "长江中下游平原", Birth = "678-12-12", Phone = "123456789", Email = "1537004059@qq.com" },
        new TestClass2(){ Name = "九江", Address = "长江中下游平原", Birth = "678-01-01", Phone = "657712345", Email = "6666677778@qq.com" },
        new TestClass2(){ Name = "宜春", Address = "东南丘陵", Birth = "1056-07-12", Phone = "778812345", Email = "yiyandingzhen@qq.com" },
        new TestClass2(){ Name = "上饶", Address = "罗霄山北面", Birth = "1234-12-12", Phone = "666875652", Email = "5712351231@qq.com" },
        new TestClass2(){ Name = "赣州", Address = "南岭", Birth = "956-12-12", Phone = "98237818923", Email = "6154231@qq.com" },
        new TestClass2(){ Name = "萍乡", Address = "靠近湖南", Birth = "8293-12-12", Phone = "231231", Email = "leehan51240@qq.com" },
    };

    IWorkbook workbook = wrap.ConvertToExcel<TestClass2>(ts);

    MemoryStream ms = new MemoryStream();

    workbook.Write(ms);

    FileStream fs = new FileStream("wdnmd.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);

    fs.Write(ms.ToArray());

    fs.Close();


}
```

