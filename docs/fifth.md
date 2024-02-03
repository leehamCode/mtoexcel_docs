## 可共用Attribute列表

<small>类和属性都可以使用</small>

#### <font color=licyan>1.DateTimeFormat 指定日期格式标签</font>

> 当Excel中需要特定的日期格式输出时,可以使用该标签.
>
> ---on property 为当前日期属性
>
> ---on class 为该class的所有日期属性

| 属性名 | 类型   | 说明     |
| ------ | ------ | -------- |
| format | string | 日期格式 |

<font color=orang><b><u>:taco:使用例:</u></b></font>

```c#
/// <summary>
/// 日期格式转化测试Model
/// </summary>
public class TestClass3
{
    [HeaderName("学校名称")]
    public string Name{get;set;}

    [HeaderName("学校地址")]
    public string Region{get;set;}

    [HeaderName("创办日期")]
    [DateTimeFormat(format ="yyyy-MM-dd")]
    [FontSets("標楷體",16,true,false,false,false,new byte[]{233,236,0})]
    public DateTime Create_date{get;set;}

    [HeaderName("排名")]
    public int Rank{get;set;}

    [HeaderName("校领导")]
    public string head_teacher{get;set;}

    [HeaderName("备注")]
    public string remark{get;set;}
}
```

```c#
public static void TestFour()
{
    List<TestClass3> listOne = new List<TestClass3>(){
        new TestClass3(){ Name = "江西农业大学",Region="江西省南昌市新建区",Create_date = new DateTime(1905,11,5),Rank = 6,head_teacher = "预演丁真",remark="A"},
        new TestClass3(){ Name = "南昌大学",Region="江西省南昌市红谷滩区",Create_date = new DateTime(1915,11,5),Rank = 1,head_teacher = "遗言丁真",remark="B"},
        new TestClass3(){ Name = "江西财经大学",Region="江西省南昌市新建区",Create_date = new DateTime(1955,11,5),Rank = 2,head_teacher = "云隐丁真",remark="C"},
        new TestClass3(){ Name = "华东交通大学",Region="江西省南昌市新建区",Create_date = new DateTime(1965,11,5),Rank = 6,head_teacher = "音乐丁真",remark="D"},
        new TestClass3(){ Name = "江西师范大学",Region="江西省南昌市青山湖区",Create_date = new DateTime(1975,11,5),Rank = 6,head_teacher = "阴影丁真",remark="E"}

    };

    WrapperConverter wrapper = new WrapperConverter();

    wrapper.basic = new BasicConverter();

    IWorkbook workbook = wrapper.ConvertToExcel<TestClass3>(listOne);

    FileStream file = new FileStream("DEMO.xlsx", FileMode.Create);

    workbook.Write(file);

    file.Close();
}
```

![图例](http://img.leepichome.top/open3.png)

#### <font color=licyan>2.BorderStyleAttribute边框样式标签</font>

> 该边框用于设置Excel各边框的样式,如厚度,颜色,方向,
>
> -----如果需要<font color=red>不同边框设置不同样式(比如左边框厚,右边框薄),需设置多项标签</font>

| 属性名       | 类型     | 说明       |
| ------------ | -------- | ---------- |
| BorderWid    | 枚举类型 | 边框的厚度 |
| BorderDirect | 枚举类型 | 边框的方向 |
| Color        | byte[]   | 边框的颜色 |

<font color=orang><b><u>:taco:使用例:</u></b></font>

```c#
public class TestClass2
{
    //如果需要设置多个方向的边框则设置多个边框Attribute
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
    
    /* 省略 */
}
```

![图例](http://img.leepichome.top/open1.png)

#### <font color=licyan>3.Horizon对齐设置标签</font>

> 对齐标签控制内容的对齐方式,
>
> -----on property <font color=red>控制单元格的对齐方式</font>
>
> -----on class <font color=red>控制标题标签(TitleAttribute)的对齐方式</font>

| 属性名          | 类型 | 说明           |
| --------------- | ---- | -------------- |
| Horizon         | 枚举 | 水平对齐的方式 |
| VerticalHorizon | 枚举 | 垂直对齐的方式 |

<font color=orang><b><u>:taco:使用例:</u></b></font>

```c#
[FreezeArea(FreezeStartCol =3,FreezeStartRow = 3)]
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
    //[CellStyle(Models.Enums.Horizon.Center, Models.Enums.VerticalHorizon.Up, false, charSet = new CharSet() { Size = 13.1d })]
    public string phone{get;set;}
}
```

![图例](http://img.leepichome.top/open4.png)

