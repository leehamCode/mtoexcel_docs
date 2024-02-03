## <font color=darkcyan>自定义表头/表尾,和控制行打印事件</font>

><font color=red>当需要设置的表头比较复杂,或者需要动态的设置表头又或者需要在表尾(行尾)计算统计行时</font>,
>
>可以传递Action进行额外的处理

```c#
/// <summary>
/// 自定义表头函数
/// </summary>
/// <value>转化中需要导出的Workbook</value>
/// <value>自定义表头需要占用的行数，如果没有，则默认一行</value>
public Action<IWorkbook> CustomHeadMethod{get;set;} = null;

/// <summary>
/// 自定义表头需要占用的行数
/// </summary>
public int? CustomHeadRows {get;set;} = null;   //使用自定义表头必须设置


/// <summary>
/// 自定义Excel尾部函数
/// </summary>
/// <value>Excel本体</value>
/// <value>单个Sheet的最后一行</value>
public Action<IWorkbook,int> CustomTailMethod{get;set;} = null;

/// <summary>
/// 行打印前事件
/// </summary>
/// <value>当前行,object为T类型当前行对象</value>
public Action<IRow,Object> item_change_before_event { get; set; }


/// <summary>
/// 打印后事件
/// </summary>
/// <value>当前行,object为T类型当前行对象</value>
public Action<IRow,Object> item_change_after_event{get;set;}
```

它们分别在打印的不同节点被执行,如下为自定义表头的执行时机

```c#
ISheet defaultSheet = workbook.CreateSheet("SheetOne");

//获取传递的泛型类型
Type type = typeof(T);
Check_Class_Attr(type);
//首先判断泛型T是否为基础数据类型

//如果泛型类型为基础数据类型,则为写一行数据
if (isBasicType(type))
{
    IRow uniqueRow = defaultSheet.CreateRow(0);
    int Count = 0;
    list.ForEach(item => {
        uniqueRow.CreateCell(Count).SetCellValue(Convert.ToString(item));
        Count++;
    });
    return workbook;
}

//如果不是基础数据类型就反射获取其属性写入Excel

PropertyInfo[] properties = type.GetProperties();

//-------------------------------------------------------------------------------------------------------------------分割线

if(CustomHeadMethod!=null&&CustomHeadRows!=null)
{
    if(CustomHeadRows<=0)
    {
        throw new CustomHeadException("自定义表头长度必须大于0!");
    }
    CustomHeadMethod.Invoke(workbook);
}
```

用例:

```c#
public class CustomTestClass
{
    [HeaderName("省份名称")]
    public string Name{get;set;}

    [HeaderName("旧时名称")]
    public string OldName{get;set;}

    [HeaderName("何处")]
    public string Address{get;set;}

    [HeaderName("河流")]
    public string River{get;set;}

    [HeaderName("山川")]
    public string Mountain{get;set;}

}
```



```c#
public static void CutomTestFive()
{
    List<CustomTestClass> listOne = new List<CustomTestClass>(){
        new CustomTestClass(){ Name = "江西", OldName = "江右/江南西道",Address="豫章", River = "赣江", Mountain="武夷山"},
        new CustomTestClass(){ Name = "湖北", OldName = "山南东道",Address="荆襄", River = "长江", Mountain="秦岭"},
        new CustomTestClass(){ Name = "甘肃", OldName = "陇右道",Address="河西", River = "黄河", Mountain="祁连山"},
        new CustomTestClass(){ Name = "山西", OldName = "河东道",Address="河东", River = "汾河", Mountain="武夷山"},

    };

    WrapperConverter wrapper = new WrapperConverter();


    wrapper.basic = new BasicConverter();
    wrapper.basic.CustomHeadRows = 3;
    wrapper.basic.CustomHeadMethod = (workbook)=>{
        var sheet = workbook.GetSheetAt(0);

        for(int i = 0;i<3;i++)
        {
            var row =  sheet.CreateRow(i);

            row.CreateCell(i).SetCellValue("wdnmd");
        }

    };
    wrapper.basic.CustomTailMethod = (workbook,LastRow)=>{
        var sheet = workbook.GetSheetAt(0);
        var row =  sheet.CreateRow(LastRow);
        row.CreateCell(0).SetCellValue("音乐丁真");
    };

    IWorkbook workbook = wrapper.ConvertToExcel<CustomTestClass>(listOne);

    FileStream file = new FileStream("DEMO.xlsx", FileMode.Create);

    workbook.Write(file);

    file.Close();
}
```

结果:

![图例](https://raw.githubusercontent.com/leehamCode/MyPics/main/open5.png)

