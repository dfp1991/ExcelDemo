package ExcelDemo;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.Iterator;
import java.util.Map;

public class TestDriverData {
    @Test(dataProvider = "excelIterator")
    public void f(Map<String,String> data){
        System.out.println("用例标题："+ data.get("comment"));
    }

    @DataProvider(name = "excelIterator")
    public Iterator<Object[]> dataProvider(){
        return new ExcelIterator("D:\\test.xls","testexcel");
    }
}
