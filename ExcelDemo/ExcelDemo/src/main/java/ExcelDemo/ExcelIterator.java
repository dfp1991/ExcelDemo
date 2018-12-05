package ExcelDemo;

import jxl.Cell;
import jxl.Workbook;
import jxl.Sheet;
import jxl.read.biff.BiffException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelIterator implements Iterator<Object[]> {

    private Workbook book = null;
    private Sheet sheet = null;

    private int rowNum = 0;
    private int curRowNo = 0;
    private int columnNum = 0;
    private String[] columnName;
    private int singleline = 0;

    public ExcelIterator(String filepath,String sheetname){
        try{
            FileInputStream fs = new FileInputStream(filepath);
            this.book = Workbook.getWorkbook(fs);
            this.sheet = book.getSheet(sheetname);
            this.rowNum = sheet.getRows();
            Cell[] c = sheet.getRow(0);
            this.columnNum = c.length;
            columnName = new String[c.length];
            for(int i = 0;i<c.length;i++){
                columnName[i] = c[i].getContents().toString();
            }
            if (this.singleline > 0){
                this.curRowNo = this.singleline;
            }else {
                this.curRowNo++;
            }
        }catch (FileNotFoundException e){
            System.out.println("文件IO异常");
        }catch (BiffException e){
            e.printStackTrace();
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    @Override
    public boolean hasNext(){
        if(this.rowNum == 0 || this.curRowNo >= this.rowNum){
            book.close();
            return false;
        }
        if(this.singleline > 0 && this.curRowNo >this.singleline){
            book.close();
            return false;
        }
        return true;
    }

    @Override
    public Object[] next(){
        Cell[] c = sheet.getRow(this.curRowNo);
        Map<String,Object> s = new LinkedHashMap<String, Object>();

        for(int i=0;i<this.columnNum;i++){
            String data = c[i].getContents().toString();
            s.put(this.columnName[i],data);
        }
        this.curRowNo++;
        return new Object[] {s};
    }

    @Override
    public void remove(){

    }
}
