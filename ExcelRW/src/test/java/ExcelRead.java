import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

/**
 * @author ak
 * @version 1.0
 * @date 2024/4/2 14:59
 */
public class ExcelRead {

    static String XLSXReadPath = "E:\\TestData\\POIRead.xlsx";
    static String XLSXWritePath = "";
    static String BIGXLSXReadPath = "";
    static String BIGXLSXWritePath = "";
    static String XLSReadPath = "E:\\TestData\\ComputerComponent.xls";
    static String XLSWritePath = "";

    @Test
    public void XLSXRead() throws IOException {
        /*一个sheet最大行数1048576，最大列数16384。*/
        DoXLSXRead();
    }

    @Test
    public void BIGXLSXRead() throws IOException {
        /*大文件*/
        DoBIGXLSXRead();
    }

    @Test
    public void XLSRead() throws IOException {
        /*一个sheet最大行数65536，最大列数256。*/
        DoXLSRead();
    }

    private void DoXLSXRead() throws IOException {
        FileInputStream fis = new FileInputStream(XLSXReadPath);
        XSSFWorkbook wb = new XSSFWorkbook(fis);

        for (Sheet sheet : wb) {
            getContent(sheet);
        }

        // 关闭流
        fis.close();
        wb.close();

    }


    private void DoBIGXLSXRead() throws IOException {
        // SXSSF
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(XLSXReadPath));
        SXSSFWorkbook wb = new SXSSFWorkbook(xssfWorkbook);
        for (Sheet sheet : wb) {
            getContent(sheet);
        }

    }

    private void DoXLSRead() throws IOException {
        // HSSF
        POIFSFileSystem pfs = new POIFSFileSystem(new File(XLSReadPath));
        HSSFWorkbook wb = new HSSFWorkbook(pfs.getRoot(), true);

        for (Sheet sheet : wb) {
            getContent(sheet);
        }

    }

    // 依次读取sheet 行 单元格 值 获取excel表格内容
    private void getContent(Sheet sheet)  {

            for (Row cells : sheet) {
                for (Cell cell : cells) {
                    // numerical to string
                    cell.setCellType(CellType.STRING);
                    String cellValue = cell.getStringCellValue();

                    // 换行处理等
                    System.out.println(cellValue);
                    //txtFile(cellValue);
                }
        }
    }

    private void txtFile(String cellValue) {
        FileWriter fw = null;
        try {
            fw = new FileWriter("E:\\TestData\\any.txt", true);
            fw.write(cellValue);
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                fw.flush();
                fw.close();
                System.out.println("write success");
            }catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
