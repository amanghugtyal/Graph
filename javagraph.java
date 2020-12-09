
import java.awt.Color;
import org.jfree.chart.*;
import org.jfree.chart.plot.*;
import org.jfree.data.category.DefaultCategoryDataset;
import java.io.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

public class javagraph
{

    private static Workbook wb;
    private static Sheet sh;
    private static FileInputStream fis;
    private static Row rw;

    public static void main(String args[]) throws FileNotFoundException, IOException
    {
        try
        {
            DefaultCategoryDataset dcd = new DefaultCategoryDataset();
            fis = new FileInputStream("./marks.xlsx");
            wb = WorkbookFactory.create(fis);
            sh = wb.getSheet("Sheet1");
            int rows = sh.getLastRowNum();
            for (int i = 1; i <= rows; i++)
            {
                rw = sh.getRow(i);
                int colomn = rw.getPhysicalNumberOfCells();
                for (int j = 0; j <= colomn - 2; j++)
                {
                    DataFormatter format = new DataFormatter();
                    Object num = format.formatCellValue(rw.getCell(j + 1));
                    int n = Integer.parseInt((String) num);
                    Object name = format.formatCellValue(rw.getCell(j));
                    dcd.setValue(n, "Marks", (Comparable) name);
                }
            }
            JFreeChart jchart = ChartFactory.createBarChart3D("Student Record", "Student Name", "Student Marks", dcd, PlotOrientation.VERTICAL, true, true, false);
            CategoryPlot plot = jchart.getCategoryPlot();
            plot.setRangeGridlinePaint(Color.blue);
            ChartFrame frame = new ChartFrame("Student Record", jchart, true);
            frame.setVisible(true);
            frame.setSize(800, 700);
            fis.close();
        } catch (IOException | NumberFormatException | EncryptedDocumentException ex)
        {
            System.out.println(ex);
        }
    }
}
