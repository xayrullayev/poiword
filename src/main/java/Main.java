import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
  public static void main(String[] args) throws Exception {
    //createSpreadSheet();
    //openSpreadSheet();
    createWordDoc();
  }

  private static void createSpreadSheet() throws IOException {
    //Create Blank workbook
    HSSFWorkbook workbook = new HSSFWorkbook();

    //Create file system using specific name
    FileOutputStream out = new FileOutputStream(new File("createworkbook.xlsx"));

    //write operation workbook using file out object
    workbook.write(out);
    out.close();
    System.out.println("createworkbook.xlsx written successfully");
  }

  private static void openSpreadSheet() throws IOException {
    File file = new File("openworkbook.xlsx");
    FileInputStream fIP = new FileInputStream(file);

    //Get the workbook instance for XLSX file
    HSSFWorkbook workbook = new HSSFWorkbook(fIP);
    workbook.write();
    if (file.isFile() && file.exists()) {
      System.out.println("openworkbook.xlsx file open successfully.");
    } else {
      System.out.println("Error to open openworkbook.xlsx file.");
    }
  }

  private static void createWordDoc() throws IOException {
    //Blank Document
    XWPFDocument document = new XWPFDocument();

    //Write the Document in file system
    FileOutputStream out = new FileOutputStream(new File("create.docx"));

    // create paragraph
    XWPFParagraph paragraph = document.createParagraph();
    RunnerDecorator run = new RunnerDecorator(paragraph.createRun());
    RunnerDecorator run2 = new RunnerDecorator(paragraph.createRun());
    RunnerDecorator run3 = new RunnerDecorator(paragraph.createRun());

    //create run
    run.setText("Lorem ipsum")
        .setBold(true)
        .setFontFamily("Verdana")
        .setFontSize(28)
        
        .setVerticalAlignment("center")
        .addBreak()
        .setFontSize(10)
        .addBreak()
        .build();

    run2.setText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ac faucibus odio.")
        .setBold(true)
        .setFontSize(18)
        .addBreak()
        .setFontSize(10)
        .addBreak()
        .build();

    RunnerDecorator runner = new RunnerDecorator(paragraph.createRun());
    runner.setText("Vestibulum neque massa, scelerisque sit amet ligula eu, congue molestie mi. Praesent ut varius sem. Nullam at porttitor arcu, nec lacinia nisi. Ut ac dolor vitae odio interdum condimentum.")
        .setBold(true)
        .setText("Vivamus dapibus sodales ex, vitae malesuada ipsum cursus convallis. Maecenas sed egestas nulla, ac condimentum orci.")
        .setBold(false)
        .setText("Mauris diam felis, vulputate ac suscipit et, iaculis non est. Curabitur semper arcu ac ligula semper, nec luctus nisl blandit. Integer lacinia ante ac libero lobortis imperdiet.")
        .setItalic(true)
        .setText("Nullam mollis convallis ipsum, ac accumsan nunc vehicula vitae.")
        .setItalic(false)
        .setText("Nulla eget justo in felis tristique fringilla. Morbi sit amet tortor quis risus auctor condimentum. Morbi in ullamcorper elit. Nulla iaculis tellus sit amet mauris tempus fringilla.")
        .addBreak()
        .build();


    RunnerDecorator runner2 = new RunnerDecorator(paragraph.createRun());
    runner2.setText("Maecenas mauris lectus, lobortis et purus mattis, blandit dictum tellus.")
        .addBreak()
        .build();


    document.write(out);
    out.close();
    System.out.println("createdoc.docx written successully");

  }
}
