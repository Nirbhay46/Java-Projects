
import java.io.FileOutputStream;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chunk;
import static com.itextpdf.text.Chunk.IMAGE;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.draw.VerticalPositionMark;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
//import sun.applet.Main;

public class MarkSheet extends MainFrame {

    private static Font catFont = new Font(Font.FontFamily.TIMES_ROMAN, 28,
            Font.BOLD);
    private static Font redFont = new Font(Font.FontFamily.TIMES_ROMAN, 21,
            Font.NORMAL, BaseColor.BLACK);
    private static Font redFont1 = new Font(Font.FontFamily.TIMES_ROMAN, 19,
            Font.NORMAL, BaseColor.BLACK);
    private static Font nameFont = new Font(Font.FontFamily.TIMES_ROMAN, 18,
            Font.BOLD, BaseColor.BLACK);
    private static Font subFont = new Font(Font.FontFamily.TIMES_ROMAN, 22,
            Font.BOLD);
    private static Font smallBold = new Font(Font.FontFamily.TIMES_ROMAN, 18,
            Font.BOLD);
    FileOutputStream file;
    int counter = 0, fetchedRow = 0;
    String roll = null;
    String enrollNo = null;
    String name = null;
    static String path1;
    int total = 0;

    public MarkSheet() {

    }

    public MarkSheet(String... enrollment) {
        //JOptionPane.showMessageDialog(null, college);
        URL url = null;
        if (college == "Lakshmi Narain College Of Technology") {
            url = MainFrame.class.getResource("/LNCT-BPL.png");
        }
        else if (college == "Oriental Institute Of Technology"){
            url = MainFrame.class.getResource("/oit.png");
        }
        File newpath = FileSystemView.getFileSystemView().getHomeDirectory();
        String path1 = newpath + "";
        MainFrame obj6 = new MainFrame();
        String FILE = path1 + "\\Marksheet.pdf";
        for (String a : enrollment) {
            counter++;
        }
        try {
            Document document = new Document(PageSize.A4.rotate(), 0, 0, 0, 0);
            FileOutputStream file = new FileOutputStream(FILE);
            PdfWriter writer = PdfWriter.getInstance(document, file);
            document.open();
            //JOptionPane.showMessageDialog(null, counter+" ");

            for (String b : enrollment) {
                Paragraph paragraph = new Paragraph(college, catFont);
                paragraph.setAlignment(Element.ALIGN_CENTER);
                Paragraph paragraph1 = new Paragraph("MarkSheet and progress report of MID SEM " + midsem + " Exam",
                        redFont);
                paragraph1.setAlignment(Element.ALIGN_CENTER);
                Paragraph para1 = new Paragraph("Branch : " + branch + " (Section " + section + ")", subFont);
                para1.setAlignment(Element.ALIGN_CENTER);
                Paragraph paragraph2 = new Paragraph();
                addEmptyLine(paragraph2, 1);

                PdfContentByte canvas = writer.getDirectContentUnder();
                Image image = Image.getInstance(url);
                image.scaleAbsolute(300, 300);
                image.setAbsolutePosition(260f, 160f);

                canvas.saveState();
                PdfGState state = new PdfGState();
                state.setFillOpacity(0.2f);
                canvas.setGState(state);
                canvas.addImage(image);
                canvas.restoreState();
                PdfPTable table = new PdfPTable(8);
                table.setTotalWidth(new float[]{50, 150, 60, 105, 60, 80, 75, 100});
                table.setLockedWidth(true);
                PdfPCell c1 = new PdfPCell(new Phrase("S.No"));
                c1.setColspan(1);
                c1.setPadding(10);
                c1.setHorizontalAlignment(Element.ALIGN_CENTER);
                table.addCell(c1);
                c1 = new PdfPCell(new Phrase("Subject"));
                c1.setColspan(1);
                c1.setPadding(10);
                c1.setHorizontalAlignment(Element.ALIGN_CENTER);
                table.addCell(c1);
                c1 = new PdfPCell(new Phrase("Marks"));
                c1.setColspan(2);
                c1.setPadding(10);
                c1.setHorizontalAlignment(Element.ALIGN_CENTER);
                table.addCell(c1);
                c1 = new PdfPCell(new Phrase("Attendance Theory"));
                c1.setColspan(2);
                c1.setPadding(10);
                c1.setHorizontalAlignment(Element.ALIGN_CENTER);
                table.addCell(c1);
                c1 = new PdfPCell(new Phrase("Attendance Practical"));
                c1.setColspan(2);
                c1.setPadding(10);
                c1.setHorizontalAlignment(Element.ALIGN_CENTER);
                table.addCell(c1);
                table.addCell(" ");
                table.addCell("  ");
                table.addCell("Max");
                table.addCell("Obtained");
                table.addCell("Held");
                table.addCell("Attendance");
                table.addCell("Held");
                table.addCell("Attendance");

                try {
                    File obj = new File(path);
                    Workbook w = Workbook.getWorkbook(obj);
                    Sheet sheet = w.getSheet(0);
                    for (int k = 0; k < sheet.getRows(); k++) {
                        Cell enrollment1 = sheet.getCell(1, k);
                        String id = enrollment1.getContents();

                        if (b.equals(id)) {

                            fetchedRow = k;
                            roll = sheet.getCell(0, fetchedRow).getContents();
                            enrollNo = sheet.getCell(1, fetchedRow).getContents();
                            name = sheet.getCell(2, fetchedRow).getContents();

                        }

                    }
                    PdfPCell c2 = new PdfPCell();
                    for (int j = 1; j < (numberOfSubjects + 1); j++) {
                        c2 = new PdfPCell(new Phrase(j + ""));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 2, 0).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(maxMarks + ""));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 2, fetchedRow).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 9, fetchedRow).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 16, fetchedRow).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 23, fetchedRow).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                        c2 = new PdfPCell(new Phrase(sheet.getCell(j + 30, fetchedRow).getContents()));
                        c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                        c2.setPadding(10);
                        table.addCell(c2);
                    }
                    c2 = new PdfPCell(new Phrase("Grand Total"));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setColspan(2);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(maximumMarks + ""));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setColspan(1);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(sheet.getCell(numberOfSubjects + 3, fetchedRow).getContents()));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(sheet.getCell(numberOfSubjects + 10, fetchedRow).getContents()));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(sheet.getCell(numberOfSubjects + 17, fetchedRow).getContents()));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(sheet.getCell(numberOfSubjects + 24, fetchedRow).getContents()));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setPadding(10);
                    table.addCell(c2);
                    c2 = new PdfPCell(new Phrase(sheet.getCell(numberOfSubjects + 31, fetchedRow).getContents()));
                    c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c2.setPadding(10);
                    table.addCell(c2);

                } catch (BiffException e) {
                    e.printStackTrace();
                } catch (IOException ex) {
                    Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
                }
                PdfPTable table3 = new PdfPTable(3);
                table3.setTotalWidth(new float[]{210, 260, 295});
                table3.setLockedWidth(true);
                table3.setWidthPercentage(80);
                PdfPCell c4 = new PdfPCell(new Phrase("Semester :" + semester, redFont1));
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setBorderColor(BaseColor.WHITE);
                c4.setColspan(1);
                table3.addCell(c4);
                c4 = new PdfPCell(new Phrase(" "));
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setBorderColor(BaseColor.WHITE);
                c4.setColspan(1);
                table3.addCell(c4);
                c4 = new PdfPCell(new Phrase("Academic Session : " + session, redFont1));
                c4.setBorderColor(BaseColor.WHITE);
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setColspan(1);
                table3.addCell(c4);
                c4 = new PdfPCell(new Phrase("Class Roll no :" + roll, redFont1));
                c4.setBorderColor(BaseColor.WHITE);
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setColspan(1);
                table3.addCell(c4);
                c4 = new PdfPCell(new Phrase("Enrollment No :" + enrollNo, redFont1));
                c4.setBorderColor(BaseColor.WHITE);
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setColspan(1);
                table3.addCell(c4);
                c4 = new PdfPCell(new Phrase("Name : " + name, nameFont));
                c4.setBorderColor(BaseColor.WHITE);
                c4.setHorizontalAlignment(Element.ALIGN_CENTER);
                c4.setColspan(1);
                table3.addCell(c4);
                document.add(paragraph);
                document.add(paragraph1);
                document.add(para1);
                document.add(paragraph2);
                document.add(table3);
                document.add(table);
                Phrase behavior = new Phrase();
                behavior.add(new Chunk("                   " + "General Behavior:  ", smallBold));
                behavior.add(new Chunk("Poor/Good/Very Good." + "                                                                                                        ", redFont1));
                Phrase advice = new Phrase();
                advice.add(new Chunk("     " + "Advice Him/Her To: ", smallBold));
                advice.add(new Chunk(" Improve the Performance/ Improve Regularity/ Work Hard.", redFont1));
                document.add(behavior);
                document.add(advice);
                Paragraph verify = new Paragraph("                  Verified by :", redFont1);
                document.add(verify);

                PdfPTable table1 = new PdfPTable(3);
                table1.setWidthPercentage(80);
                PdfPCell c2 = new PdfPCell(new Phrase("Class Officer", smallBold));
                c2.setHorizontalAlignment(Element.ALIGN_LEFT);
                c2.setBorderColor(BaseColor.WHITE);
                c2.setColspan(1);
                table1.addCell(c2);
                c2 = new PdfPCell(new Phrase(HOD, smallBold));
                c2.setHorizontalAlignment(Element.ALIGN_CENTER);
                c2.setBorderColor(BaseColor.WHITE);
                c2.setColspan(1);
                table1.addCell(c2);
                c2 = new PdfPCell(new Phrase("Dr Rakesh Mowar", smallBold));
                c2.setBorderColor(BaseColor.WHITE);
                c2.setHorizontalAlignment(Element.ALIGN_RIGHT);
                c2.setColspan(1);
                table1.addCell(c2);
                document.add(table1);

                PdfPTable table2 = new PdfPTable(3);
                table2.setWidthPercentage(80);
                PdfPCell c3 = new PdfPCell(new Phrase("  "));
                c3.setHorizontalAlignment(Element.ALIGN_LEFT);
                c3.setBorderColor(BaseColor.WHITE);
                c3.setColspan(1);
                table2.addCell(c3);
                c3 = new PdfPCell(new Phrase("(HOD)", redFont1));
                c3.setHorizontalAlignment(Element.ALIGN_CENTER);
                c3.setBorderColor(BaseColor.WHITE);
                c3.setColspan(1);
                table2.addCell(c3);
                c3 = new PdfPCell(new Phrase("Principal LNCT,Bhopal", redFont1));
                c3.setBorderColor(BaseColor.WHITE);
                c3.setHorizontalAlignment(Element.ALIGN_RIGHT);
                c3.setColspan(1);
                table2.addCell(c3);
                document.add(table2);

                document.newPage();

            }

            document.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (Desktop.isDesktopSupported()) {
            try {
                File file = new File(FILE);
                Desktop.getDesktop().open(file);
                //Desktop.getDesktop().browse(new URI("E:\\test\\hello.pdf"));
            } catch (IOException ex) {
                Logger.getLogger(MarkSheet.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

    }

    private void addEmptyLine(Paragraph paragraph1, int number) {
        for (int i = 0; i < number; i++) {
            paragraph1.add(new Paragraph(" "));
        }

    }
}
