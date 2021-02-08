package POI;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

public class POIActiveCellSample {

    public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
          String filePath = "testfile.xlsx";
          Workbook workbook = null;
          FileInputStream in = null;
          OutputStream os = null;
          try {
                in = new FileInputStream(filePath);
                workbook = WorkbookFactory.create(in);

                Sheet sheet = workbook.getSheetAt(0);

                // �����ŃA�N�e�B�u�ȃZ����ݒ�
                // ���̗�ł�3�s�ڂ�3��ڂ��A�N�e�B�u�Z���ɐݒ�i�s�E���0����n�܂�̂ň����́u2,2�v�ƂȂ�j
                CellAddress address = new CellAddress(2, 2);
                sheet.setActiveCell(address);

                os = new FileOutputStream(filePath);
                workbook.write(os);

          } finally {
                if (in != null) {
                      in.close();
                }
                if (os != null) {
                      os.close();
                }
                if (workbook != null) {
                      workbook.close();
                }
          }
    }

}