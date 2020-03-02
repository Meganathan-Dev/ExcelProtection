import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class HSSF {

	public static void main(String[] args) throws IOException {
		Biff8EncryptionKey.setCurrentUserPassword("pass");
		POIFSFileSystem fs = new POIFSFileSystem(new File("file.xls"), true);
		HSSFWorkbook hwb = new HSSFWorkbook(fs.getRoot(), true);
		Biff8EncryptionKey.setCurrentUserPassword(null);
		
		Sheet sheet = hwb.createSheet();
		Row row1 = sheet.createRow(0);
		Cell cell = row1.getCell(0);
		cell.setCellValue("Test");
		
		hwb.close();
		fs.close();
	}
	
}
