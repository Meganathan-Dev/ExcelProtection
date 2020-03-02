import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSF {
	public static void main(String[] args) {
		try (POIFSFileSystem fs = new POIFSFileSystem()) {
			EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
			// EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
			
			Encryptor enc = info.getEncryptor();
			enc.confirmPassword("password");

			// Read in an existing OOXML file and write to encrypted output stream
			File file = File.createTempFile("tempFile", ".xlsx");
			FileOutputStream fos = new FileOutputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook();

			//add sheet, rows and/or columns
			workbook.write(fos);
			fos.close();
			// don't forget to close the output stream otherwise the padding bytes aren't added
			try (OPCPackage opc = OPCPackage.open(file, PackageAccess.READ_WRITE);
			OutputStream os = enc.getDataStream(fs)) {
			opc.save(os);
			}
		}
	}
}
