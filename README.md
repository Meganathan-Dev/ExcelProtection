This method of creating Excel is based on Apache-POI.

XLS is a proprietary binary format whereas XLSX is based on open XML format. In XLS, the information is directly stored to a binary format BIFF (Binary Interchange File Format). 

Whereas the information in an XLSX file is stored in a text file that uses XML to define all its parameters. So we need different Workbook object depending on the Excel format as shown below.


XLSX - XSSFWorkbook()

XLS - HSSFWorkbook()


---

XLS


Creating Excel-XLS is simple as shown below.

Biff8EncryptionKey.setCurrentUserPassword("pass");

POIFSFileSystem fs = new POIFSFileSystem(new File("file.xls"), true);

HSSFWorkbook hwb = new HSSFWorkbook(fs.getRoot(), true);

Biff8EncryptionKey.setCurrentUserPassword(null);



---

XLSX


Creating Excel-XLSX is tricky than Excel-XLS.

try (POIFSFileSystem fs = new POIFSFileSystem()) {

  EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);

  // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);

  Encryptor enc = info.getEncryptor();

  enc.confirmPassword("foobaa");

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
