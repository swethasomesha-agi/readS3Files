package com.arisglobal.reads3files;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.boot.autoconfigure.jdbc.DataSourceTransactionManagerAutoConfiguration;
import org.springframework.boot.autoconfigure.orm.jpa.HibernateJpaAutoConfiguration;
import org.springframework.context.ApplicationContext;

@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class, DataSourceTransactionManagerAutoConfiguration.class, HibernateJpaAutoConfiguration.class})
public class ReadS3FilesApplication {

    static String EMAIL_INBOUND_S3_PREFIX = "email_inbound";
    static String S3_SEPARTOR = "/";
    static final String DDE_FILE = "ddeInputFile.pdf";

    public static void main(String[] args) {
        ApplicationContext applicationContext = SpringApplication.run(ReadS3FilesApplication.class, args);
        WritePageNumbers writePageNumbers = applicationContext.getBean(WritePageNumbers.class);
        try {
//            writePageNumbers.emailAttachmentDetail();
//            writePageNumbers.editablePdfAttachmentDetail();
//            writePageNumbers.prepareFlattenedFiles();
            writePageNumbers.identifyMultilineInWordDoc();

//            writePageNumbers.processWordAttachmentTest("C:\\Sample_forms\\BMS\\AE-ORP-1572279 ORP0001101.docx",null);
//            writePageNumbers.processWordAttachmentTest("C:\\Sample_forms\\normal.docx", null);

//            writePageNumbers.updateResultSheet();
            System.exit(0);

//            writePageNumbers.readExcelFileRowByRow();
//            writePageNumbers.readExcelFileRowByRowForOriginal();
//            writePageNumbers.readExcelFileRowByRowCopyImpacted();
//            writePageNumbers.readExcelFileRowByRowCopyWithoutWorldWideBMS();
        } catch (Exception e) {
            System.out.println("Exception in Main" + e);
        }

    }

    public static String getPrefix(String messageUid) {
        return EMAIL_INBOUND_S3_PREFIX +
                S3_SEPARTOR + messageUid + S3_SEPARTOR;
    }

    public static String concatAll(Object... args) {
        StringBuilder concatinatedStrBuff = new StringBuilder();
        if (null != args && args.length > 0) {
            for (int i = 0; i < args.length; i++) {
                concatinatedStrBuff.append(args[i]);
            }
        }
        return concatinatedStrBuff.toString();
    }

}
