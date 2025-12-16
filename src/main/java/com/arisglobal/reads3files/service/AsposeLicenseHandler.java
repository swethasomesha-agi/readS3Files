package com.arisglobal.reads3files.service;

import com.aspose.words.License;

import java.io.InputStream;

public class AsposeLicenseHandler {

    public void initialiseAsposeLicense(String licenseName) throws Exception {

        License license = new License();
        license.setLicense(licenseName);
    }


    public void initialiseAsposeLicense(InputStream inputStream) throws Exception {

        License license = new License();
        license.setLicense(inputStream);
    }

    public void initialiseAsposeLicenseForExcel(String licenseName) throws Exception {
        com.aspose.cells.License license = new com.aspose.cells.License();
        license.setLicense(licenseName);
    }

    public void initialiseAsposeLicenseForExcel(InputStream inputStream) throws Exception {
        com.aspose.cells.License license = new com.aspose.cells.License();
        license.setLicense(inputStream);
    }

    public void initialiseAsposeLicenseForPDFKit(String licenseName) throws Exception {
		com.aspose.pdf.License license = new com.aspose.pdf.License();
		license.setLicense(licenseName);
	}

    public void initialiseAsposeLicenseForPDFKit(InputStream inputStream) throws Exception {
        com.aspose.pdf.License license = new com.aspose.pdf.License();
        license.setLicense(inputStream);
    }

    public void initialiseAsposeLicenseForPDF(String licenseName) throws Exception {
        com.aspose.pdf.License license = new com.aspose.pdf.License();
        license.setLicense(licenseName);
    }

    public void initialiseAsposeLicenseForPDF(InputStream inputStream) throws Exception {
        com.aspose.pdf.License license = new com.aspose.pdf.License();
        license.setLicense(inputStream);
        System.out.println("Is Licensed: " + com.aspose.pdf.Document.isLicensed());
    }


}
