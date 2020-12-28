import org.apache.commons.lang3.text.StrSubstitutor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.*;
import java.util.Map;

public class ApachePoiJodConverter implements Converter {

    @Override
    public byte[] convert(File templateFile, Map<String, Object> data) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        XWPFDocument temp = new XWPFDocument(new FileInputStream(templateFile));
        CTBody body = temp.getDocument().getBody();
        String mergedStr = StrSubstitutor.replace(body.xmlText(), data);
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        body.set(CTBody.Factory.parse(mergedStr, optionsOuter));
        temp.write(out);

        InputStream in = new ByteArrayInputStream(out.toByteArray());
        DocumentConverter converter = DocumentConverterInstanceGenerator.getInstance();
        ByteArrayOutputStream finalOutput = new ByteArrayOutputStream();
        converter.convert(in).to(finalOutput).as(DefaultDocumentFormatRegistry.PDF).execute();
        return finalOutput.toByteArray();
    }

    private static class DocumentConverterInstanceGenerator {
        private static DocumentConverter instance;

        private DocumentConverterInstanceGenerator() {
        }

        public static DocumentConverter getInstance() throws OfficeException {
            if (instance == null) {
                OfficeManager LIBRE_OFFICE = LocalOfficeManager.builder().officeHome("E:\\pf\\LibreOffice").install().build();
                LIBRE_OFFICE.start();
                instance = LocalConverter.make(LIBRE_OFFICE);
            }
            return instance;
        }
    }
}
