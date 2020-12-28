import org.apache.commons.lang3.text.StrSubstitutor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.*;
import java.nio.file.Files;
import java.util.Map;

public class PandocConverter implements Converter {

    @Override
    public byte[] convert(File templateFile, Map<String, Object> data) throws Exception {
        final String tempPdfName = "temp.pdf";
        File tempDoc = new File("temp.docx");
        OutputStream out = new FileOutputStream(tempDoc);
        XWPFDocument temp = new XWPFDocument(new FileInputStream(templateFile));
        CTBody body = temp.getDocument().getBody();
        String mergedStr = StrSubstitutor.replace(body.xmlText(), data);
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        body.set(CTBody.Factory.parse(mergedStr, optionsOuter));
        temp.write(out);
        out.close();
        Runtime.getRuntime().exec(String.format(" pandoc \"%s\" --from=docx --to=pdf --output=\"%s\" ", tempDoc.getAbsolutePath(), tempPdfName));
        File tempPdf = new File(tempPdfName);
        byte[] buffer = Files.readAllBytes(tempPdf.toPath());
        tempDoc.delete();
        tempPdf.delete();
        return buffer;
    }
}
