import org.apache.commons.lang3.text.StrSubstitutor;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.*;
import java.util.Map;

public class ApachePoiConverter implements Converter {

    @Override
    public byte[] convert(File templateFile, Map<String, Object> data) throws Exception {
        return getMergedPDF(templateFile, data);
    }

    public byte[] getMergedPDF(File templateDoc, Map<String, Object> model) throws IOException, XmlException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        InputStream byteInputStream = new ByteArrayInputStream(applyDataToTemplate(templateDoc, model));
        PdfConverter.getInstance().convert(new XWPFDocument(byteInputStream), output, PdfOptions.create());
        return output.toByteArray();
    }

    private byte[] applyDataToTemplate(File template, Map<String, Object> data) throws IOException, XmlException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        XWPFDocument temp = new XWPFDocument(new FileInputStream(template));
        CTBody body = temp.getDocument().getBody();
        String mergedStr = StrSubstitutor.replace(body.xmlText(), data);
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        body.set(CTBody.Factory.parse(mergedStr, optionsOuter));
        temp.write(out);
        return out.toByteArray();
    }

}
