import java.io.*;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;

public class App {
    /**
     * Mail merge template based on Mail merge field
     * Ex:
     * Hello, <<name>>
     */
    private static final File MAIL_MERGE_FIELD_TEMPLATE = new File("mail-merge-field-template.docx");

    /**
     * Mail merge template based on Variable replacement
     * Ex:
     * Hello, ${name}
     */
    private static final File MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE = new File("mail-merge-variable-replacement-template.docx");

    public static void main(String[] args) throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("name", "Bhuwan");
        data.put("age", 26);
        data.put("email", "bhuwansharma.1996@gmail.com");
        apachePOIConversion(data);
        apachePOIJodConversion(data);
        asposeWordConversion(data);
        pandocConversion(data);
    }

    private static void apachePOIConversion(Map<String, Object> data) throws Exception {
        File outputApachePoi = new File("output-apache-poi.pdf");
        Converter poiConverter = new ApachePoiConverter();
        Files.write(outputApachePoi.toPath(), poiConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, data));
    }

    private static void apachePOIJodConversion(Map<String, Object> data) throws Exception {
        File outputJODWord = new File("output-jod-word.pdf");
        Converter jodConverter = new ApachePoiJodConverter();
        Files.write(outputJODWord.toPath(), jodConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, data));
    }

    private static void asposeWordConversion(Map<String, Object> data) throws Exception {
        File outputAsposeWord = new File("output-aspose-word.pdf");
        Converter asposeConverter = new AsposeWordConverter();
        Files.write(outputAsposeWord.toPath(), asposeConverter.convert(MAIL_MERGE_FIELD_TEMPLATE, data));
    }

    private static void pandocConversion(Map<String, Object> data) throws Exception {
        File outputPandoc = new File("output-pandoc.pdf");
        Converter pandocConverter = new PandocConverter();
        Files.write(outputPandoc.toPath(), pandocConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, data));
    }

}
