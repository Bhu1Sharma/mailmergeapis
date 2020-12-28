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


    private static final Map<String, Object> DUMMY_DATE = new HashMap<>();

    static {
        DUMMY_DATE.put("name", "Bhuwan");
        DUMMY_DATE.put("age", 26);
        DUMMY_DATE.put("email", "bhuwansharma.1996@gmail.com");
    }

    public static void main(String[] args) throws Exception {
        /**
         *  apachePOIConversion() uses Apache POI Library for variable replacement
         *  For conversion from DOCX to PDF it uses opensagres PdfConverter
         */
        apachePOIConversion();

        /**
         * apachePOIJodConversion uses Apache POI Library for variable replacement
         * For conversion from DOCX to PDF it uses JODConverter as middleware and LibreOffice as engine
         */
        apachePOIJodConversion();

        /**
         * asposeWordConversion() uses aspose.word API to do mail merge on Merge field in DOCX and then convert it to PDF.
         */
        asposeWordConversion();

        /**
         * pandocConversion() uses Apache POI Library for variable replacement
         * For Conversion from DOCX to PDF it uses pandoc API
         */
        pandocConversion();
    }

    private static void apachePOIConversion() throws Exception {
        File outputApachePoi = new File("output-apache-poi.pdf");
        Converter poiConverter = new ApachePoiConverter();
        Files.write(outputApachePoi.toPath(), poiConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, DUMMY_DATE));
    }

    private static void apachePOIJodConversion() throws Exception {
        File outputJODWord = new File("output-jod-word.pdf");
        Converter jodConverter = new ApachePoiJodConverter();
        Files.write(outputJODWord.toPath(), jodConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, DUMMY_DATE));
    }

    private static void asposeWordConversion() throws Exception {
        File outputAsposeWord = new File("output-aspose-word.pdf");
        Converter asposeConverter = new AsposeWordConverter();
        Files.write(outputAsposeWord.toPath(), asposeConverter.convert(MAIL_MERGE_FIELD_TEMPLATE, DUMMY_DATE));
    }

    private static void pandocConversion() throws Exception {
        File outputPandoc = new File("output-pandoc.pdf");
        Converter pandocConverter = new PandocConverter();
        Files.write(outputPandoc.toPath(), pandocConverter.convert(MAIL_MERGE_VARIABLE_REPLACEMENT_TEMPLATE, DUMMY_DATE));
    }

}
