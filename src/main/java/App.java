import java.io.*;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;

public class App {

    public static void main(String[] args) throws Exception {
        File inputApachePoi = new File("input-apachepoi.docx");
        File inputAsposeWord = new File("input-aspose-word.docx");
        File outputApachePoi = new File("output-apache-poi.pdf");
        File outputAsposeWord = new File("output-aspose-word.pdf");
        Map<String, Object> data = new HashMap<>();
        data.put("name", "Bhuwan");
        data.put("age", 26);
        data.put("email", "bhuwansharma.1996@gmail.com");
        Converter poiConverter = new ApachePoiConverter();
        Converter asposeConverter = new AsposeWordConverter();
        Files.write(outputApachePoi.toPath(), poiConverter.convert(inputApachePoi, data));
        Files.write(outputAsposeWord.toPath(), asposeConverter.convert(inputAsposeWord, data));
    }

}
