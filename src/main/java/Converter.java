import java.io.File;
import java.util.Map;

public interface Converter {
    byte[] convert(File templateFile, Map<String, Object> data) throws Exception;
}
