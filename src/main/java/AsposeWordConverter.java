import com.aspose.words.Document;
import com.aspose.words.MailMerge;
import com.aspose.words.SaveFormat;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Map;

public class AsposeWordConverter implements Converter {
    @Override
    public byte[] convert(File templateFile, Map<String, Object> data) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Document document = new Document(templateFile.getPath());
        int size = data.size();
        String[] keys = new String[size];
        String[] values = new String[size];
        int idx = 0;
        for (Map.Entry<String, Object> entry : data.entrySet()) {
            keys[idx] = entry.getKey();
            values[idx] = entry.getValue().toString();
            idx = idx + 1;
        }
        MailMerge merge = document.getMailMerge();
        merge.getFieldNames();
        merge.execute(keys, values);
        document.save(out, SaveFormat.PDF);
        return out.toByteArray();
    }
}
