package example.office.utils;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

public class CharsetUtils {
    public static Charset detectCharset(byte[] buffer) {
        if(buffer[0] == (byte)0xEF && buffer[1] == (byte)0xBB && buffer[2] == (byte)0xBF) {
            return StandardCharsets.UTF_8;
        }

        if(buffer[0] == (byte)0xFF && buffer[1] == (byte)0xFE) {
            return StandardCharsets.UTF_16;
        }

        if(buffer[0] == (byte)0xFE && buffer[1] == (byte)0xFF) {
            return StandardCharsets.UTF_16BE;
        }

        // fallback to system default charset
        return Charset.defaultCharset();
    }
}
