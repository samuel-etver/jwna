package jwna;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;

public class ConfigFile {
  private final HashMap<String, String> mList = new HashMap<>();


  public void clear() {
    mList.clear();
  }


  public String read(String key, String defValue) {
    final String value = mList.get(key);
    return value == null ? defValue : value;
  }


  public boolean read(String key, boolean defValue) {
    final String valueStr = read(key, null);
    boolean value = defValue;
    
    if (valueStr.equalsIgnoreCase("true")) {
      value = true;
    }
    else if (valueStr.equalsIgnoreCase("false")) {
      value = false;
    }
    
    return value;
  }  

  
  public int read(String key, int defValue) {
    final String valueStr = read(key, null);
    int result = defValue;
    
    try {
      result = Integer.parseInt(valueStr);
    }
    catch(Exception exception) {
    }
    
    return result;
  }

  
  public boolean load(File file) {
    boolean done = false;
    
    try (BufferedReader reader =
            new BufferedReader(
            new InputStreamReader(
            new FileInputStream(file), StandardCharsets.UTF_8))) {
      String line;

      while ((line = reader.readLine()) != null) {
        final int pos = line.indexOf('=');
        if (pos > 0) {
            final String key = line.substring(0, pos);
            final String value = line.substring(pos + 1);
            mList.put(key, value);
        }
      }
      done = true;
    }
    catch(Exception exception) {
    }

    return done;
  }
}
