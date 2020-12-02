package jwna;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
import java.util.Timer;
import java.util.TimerTask;
import java.util.logging.LogRecord;
import java.util.logging.Logger;

public class Common {
  public static final String APP_NAME = "jwna";
  private static final String CFG_FILE_NAME = APP_NAME + ".config";
  private static final String STANS_FILE_NAME = "stans.config";
  
  public static final int STAN_1000 = 1000;
  public static final int STAN_800  = 800;
  public static final String STAN_I = "ID";
  public static final String STAN_O = "OD";
  public static final String TEST_IP_ADDR_800_I = "TEST_800_I";
  public static final String TEST_IP_ADDR_800_O = "TEST_800_O";
  public static final String TEST_IP_ADDR_1000_I = "TEST_1000_I";
  public static final String TEST_IP_ADDR_1000_O = "TEST_1000_O";  
  
  private static final HashMap<String, Stan> sStans = new HashMap<>();
  
  private static String sAppPath;
  private static String sArchivePath;
  
  private static final Logger sLogger = Logger.getLogger(Jwna.class.getName());
  private static LogFileHandler sLogFileHandler;


  public static class Stan {
    private final String mIp;
    private final Integer mSize;
    private final Integer mNumber;
    private final String mType;
    

    private Stan(String ip, Integer size, Integer number, String type) {
      mIp = ip;
      mSize = size;
      mNumber = number;
      mType = type;
    }
    

    public String getIp() {
      return mIp;
    }
    

    public Integer getSize() {
      return mSize;
    }
    

    public Integer getNumber() {
      return mNumber;
    }
    

    public String getType() {
      return mType;
    }    
  }
  

  public static void initWithAppPath(String path) {
    sAppPath = path;
    sLogFileHandler = new LogFileHandler();
    sLogger.addHandler(sLogFileHandler);
  }
  

  public static void init() {
    initWithAppPath(new File("").getAbsolutePath());
  }
  

  public static void load() {
    final ConfigFile cfg = new ConfigFile();
    
    cfg.load(Paths.get(sAppPath, CFG_FILE_NAME).toFile());
    
    String tmpStr = cfg.read("ArchivePath", "");
    if(new File(tmpStr).isDirectory()) {
      sArchivePath = tmpStr;
    }
    else {
      System.out.println(String.format("Archive path was not found [%s].", 
       tmpStr));
      System.exit(0);
    }
    
    loadStans();
  }
  
  
  public static Stan getStan(String ip) {
    return sStans.get(ip);
  }
  
  
  private static void loadStans() {   
    final File stansFile = Paths.get(sAppPath, STANS_FILE_NAME).toFile();
    try (BufferedReader reader = 
            new BufferedReader(
            new InputStreamReader(
            new FileInputStream(stansFile), StandardCharsets.UTF_8))) {
      String line;

      while ((line = reader.readLine()) != null) {
        final String[] fields = line.split(",");  
        final Stan newStan = createStan(fields);
        if (newStan != null) {
          addStan(newStan);
        }
      }
    }
    catch(Exception ex) {          
    }
    addStan(new Stan("127.0.0.1",         STAN_800,  1, STAN_I));
    addStan(new Stan(TEST_IP_ADDR_800_I,  STAN_800,  1, STAN_I));
    addStan(new Stan(TEST_IP_ADDR_800_O,  STAN_800,  1, STAN_O));
    addStan(new Stan(TEST_IP_ADDR_1000_I, STAN_1000, 1, STAN_I));
    addStan(new Stan(TEST_IP_ADDR_1000_O, STAN_1000, 1, STAN_O));
  }    
  
  
  private static void addStan(Stan stan) {
    sStans.put(stan.getIp(), stan);
  }
  
  
  private static Stan createStan(String[] fields) {
    Stan stan = null;

    creation: {
      if (fields == null || fields.length != 4) {
        break creation;
      }

      final String ip = fields[0].trim();

      final String pipeSizeStr = fields[1].trim();
      int pipeSize = -1;
      try {
        pipeSize = Integer.parseInt(pipeSizeStr);
      }
      catch(Exception ex) {
        break creation;
      }
      switch(pipeSize) {
        case STAN_800:
        case STAN_1000:
          break;
        default:
          break creation;
      }

      final String stanNumStr = fields[2].trim();
      int stanNum = -1;
      try {
        stanNum = Integer.parseInt(stanNumStr);
      }
      catch(Exception ex) {              
      }
      if (stanNum < 1) {
        break creation;
      }

      final String stanType = fields[3].trim().toUpperCase();
      if (!stanType.equals(STAN_I)
          && !stanType.equals(STAN_O)) {
        break creation;
      }

      stan = new Stan(ip, pipeSize, stanNum, stanType);
    } 

    return stan;
  }
  

  public static String getAppPath() {
    return sAppPath;
  }
  

  public static String getArchivePath() {
    return sArchivePath;
  }
  

  public static Logger getLogger() {
    return sLogger;
  }
  

  private static class LogFileHandler extends java.util.logging.Handler {
    private BufferedOutputStream mOutputStream;
    private final Object mSync = new Object();
    private LocalDate mFileDate;
    private final Timer mUpdateTimer = new Timer();
    private final Charset mCharset = Charset.forName("UTF8");
    private final DateTimeFormatter mDateTimeFormatter =
     DateTimeFormatter.ISO_DATE_TIME;
    

    private LogFileHandler() {         
      recreateIfNeeded(LocalDate.now());
      mUpdateTimer.scheduleAtFixedRate(new TimerTask() {
        @Override
        public void run() {
          recreateIfNeeded(LocalDate.now());
        }}, 0, 10000);     
    }
    

    @Override
    public void close() throws SecurityException {          
      synchronized(mSync) {        
        final OutputStream stream = mOutputStream;
        mOutputStream = null;
        if(stream != null) {
          try {
            stream.close();
          }
          catch(Exception ex) {            
            handleException(ex);
          }
        }
      }
    }
    

    @Override
    public void flush() {      
      synchronized(mSync) {
        if(mOutputStream != null) {
          try {
            mOutputStream.flush();
          }
          catch(Exception ex) {            
            handleException(ex);
          }
        }
      }
    }
    

    @Override
    public void publish(LogRecord record) {
      final LocalDateTime dttm = new Date(record.getMillis())
       .toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();  
      final Throwable thrown = record.getThrown();
      final StringBuffer buff = new StringBuffer();
      buff.append(dttm.format(mDateTimeFormatter))   
          .append(" (")
          .append(record.getLevel().toString())
          .append(")\n")
          .append(record.getMessage())
          .append("\n");
      if(thrown != null) {
        for(StackTraceElement el: thrown.getStackTrace()) {
          buff.append(" ")
              .append(el.toString())
              .append("\n");
        }
      }
      buff.append("\n");
      final byte[] byteBuff = buff.toString().getBytes(mCharset);
      final LocalDate date = dttm.toLocalDate();      
      
      synchronized(mSync) {
        recreateIfNeeded(date);
        if(mOutputStream != null) {          
          try {
            mOutputStream.write(byteBuff);
          }
          catch(Exception ex) {  
            handleException(ex);
          }
        }
      }
    }
    

    private void recreateIfNeeded(LocalDate date) {            
      synchronized(mSync) {
        if(!(mFileDate != null
           && mFileDate.getMonthValue() == date.getMonthValue()
           && mFileDate.getYear() == date.getYear())) {                
          close();
          
          BufferedOutputStream stream = null;
          
          try {
            final Path filePath = getFilePath(date);          
            final Path folderPath = filePath.getParent();
            if(!Files.exists(folderPath)) {
              Files.createDirectories(folderPath);
            }
            final FileOutputStream fileStream =
             new FileOutputStream(filePath.toFile(), Files.exists(filePath));
            stream = new BufferedOutputStream(fileStream);
            mFileDate = date;
          }
          catch(Exception ex) {  
            handleException(ex);
          }
          
          mOutputStream = stream;
        }
      }
    }
    

    private Path getFilePath(LocalDate date) {
      final String fileName = String.format("%s (%04d-%02d).log", 
       APP_NAME, date.getYear(), date.getMonthValue());
      return Paths.get(sAppPath, "Log", fileName);
    }
    

    @SuppressWarnings("CallToPrintStackTrace")
    private void handleException(Exception ex) {
      ex.printStackTrace();
    }            
  }
}
