package jwna;

import java.time.LocalDateTime;
import java.util.Map;
import java.util.TreeMap;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WeldingData {
  private static final String[] EMPTY_DATA_LINES = new String[0];    
  private boolean mLast = false;
  private String mIpAddr;
  private Integer mVersion;
  private String mData;
  private String[] mDataLines = EMPTY_DATA_LINES;
  private byte[] mMsg;
  private final TreeMap<String, byte[]> mParams = new TreeMap<>();
  private HSSFWorkbook mWorkbook;
  private HSSFSheet[] mSheets;
  private int mRows = 0;
  private int mStep = 3;
  private String mPipeNumber;
  private String mPipeThickness;
  private String mStanName;
  private String mStanType;
  private Integer mStanNumber;
  private Integer mStanSize;
  private LocalDateTime mDate;
  private String mDstFilePath;
  private String mDstFileName;
  

  public boolean isLast() {
    return mLast;
  }
  

  public void setLast(boolean newLast) {
    mLast = newLast;
  }  
  

  public String getIpAddr() {
    return mIpAddr;
  }
  

  public void setIpAddr(String newAddr) {
    mIpAddr = newAddr;
  }
  

  public Integer getVersion() {
    return mVersion;
  }
  

  public void setVersion(Integer newVersion) {
    mVersion = newVersion;
  }
  

  public String getData() {
    return mData;
  }
  

  public void setData(String newData) {
    mData = newData;    
    if(mData == null) {
      mDataLines = EMPTY_DATA_LINES;
    }
    else {
      mDataLines = mData.replace("\015", "\012")
                        .replace("\012\012", "\012")
                        .split("\012");     
    }
  }
  

  public String[] getDataLines() {
    return mDataLines;
  }
  

  public byte[] getMsg() {
    return mMsg;
  }
  

  public void setMsg(byte[] newMsg) {
    mMsg = newMsg;
  }
  

  public Map<String, byte[]> getParams() {
    return mParams;
  }
  

  public byte[] getParam(String field) {
    return mParams.get(field);
  }

  
  public void setParam(String field, byte[] values) {
    mParams.put(field, values);  
  }
  

  public HSSFWorkbook getWorkbook() {
    return mWorkbook;
  }
  

  public void setWorkbook(HSSFWorkbook newWorkbook) {
    mWorkbook = newWorkbook;
  }
  

  public HSSFSheet[] getSheets() {
    return mSheets;
  }
  

  public void setSheets(HSSFSheet[] newSheets) {
    mSheets = newSheets;
  }
  

  public int getRows() {
    return mRows;
  }
  

  public void setRows(int newRows) {
    mRows = newRows;
  }
  
   

  public int getStep() {
    return mStep;
  }
  

  public void setStep(int newStep) {
    mStep = newStep;
  }
  

  public String getPipeNumber() {
    return mPipeNumber;
  }
  

  public void setPipeNumber(String newPipeNumber) {
    mPipeNumber = newPipeNumber;
  }
  

  public String getPipeThickness() {
    return mPipeThickness;
  }
  

  public void setPipeThickness(String newPipeThickness) {
    mPipeThickness = newPipeThickness;
  }
  

  public String getStanName() {
    return mStanName;
  }
  

  public void setStanName(String newStanName) {
    mStanName = newStanName;
  }
  

  public String getStanType() {
    return mStanType;
  }
  

  public void setStanType(String newStanType) {
    mStanType = newStanType;
  }
  

  public Integer getStanNumber() {
    return mStanNumber;
  }
  

  public void setStanNumber(Integer newStanNumber) {
    mStanNumber = newStanNumber;
  }
  

  public Integer getStanSize() {
    return mStanSize;
  }
  

  public void setStanSize(Integer newStanSize) {
    mStanSize = newStanSize;
  }
  

  public LocalDateTime getDate() {
    return mDate;
  }
  

  public void setDate(LocalDateTime newDate) {
    mDate = newDate;
  }

  
  public String getDstFilePath() {
    return mDstFilePath;
  }
  

  public void setDstFilePath(String filePath) {
    mDstFilePath = filePath;
  }
  

  public String getDstFileName() {
    return mDstFileName;
  }
  

  public void setDstFileName(String fileName) {
    mDstFileName = fileName;
  }
}
