package jwna;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Row;

public class DataSaver {
  private final String mSrcFilePath; 
  private final Pattern mCellNamePattern =
   Pattern.compile("([A-Za-z]+)([0-9]+)");
  

  public DataSaver() {
    mSrcFilePath = Common.getAppPath();
  }
  

  public void save(WeldingData wd) throws IOException {
    final LocalDateTime startTime = LocalDateTime.now();
    wd.setDate(startTime);
    Integer version = wd.getVersion();
    switch(version) {
      case 0x100:
        save1x00(wd);
        break;
      case 0x110:
        save1x10(wd);
    }
  }


  private void save1x10(WeldingData wd) throws IOException {
    save1x00(wd);
  }
  

  private void save1x00(WeldingData wd) throws IOException {
    final Common.Stan stan = Common.getStan(wd.getIpAddr());
    if(stan == null) {
        return;
    }
    wd.setStanSize(stan.getSize());
    wd.setStanNumber(stan.getNumber());
    wd.setStanType(stan.getType());
    wd.setStanName(wd.getStanType() + Integer.toString(wd.getStanNumber()));

    String[] path_pair = getSrcFileName(wd);
    final String src_path = path_pair[0];
    final String src_fn   = path_pair[1];
    path_pair = getDstFileName(wd);
    final String dst_path = path_pair[0];
    final String dst_fn   = path_pair[1];
    if(!Files.isDirectory(Paths.get(dst_path)))
      Files.createDirectories(Paths.get(dst_path));
    Files.copy(Paths.get(src_path, src_fn), Paths.get(dst_path, dst_fn));
    wd.setDstFilePath(dst_path);
    wd.setDstFileName(dst_fn);

    openWb(wd, Paths.get(dst_path, dst_fn).toString());
    wd.setRows(getRowsMax(wd));

    final ParamConverter intToEx = (value, flags) -> {
      if(value == null)
        return 0.0;
      return value;
    };

    final ParamConverter vToEx = intToEx;
            
    final ParamConverter iToEx = intToEx;
            
    final ParamConverter valuesTicksToEx = (value, flags) -> {
      if(value != null)
        return (value*0x100 + flags)/1000.0;
      return 0.0;
    };            
            
    final ParamConverter vwmToEx = intToEx;
            
    final ParamConverter headPosToEx = (value, flags) -> {
      if(value != null) {
        int v = value.intValue();
        if((v & 0x8000) != 0)
          v -= 65536;
        return v/10.0;  
      }                  
      return intToEx.convert(value, flags);
    };

    final ParamConverter headPosDiffToEx = intToEx;
            
    final ParamConverter edgeSizeToEx = (value, flags) -> {
      if(value != null)
        return value/10.0;
      return 0.0;
    };
            
    final ParamConverter caretIdToEx = (value, flags) -> {
      if(value != null)
        return intToEx.convert(value/100.0, flags);
      return intToEx.convert(value, flags);
    };
            
    final ParamConverter wireIdToEx = (value, flags) -> {
      if(value != null)
        return value/100.0;
      return 0.0;
    };
            
    final ParamConverter autoVertToEx = (value, flags) -> {
      if(value != null)
        return value + 1.0;
      return 0.0;
    };
            
    final ParamConverter autoHorzToEx = (value, flags) -> {
      if(value != null)
        return value + 3;
      return 0.0;
    };
            
    final ParamConverter pipeDetectionToEx = (value, flags) -> {
      if(value != null)
        return value + 5;
      return 0.0;
    };

    final ParamConverter pipeSpeedToEx = intToEx;

    final ParamConverter wireVelocityToExp = intToEx;

    final ParamConverter currentFvToExp = intToEx;

    final ParamConverter downForceToEx = intToEx;

    final ParamConverter wireSpaceToEx = intToEx;

    final ParamConverter alphaToEx = intToEx;

    final ParamConverter weldingHeadHorzCenterToEx = intToEx;

    final ParamConverter fluxValveToEx = intToEx;

    final ParamConverter fluxAvailableToEx = intToEx;

    final ParamConverter fluxPressureToEx = intToEx;

    final ParamConverter pipeRotateLeftCmdToEx = intToEx;

    final ParamConverter pipeRotateRightCmdToEx = intToEx;

    final ParamConverter controlPositionToEx = intToEx;
    
    final FieldExportable export_field = (col, field, converter) -> {      
      exportField(wd, col, field, converter);
    };
    
    final FieldsExportable export_fields = (col, field_base, converter) -> {      
      exportFields(wd, col, genFields(wd, field_base), converter);
    };

    
    if(wd.getStanType().equals(Common.STAN_O)) {               
      exportField_Meter(wd, 1);
      export_fields.export(2, 2, iToEx);
      export_fields.export(6, 1, vToEx);
      export_fields.export(10, 10, vwmToEx);
      export_fields.export(14, 4, wireIdToEx);
      exportFields_IdMax(wd, 18, wireIdToEx);
      export_field.export(19, "5_0", pipeSpeedToEx);
      export_field.export(20, "6_0", caretIdToEx);
      export_field.export(21, "24_0", autoVertToEx);
      export_field.export(22, "25_0", autoHorzToEx);
      export_field.export(23, "26_0", pipeDetectionToEx);
      export_field.export(24, "11_0", edgeSizeToEx);
      export_field.export(25, "12_0", edgeSizeToEx);
      export_field.export(26, "13_0", edgeSizeToEx);
      export_field.export(27, "17_0", headPosToEx);
      export_field.export(28, "27_0", headPosToEx);
      exportField_ResultGeom(wd, 29, headPosDiffToEx);
      export_field.export(30, "19_0", headPosToEx);
      export_field.export(31, "41_0", controlPositionToEx);
      export_field.export(32, "37_0", pipeRotateLeftCmdToEx);
      export_field.export(33, "38_0", pipeRotateRightCmdToEx);
      exportHorzHeadPosition(wd, 34);
      export_fields.export(36, 3, wireVelocityToExp);
      export_fields.export(40, 18, currentFvToExp);
      export_fields.export(44, 22, iToEx);
      export_fields.export(48, 30, alphaToEx);
      export_field.export(52, "20_0", pipeSpeedToEx);
      export_field.export(53, "33_0", weldingHeadHorzCenterToEx);
      export_field.export(54, "34_0", fluxValveToEx);
      export_field.export(55, "35_0", fluxAvailableToEx);
      export_field.export(56, "36_0", fluxPressureToEx);
      export_field.export(57, "28_0", valuesTicksToEx);
      export_field.export(58, "49_0", iToEx);
      export_field.export(59, "50_0", iToEx);
      export_fields.export(60, 21, vToEx);
    }
    else {            
      exportField_Meter(wd, 1);
      export_fields.export(2, 2, iToEx);                
      export_fields.export(5, 1, vToEx);                
      export_fields.export(8, 10, vwmToEx);                
      export_fields.export(11, 4, wireIdToEx);                
      exportFields_IdMax(wd, 14, wireIdToEx);               
      export_field.export(15, "5_0",   pipeSpeedToEx);
      export_field.export(16, "6_0",   caretIdToEx);                
      export_field.export(17, "24_0",  autoVertToEx);
      export_field.export(18, "19_0", headPosToEx);                
      export_field.export(19, "14_0", downForceToEx);
      export_field.export(20, "15_0", downForceToEx);
      export_field.export(21, "16_0", downForceToEx);
      export_fields.export(22, 3, wireVelocityToExp);
      export_fields.export(25, 18, currentFvToExp);
      export_fields.export(28, 22, iToEx);
      export_field.export(31, "7_0", wireSpaceToEx);
      export_fields.export(32, 30, alphaToEx);
      export_field.export(35, "20_0", pipeSpeedToEx);        
      export_field.export(36, "28_0", valuesTicksToEx);
      export_field.export(37, "49_0", iToEx);
      export_field.export(38, "50_0", iToEx);
      export_fields.export(39, 21, vToEx);
    }

    exportManualControl(wd);
    exportDateTime(wd);
    exportStanName(wd);
    exportIx(wd);
    exportUx(wd);
    exportWeldingSpeed(wd);
    exportPersonnelNo(wd);
    exportPipeNumber(wd);
    exportPipeThickness(wd);
    exportPipeDiameter(wd);
    exportRegOptions(wd);              
    exportWeldParams(wd); 
    exportCurrentFvChange(wd, genFields(wd, 31), wireVelocityToExp);
    exportSSxFvChange(wd, genFields(wd, 32), wireVelocityToExp);
    exportPipeVelocityFVChange(wd, "48_0", pipeSpeedToEx);
    exportWelderState(wd, 1);
    exportWelderState(wd, 2);
    exportWelderState(wd, 3);
    exportWelderState(wd, 4);
    exportWireDiameter(wd, 1);
    exportWireDiameter(wd, 2);
    exportWireDiameter(wd, 3);
    
    closeWb(wd);
  }


  private void openWb(WeldingData wd, String filename) throws IOException {
    try (FileInputStream stream = new FileInputStream(filename)) {
      final HSSFWorkbook workbook = new HSSFWorkbook(stream);
      wd.setWorkbook(workbook);
      final String[] sheetNames = {
        "Data",
        "Setup",
        "Reg",
        "Changes",
        "Params"};
      final HSSFSheet[] sheets = new HSSFSheet[sheetNames.length];
      for(int i = 0; i < sheets.length; i++) {
        sheets[i] = workbook.getSheet(sheetNames[i]); 
      }
      wd.setSheets(sheets);
    } 
  }


  private void closeWb(WeldingData wd) throws IOException {
    try(HSSFWorkbook workbook = wd.getWorkbook()) {
      wd.setWorkbook(null);
      final File file =
       Paths.get(wd.getDstFilePath(), wd.getDstFileName()).toFile();
      workbook.write(file);
    }
  }


  private HSSFSheet getSheet(WeldingData wd, int index) {
    return wd.getSheets()[index];
  }
 

  private void setCell(WeldingData wd, int sheet_index, String cell, Object value) {
    if(value != null && value instanceof String) {
      conv: {
        try {
          value = Integer.parseInt((String)value);
          break conv;
        }
        catch(Exception ex) {          
        }
        try {
          value = Double.parseDouble((String)value);
          break conv;
        }
        catch(Exception ex) {          
        }
        try {
          value = Double.parseDouble(((String)value).replace(',', '.'));
          break conv;
        }
        catch(Exception ex) {          
        }        
      }
    }
    
    final CellReference cr = new CellReference(cell);
    Row rowObj = getSheet(wd, sheet_index).getRow(cr.getRow());
    if(rowObj == null) {
      rowObj = getSheet(wd, sheet_index).createRow(cr.getRow());
    }
    Cell cellObj = rowObj.getCell(cr.getCol()); 
    if(cellObj == null) {
      cellObj = rowObj.createCell(cr.getCol());
    }
    if(value == null) {
      cellObj.setCellValue("");
    }
    else if(value instanceof Integer) {
      cellObj.setCellValue((Integer)value);
    }
    else if(value instanceof Float) {
      cellObj.setCellValue((Float)value);      
    }
    else if(value instanceof Double) {
      cellObj.setCellValue((Double)value);
    }
    else if(value instanceof String) {
      cellObj.setCellValue((String)value);
    }
    else {
      cellObj.setCellValue(value.toString());
    }
  }              
    

  private String dateToStr(WeldingData wd) {
    final LocalDateTime dt = wd.getDate();
    return String.format("%d-%02d-%02d__%02d-%02d-%02d",
      dt.getYear(), dt.getMonthValue(), dt.getDayOfMonth(), 
      dt.getHour(), dt.getMinute(), dt.getSecond());
  }


  private String getDataValue(WeldingData wd, String key) {
    final int keyLen = key.length();    
    final String[] lines = wd.getDataLines();
    for(String line: lines) {
      if(line.length() >= keyLen) {
        if(line.substring(0, keyLen).equals(key)) {
          return line.substring(keyLen);
        }
      }
    }
    return null;
  }              
    

  private int getFieldSize(WeldingData wd, String field) {
    final Map<String, byte[]> params = wd.getParams();
    if(params.containsKey(field)) {
        return params.get(field).length/3/wd.getStep();
    }
    return 0;
  }
  

  private int getRowsMax(WeldingData wd) {
    final Map<String, byte[]> params = wd.getParams();
    int max = 0;
    for(Map.Entry<String, byte[]> entry: params.entrySet()) {
      final int v = getFieldSize(wd, entry.getKey());
      if(v > max) 
        max = v;
    }
    return max;
  }


  private String[] getSrcFileName(WeldingData wd) {   
    String src_fn;
    if(Common.STAN_I.equals(wd.getStanType())) {
      src_fn = 
       wd.getStanSize() == Common.STAN_1000
       ? "template-i.xls"
       : "template-i-800.xls";
    }        
    else {
      src_fn =
       wd.getStanSize() == Common.STAN_1000
       ? "template-o.xls"
       : "template-o-800.xls";
    }
    return new String[] {mSrcFilePath, src_fn};
  }      


  private String[] getDstFileName(WeldingData wd) {
    final String dst_path = Paths.get(Common.getArchivePath(),
     wd.getStanSize().toString(),
     wd.getStanType() + wd.getStanNumber().toString(), 
     Integer.toString(wd.getDate().getYear()),
     Integer.toString(wd.getDate().getMonthValue())).toString();
    String dst_fn;

    if(wd.getVersion() == 0x100) {
      dst_fn = dateToStr(wd) + ".xls";
    }
    else {
      String dst_fn_base = wd.getStanSize().toString() + "_" + 
        wd.getStanType() + "-" + wd.getStanNumber().toString() + "____" +
        Integer.toString(wd.getDate().getYear() % 100) + 
        "00" + wd.getPipeNumber()+ "_" + wd.getPipeThickness() + "_data";
      dst_fn = dst_fn_base + ".xls";
      if(Files.isRegularFile(Paths.get(dst_path, dst_fn))) {
        int dst_fn_index = 0;
        dst_fn_base  += "_";
        boolean found = true;
        while(found) {
          dst_fn_index += 1;
          dst_fn = dst_fn_base + Integer.toString(dst_fn_index) + ".xls";
          found = Files.isRegularFile(Paths.get(dst_path, dst_fn));
        }
      }                
    }
    return new String[] {dst_path, dst_fn};
  }              


  private void exportManualControl(WeldingData wd) {
    String v = getDataValue(wd, "ManualControl=");
    if(v != null && v.length() > 0) {
      switch(v.substring(0, 1).toUpperCase()) {
        case "F": 
          v = "Auto";
          break;
        case "T":
          v = "Hand";
      }
    }
    setCell(wd, 1, "O4", v);
  }      


  private void exportWeldingSpeed(WeldingData wd) {
    final String v = getDataValue(wd, "WeldingSpeedFv=");
    setCell(wd, 1, "B29", v);
  }


  private void exportPipeNumber(WeldingData wd) {
    final String v = getDataValue(wd, "PipeNumber=");
    setCell(wd, 1, "C3", v);
  }
  

  private void exportPersonnelNo(WeldingData wd) {
    final String v = getDataValue(wd, "PersonnelNo=");
    setCell(wd, 1, "O3", v);
  }
  

  private void exportPipeThickness(WeldingData wd) {
    final String v = getDataValue(wd, "PipeThickness=");
    setCell(wd, 1, "O21", v);
  }      


  private void exportPipeDiameter(WeldingData wd) {
    final String v = getDataValue(wd, "PipeDiameter=");
    setCell(wd, 1, "O20", v);
  }
        

  private void exportDateTime(WeldingData wd) {
    final LocalDateTime dttm = wd.getDate();
    final String dt = dttm.format(DateTimeFormatter.ofPattern("d.MM.yyyy"));
    final String tm = dttm.format(DateTimeFormatter.ofPattern("h:mm a"));
    setCell(wd, 1, "I2", dt);
    setCell(wd, 1, "L2", tm);
  }


  private void exportPipeMeterFK(WeldingData wd) {
    final String v = getDataValue(wd, "PipeMeterFK=");
    setCell(wd, 2, "B28", v);
  }


  private void exportStanName(WeldingData wd) {
    setCell(wd, 1, "F2", wd.getStanName());
  }


  private void exportStProvarActive(WeldingData wd) {
    final String v = getDataValue(wd, "StProvarActive=");
    setCell(wd, 4, "B20", v);
  }      


  private void exportProvarActive(WeldingData wd) {
    final String v = getDataValue(wd, "ProvarActive=");
    setCell(wd, 4, "B31", v);
  }


  private void exportWireDiameter(WeldingData wd, int arc) {
    if(1 <= arc && arc <= 3) {
      final String[] cells = {"C9", "F9", "I9"};
      final String v = 
       getDataValue(wd, "WireDiameter" + Integer.toString(arc) + '=');
      setCell(wd, 1, cells[arc - 1], v);
    }
  }


  private void exportWelderState(WeldingData wd, int arc) {
    final String v = 
     getDataValue(wd, "WelderState" + Integer.toString(arc) + '=');
    if(1 <= arc && arc <= 4 && v != null) {
      final String[] cells = {"O7", "O8", "O9", "O10"};
      final TreeMap<String, String> d = new TreeMap<>();
      if(arc == 1) {
        d.put("0", "Source:x1;Master:Off;Slave:Off");
        d.put("1", "Source:x2;Master:Off;Slave:Off");
        d.put("2", "Source:x1;Master:On;Slave:Off");
        d.put("3", "Source:x2;Master:On;Slave:Off");
        d.put("4", "Source:x1;Master:Off;Slave:On");
        d.put("5", "Source:x2;Master:Off;Slave:On");
        d.put("6", "Source:x1;Master:On;Slave:On");
        d.put("7", "Source:x2;Master:On;Slave:On");
      }        
      else {
        d.put("0", "MED");
        d.put("1", "MAX");
      }
      setCell(wd, 1, cells[arc - 1], d.get(v));
    }
  }
        

  private void exportIx(WeldingData wd) {
    final String[] cols = {"C", "F", "I", "L"};
    for(int i = 1; i < 5; i++) {
      final String v = getDataValue(wd, "Ix" + Integer.toString(i) + "=");
      if(v != null) {
        final String col = cols[i - 1];
        setCell(wd, 1, col +  "7", "1");
        setCell(wd, 1, col + "13", v);
      }
    }
  }


  private void exportUx(WeldingData wd) {
    final String[] cols = {"C", "F", "I", "L"}; 
    for(int i = 1; i < 5; i++) {
      final String v = getDataValue(wd, "Ux" + Integer.toString(i) + "=");
      if(v != null) 
        setCell(wd, 1, cols[i - 1] + "14", v + " B");
    }
  }              


  private void exportGroup(WeldingData wd, int sheetIndex, String[] cols,
     String group, int startRow) {
    for(int j = 1; j < 5; j++) {
      final String key = group + "-" + Integer.toString(j) + "-";
      final int keyLen = key.length();
      final String[] dataLines = wd.getDataLines();
      int row = startRow;
      String v;      
      for (String line : dataLines) {
        if(line.length() >= keyLen && line.substring(0, keyLen).equals(key)) {
          final String cell = cols[j] + Integer.toString(row++);
          final int index = line.indexOf('=');
          if(index >= 0)
            v = line.substring(index + 1);
          else
            v = "";
          setCell(wd, sheetIndex, cell, v);
        }
      }
    }
  }


  private void exportOptions(WeldingData wd, String[] cols,
     String group, int startRow) {
    exportGroup(wd, 2, cols, group, startRow);
  }


  private void exportParamsRow(WeldingData wd, String[] cols,
     String group, int startRow) {
    exportGroup(wd, 4, cols, group, startRow);
  }
                 

  private void exportRegOptions(WeldingData wd) {
    final String[] cols = {null, "B", "C", "D", "E"};
    exportOptions(wd, cols, "Reg", 3);
    exportPipeMeterFK(wd);
    exportOptions(wd, cols, "BefStartWireFeed", 30);
    exportOptions(wd, cols, "WeldingUp", 37);
    exportOptions(wd, cols, "WeldingDown", 51);
  }  


  private void exportWeldParams(WeldingData wd) {
    final String[] cols = {null, "C", "D", "E", "F"};
    exportParamsRow(wd, cols, "WeldPar", 6);
    exportParamsRow(wd, cols, "StartPar", 13);
    exportParamsRow(wd, cols, "StProvarPar", 22);
    exportParamsRow(wd, cols, "DnProvarPar", 33);
    exportParamsRow(wd, cols, "AutoPar", 44);
    exportParamsRow(wd, cols, "DownPar", 51);
    exportParamsRow(wd, cols, "CaretPar", 66);
    exportStProvarActive(wd);
    exportProvarActive(wd);
  }              
  

  private int colNameToIndex(String name) {
    int index = 0;
    final int n = name.length();
    final int period = 'Z' - 'A' + 1;
    for(int i = 0; i < n; i++) {
      index = (name.charAt(i) - 'A' + 1) + index*period;
    }
    return index;
  }
    

  private String createExcelColumnName(WeldingData wd, int col) {
    col -= 1;        
    final int n = 'Z' - 'A' + 1;
    if(col >= n) {
      if(col >= n*2) {
        if(col >= n*3) {
          return "C" + (char)('A' + col - 3*n);
        }
        return "B" + (char)('A' + col - 2*n);
      }
      return "A" + (char)('A' + col - n);
    }            
    return  "" + (char)('A' + col);
  }              


  private void setRange(WeldingData wd, int sheet_index, String range_start, 
     String range_end, Object values) {
    final Matcher matcher = mCellNamePattern.matcher(range_start);
    String colName;
    String rowName;
    if(matcher.find()) {
      colName = matcher.group(1);
      rowName = matcher.group(2);
    }
    else {
      colName = null;
      rowName = null;
    }
    final int rowStart = Integer.parseInt(rowName);
    final int colStart = colNameToIndex(colName);
    
    matcher.reset(range_end);
    if(matcher.find()) {
      colName = matcher.group(1);
      rowName = matcher.group(2);
    } 
    else {
      colName = null;
      rowName = null;
    }
    final int rowEnd = Integer.parseInt(rowName);
    final int colEnd = colNameToIndex(colName);
    
    Object v;
        
    for(int col = colStart; col <= colEnd; col++) {
      colName = createExcelColumnName(wd, col);
      final int index_c = col - colStart;

      int index_r = 0;
      for(int row = rowStart; row <= rowEnd; row++) {
        if(values == null) {
          v = null;
        }
        else if(!values.getClass().isArray()) {
          v = values;
        }
        else {
          Object col_values = ((Object[])values)[index_r];
          if(col_values == null) {
            v = null;
          }
          else if(!col_values.getClass().isArray()) {
            v = col_values;
          }
          else {
            v = ((Object[])col_values)[index_c];
          }
        }
        setCell(wd, sheet_index, colName + Integer.toString(row), v);                    
        index_r += 1;
      }
    }   
  }                      
    

  private void exportField_Meter(WeldingData wd, int col) {
    LinkedList<Double> values = new LinkedList<>();
    final String colName = createExcelColumnName(wd, col);
    final int n = wd.getRows()*wd.getStep();
    final int step = wd.getStep();
    for(int i = 0; i < n; i += step) {
      values.add((double)(i*2));
    }    
    setRange(wd, 0, colName + "2", colName + Integer.toString(wd.getRows() + 1),
     values.toArray());
  }              
            

  private String[] genFields(WeldingData wd, int prm) {
    LinkedList<String> fields = new LinkedList<>();
    final int n = wd.getStanType().equals(Common.STAN_O) ? 5 : 4;
    for(int i = 1; i < n; i++)
      fields.add(Integer.toString(prm) + "_" + Integer.toString(i));
    return fields.toArray(new String[n - 1]);
  }


  private void exportHorzHeadPosition(WeldingData wd, int col) {
    final String vStr = getDataValue(wd, "HorzHeadPositionRange=");
    Double v = null;
    conv: {
      try {
        v = Double.parseDouble(vStr);      
        break conv;
      }
      catch(Exception ex) {
      }
      try {
        v = Double.parseDouble(vStr.replace(",", "."));
        break conv;
      }
      catch(Exception ex) {        
      }
    }
    
    if(v != null) {
      final int n = wd.getRows();
      Double[] values = new Double[n];
      Arrays.fill(values, v);
      String colName = createExcelColumnName(wd, col);
      setRange(wd, 0, colName + "2", colName + Integer.toString(n + 1), values);
      Arrays.fill(values, -v);
      colName = createExcelColumnName(wd, col + 1);
      setRange(wd, 0, colName + "2", colName + Integer.toString(n + 1), values);
    }
  } 
  

  private static int getParamArg0(byte[] param, int index) {
    return (((int)param[index])     & 0xFF) +
           (((int)param[index + 1]) & 0xFF)*0x100; 
  }
  

  private static int getParamArg1(byte[] param, int index) {
    return ((int)param[index + 2]) & 0xFF;
  }  
    

  private void exportField(WeldingData wd, int col, String field,
     ParamConverter converter) {    
    final String colName = createExcelColumnName(wd, col);
    final int step = wd.getStep();
    final int n = getFieldSize(wd, field)*step;
    final int max = wd.getRows()*step;
    final Double[] values = new Double[max];
    final byte[] param = wd.getParam(field);
    int index = 0;
    
    int i;
    int j;
    for(i = 0; i < n; i += step) {
      j = 3*i;
      values[index++] = converter.convert(
       (double)getParamArg0(param, j), getParamArg1(param, j));
    }
    for(; i < max; i += step) {
      values[index++] = converter.convert(null, null);
    }

    setRange(wd, 0, colName + "2", colName + Integer.toString(wd.getRows() + 1),
     values);
  }
  
    

  private void exportFields(WeldingData wd, int col, String[] fields, 
     ParamConverter converter) {
    for(String field: fields) {
      exportField(wd, col++, field, converter);
    }
  }
  

  private void exportField_ResultGeom(WeldingData wd, int col, 
     ParamConverter converter) {
    final List<Double> values = new ArrayList<>();
    final String colName = createExcelColumnName(wd, col);

    final String field1 = "17_0";
    final String field2 = "27_0";
    
    final int step = wd.getStep();

    final int n1 = getFieldSize(wd, field1)*step;
    final int n2 = getFieldSize(wd, field2)*step;
    
    final byte[] param1 = wd.getParam(field1);
    final byte[] param2 = wd.getParam(field2);
    
    int v;
    double v1;
    double v2;
    int index;

    for(int i = 0; i < wd.getRows()*step; i += step) {
      index = 3*i;
      if(i < n1) {
        v = getParamArg0(param1, index);
        if((v & 0x8000) != 0)
            v -= 65536;
        v1 = v/10.0;
      }            
      else {
        v1 = 0.0;
      }
      if(i < n2) {
        v = getParamArg0(param2, index);
        if((v & 0x8000) != 0)
          v -= 65536;
        v2 = v/10.0;
      }            
      else {
        v2 = 0.0;
      }
      values.add(converter.convert(v1 - v2, 0));
    }

    setRange(wd, 0, colName + "2", colName + Integer.toString(wd.getRows() + 1),
     values.toArray());
  }


  private void exportFields_IdMax(WeldingData wd, int col, 
     ParamConverter converter) {
    final String colName = createExcelColumnName(wd, col);

    final String txt = getDataValue(wd, "WireIdMax=");
    Double v;
    conv: {
      try {
        v = Double.parseDouble(txt);
        break conv;
      }
      catch(Exception ex) {        
      }
      try {
        v = Double.parseDouble(txt.replace(',', '.'));
        break conv;
      }
      catch(Exception ex) {        
      }
      v = 0.0;
    }            

    v = converter.convert(v, 0);

    setRange(wd, 0, colName + "2", 
     colName + Integer.toString(wd.getRows() + 1), v);
  } 
  

  private void exportCurrentFvChange(WeldingData wd, String[] fields, 
     ParamConverter converter) {
    final int step = wd.getStep();
    final List<Double> values = new ArrayList<>();
    
    for(String field: fields) {
      int index = field.indexOf('_');
      if( index < 0)
        continue;
      final String arc = field.substring(index+1);
      String col1;
      String col2;
      switch(arc) {
        case "1":
          col1 = "B";
          col2 = "C";
          break;
        case "2":
          col1 = "D";
          col2 = "E";
          break;
        case "3":
          col1 = "F";
          col2 = "G";
          break;
        case "4":
          col1 = "H";
          col2 = "I";
          break;
        default:
          continue;
      }

      values.clear();
      final int n = getFieldSize(wd, field)*step;
      final byte[] param = wd.getParam(field);

      for(int i = 0; i < n; i += step) {
        index = 3*i;
        values.add(converter.convert(
         (double)getParamArg0(param, index), getParamArg1(param, index)));
      }            

      Double Fv = 0.0;
      if(n > 0) 
        Fv = values.get(0);

      int row = 6;
      setCell(wd, 3, col2 + Integer.toString(row), Fv);
      int i = 0;
      for(Double v: values) {
        if(Math.abs(v - Fv) > 0.001) {
          row++;
          setCell(wd, 3, col1 + Integer.toString(row), i);
          setCell(wd, 3, col2 + Integer.toString(row), v);
          Fv = v;
        }
        i += 6;
      }
    }
  }
  

  private void exportSSxFvChange(WeldingData wd, String[] fields,
     ParamConverter converter) {
    final int step = wd.getStep();
    final List<Double> values = new ArrayList<>();
    
    for(String field: fields) {
      int index = field.indexOf('_');
      if(index < 0)
          continue;
      final String arc = field.substring(index+1);
      String col1;
      String col2;
      switch(arc) {
        case "1":
          col1 = "K";
          col2 = "L";
          break;
        case "2":
          col1 = "M";
          col2 = "N";
          break;
        case "3":
          col1 = "O";
          col2 = "P";
          break;
        case "4":
          col1 = "Q";
          col2 = "R";
          break;
        default:
          continue;
      }

      values.clear();
      final int n = getFieldSize(wd, field)*step;
      final byte[] param = wd.getParam(field);

      for(int i = 0; i < n; i += step) {
        index = 3*i;
        values.add(converter.convert(
         (double)getParamArg0(param, index), getParamArg1(param, index)));
      }            

      Double Fv = 0.0;
      if(n > 0)
        Fv = values.get(0);

      int row = 6;
      setCell(wd, 3, col2 + Integer.toString(row), Fv);
      int i = 0;
      for(Double v: values) {
        if(Math.abs(v - Fv) > 0.001) {
          row++; 
          setCell(wd, 3, col1 + Integer.toString(row), i);
          setCell(wd, 3, col2 + Integer.toString(row), v);
          Fv = v;
        }            
        i += 6;
      }    
    }
  }              


  private void exportPipeVelocityFVChange(WeldingData wd, String field, 
     ParamConverter converter) {
    int row = 6;
    final List<Double> values = new ArrayList<>();
    final String col1 = "T";
    final String col2 = "U";
    Double Fv = 0.0;
    final int step = wd.getStep();
    final int n = getFieldSize(wd, field)*step;
    final byte[] param = wd.getParam(field); 
    int index;

    for(int i = 0; i < n; i+= step) {
      index = 3*i;
      values.add(converter.convert(
       (double)getParamArg0(param, index), getParamArg1(param, index)));
    }            

    if(n > 0)
      Fv = values.get(0);

    setCell(wd, 3, col2 + Integer.toString(row), Fv);
    int i = 0;
    for(Double v: values) {
      if(Math.abs(v - Fv) > 0.001) {
        row++; 
        setCell(wd, 3, col2 + Integer.toString(row), v);
        setCell(wd, 3, col1 + Integer.toString(row), i);
        Fv = v;
      }            
      i += 6;
    }            
  }                  
  

  private interface ParamConverter {  
    Double convert(Double value, Integer flags);
  }
  

  private interface FieldExportable {
    void export(int col, String field, ParamConverter converter);
  }

  
  private interface FieldsExportable {
    void export(int col, int field_base, ParamConverter converter);
  }
}
