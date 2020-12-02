import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import jwna.Common;
import jwna.DataSaver;
import jwna.Jwna;
import jwna.WeldingData;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

public class SaverTest {
  @BeforeClass
  public static void setUpClass() throws IOException, URISyntaxException {
    File file = new File(
     Jwna.class.getProtectionDomain().getCodeSource().getLocation().toURI());    
    Path rootPath = null;
    while(file != null) {
      Path path = Paths.get(file.getAbsolutePath(), "run");
      if(Files.exists(path, LinkOption.NOFOLLOW_LINKS)) {
        rootPath = path;
        break;
      }
      file = file.getParentFile();
    }    
    if(rootPath == null) {
      throw new IOException("Root folder[run] was not found.");
    }
    Common.initWithAppPath(rootPath.toString()); 
    Common.load();
  }
 
  
  @Test
  public void testWeldingDataDef() {
    final WeldingData wd = new WeldingData();
    
    assertFalse(wd.isLast());
    assertNull(wd.getVersion());
    assertNull(wd.getIpAddr());
    assertNull(wd.getData());
    assertNotNull(wd.getDataLines());
    assertNull(wd.getMsg());
    assertTrue(wd.getParams().isEmpty());
    assertTrue(wd.getRows() == 0);
    assertNull(wd.getPipeNumber());
    assertNull(wd.getPipeThickness());
    assertNull(wd.getStanName());
    assertNull(wd.getStanType());
    assertNull(wd.getStanNumber());
    assertNull(wd.getStanSize());
    assertNull(wd.getDate());
    assertNull(wd.getDstFilePath());
    assertNull(wd.getDstFileName());
  }
  
  
  @Test
  public void testWeldingData() {
    final WeldingData wd = new WeldingData();
    
    final boolean last = true;
    wd.setLast(last);
    assertTrue(last == wd.isLast());
    
    final String ipAddr = "127.0.0.1";
    wd.setIpAddr(ipAddr);
    assertTrue(ipAddr.equals(wd.getIpAddr()));
    
    final int version = 0x100;
    wd.setVersion(version);
    assertTrue(version == wd.getVersion());
    
    final int step = 3;
    wd.setStep(step);
    assertTrue(step == wd.getStep());
  }
  
  
  @Test
  public void testStans() {    
    Common.Stan stan;
    
    assertNotNull(Common.getStan("127.0.0.1"));
    
    stan = Common.getStan(Common.TEST_IP_ADDR_800_I);
    assertNotNull(stan);
    assertTrue(stan.getSize() == Common.STAN_800);
    assertTrue(stan.getType().equals(Common.STAN_I));
    
    stan = Common.getStan(Common.TEST_IP_ADDR_800_O);
    assertNotNull(stan);
    assertTrue(stan.getSize() == Common.STAN_800);
    assertTrue(stan.getType().equals(Common.STAN_O));
    
    stan = Common.getStan(Common.TEST_IP_ADDR_1000_I);
    assertNotNull(stan);
    assertTrue(stan.getSize() == Common.STAN_1000);
    assertTrue(stan.getType().equals(Common.STAN_I));
    
    stan = Common.getStan(Common.TEST_IP_ADDR_1000_O);
    assertNotNull(stan);
    assertTrue(stan.getSize() == Common.STAN_1000);
    assertTrue(stan.getType().equals(Common.STAN_O));
  }

  
  @Test
  public void testSaveVer1x00() throws IOException {
    testSaveVerX(0x100);
  }
  
  
  @Test
  public void testSaveVer1x10() throws IOException {
    testSaveVerX(0x110);
  }
  
  
  public void testSaveVerX(int version) throws IOException {
    final DataSaver saver = new DataSaver();

    WeldingData wd;
    File file;
    
    wd = generateWeldingData(version);
    wd.setIpAddr(Common.TEST_IP_ADDR_800_I);
    saver.save(wd);  
    
    wd = generateWeldingData(version);
    wd.setIpAddr(Common.TEST_IP_ADDR_800_O);
    saver.save(wd);
    
    wd = generateWeldingData(version);
    wd.setIpAddr(Common.TEST_IP_ADDR_1000_I);
    saver.save(wd);
    
    wd = generateWeldingData(version);
    wd.setIpAddr(Common.TEST_IP_ADDR_1000_O);
    saver.save(wd);
  }  
  
  private WeldingData generateWeldingData(int version) {
    final WeldingData wd = new WeldingData();
    wd.setVersion(version);
    wd.setPipeNumber("12345");
    wd.setPipeThickness("15");
    
    final int arcsCount = 4;
    final String[] arcFieldNames = {
    };   
    final String[] otherFieldNames = {      
    };
    final List<String> fieldNames = 
     new ArrayList<>(arcsCount*arcFieldNames.length + otherFieldNames.length);
    for(String name: arcFieldNames) {
      for(int i = 0; i < arcsCount; i++) {
        fieldNames.add(name + "_" + Integer.toString(i));
      }
    }
    fieldNames.addAll(Arrays.asList(otherFieldNames));
       
    class Offset {
      int value;
    }
    final Offset offset = new Offset();
    final int n = 1000;
    byte[] param = new byte[3*n];
    fieldNames.forEach(name -> {
      offset.value += 1;
      for(int i = 0; i < n; i++) {
        final int index = 3*i;
        final int v = i + offset.value;
        param[index] = (byte)(0xFF & v);
        param[index + 1] = (byte)(0xFF & (v >> 8));
        param[index + 2] = 0;
      }
      wd.setParam(name, param);
    });
    
    return wd;
  }
}
