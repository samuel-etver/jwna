package jwna;

import java.util.logging.Level;
import java.util.logging.Logger;

public class Jwna {
  public static void main(String[] args) {
    final Server server = Server.getInstance();
    try {
      server.init();
      server.startup();
    }
    catch(Exception ex) {
      final Logger log = Common.getLogger();
      if(log != null) {
        log.log(Level.SEVERE, "Startup server failure", ex);
      }
      System.exit(1);      
    }
    server.run();
  }  
}
