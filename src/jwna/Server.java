package jwna;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.IOException;
import java.net.ServerSocket;
import java.net.Socket;
import java.nio.charset.Charset;
import java.util.Arrays;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.logging.Level;
import java.util.logging.Logger;


public final class Server {
  private static final int SOCK_PORT           = 10000;
  private static final int SOCK_TIMEOUT        = 5000;
        
  private static final int PACKET_FAST_KEY     = 0x01230123;
  private static final int PACKET_KEY_LENGTH   = 10;
  private static final byte[] PACKET_KEY       = new byte[PACKET_KEY_LENGTH];
  private static final byte[] PACKET_SUCCEEDED = {(byte)0xFF, (byte)0xFF};
  
  private static final int PACKET_ID_BEGIN_REQUEST = 1;
  private static final int PACKET_ID_END_REQUEST   = 2;
  private static final int PACKET_ID_WRITE_DATA    = 3;
  private static final int PACKET_ID_WRITE_MSG     = 4;
  private static final int PACKET_ID_WRITE_PARAM   = 5;
  
  static {
    final String key = "WELDING";
    final int n = Math.min(PACKET_KEY_LENGTH, key.length());
    int i;
    for(i = 0; i < n; i++) {
      PACKET_KEY[i] = (byte)key.charAt(i);
    }
    for(; i < PACKET_KEY_LENGTH; i++) {
      PACKET_KEY[i] = (byte)' ';
    }
  }
  
  private static final Server sInstance = new Server();
  private ServerSocket mServerSock;
  private final LinkedBlockingQueue<WeldingData> mWeldingDataQueue = 
   new LinkedBlockingQueue<>();       
  private final Charset mActiveCharset = Charset.forName("windows-1251");
  

  public static Server getInstance() {
    return sInstance;
  }
  

  private Server() {    
  }  
  

  public void init() throws IOException {
    System.out.print("Loading... ");
    
    Common.init();
    Common.load();
    
    System.out.println("Done.");
  }
  

  public void startup() throws IOException {
    System.out.print("Starting... ");
    
    mServerSock = new ServerSocket(SOCK_PORT);        
    new Thread(new SaveThread()).start();
    
    System.out.println("Done.");
  }
  

  public void run() {    
    System.out.println("Running...");
    
    while(true) {
      try {
        final Socket s = mServerSock.accept();
        new Thread(new CommThread(s)).start();
      }
      catch(Exception ex) {   
        logException(ex);
      }
    }
  }
  

  private static int bytesToU16(byte[] buff, int offset) {
    return ((int)buff[offset]     & 0xFF) +
           ((int)buff[offset + 1] & 0xFF)*0x100;
  }
  

  private static int bytesToU16(byte[] buff) {
    return bytesToU16(buff, 0);
  }
  

  private static int bytesToU32(byte[] buff, int offset) {
    int result = 0;
    for(int i = 3; i >= 0; i--) {
      result = (result << 8) | ((int)buff[i + offset] & 0xFF);
    }
    return result;
  }
  

  private void logException(Exception ex) {
    final Logger logger = Common.getLogger();
    logger.log(Level.SEVERE, ex.getMessage(), ex);    
  }  
  

  private class CommThread implements Runnable {
    private final Socket mSock;

    CommThread(Socket sock) {
      mSock = sock;
    }
    

    @Override
    public void run() {   
      final WeldingData wd = new WeldingData();
      
      try (Socket s = mSock) {
        s.setSoTimeout(SOCK_TIMEOUT);
        final String ip = s.getInetAddress().getHostAddress();
        wd.setIpAddr(ip);

        while(!wd.isLast()) {
          final byte[] request = recv();
          final byte[] reply   = exec(request, wd);
          if(reply == null)
            break;
          send(reply);
        }

        if(wd.isLast())          
          mWeldingDataQueue.put(wd);        
      }
      catch(Exception ex) {   
        logException(ex);
      }
    }
    

    void send(byte[] reply) throws IOException {  
      final DataOutputStream outputStream =
       new DataOutputStream(mSock.getOutputStream());
      final int sendLen = reply.length;
      outputStream.writeByte((sendLen)      & 0xFF);
      outputStream.writeByte((sendLen >> 8) & 0xFF);
      outputStream.write(reply);     
      outputStream.flush();
    }
    

    byte[] recv() throws IOException {
      final DataInputStream inputStream = 
       new DataInputStream(mSock.getInputStream());
      
      final byte[] packetLenBuff = new byte[2];
      for(int i = 0; i < packetLenBuff.length; i++)
        packetLenBuff[i] = inputStream.readByte();
      final int recvLen = bytesToU16(packetLenBuff);
      
      final byte[] packetBodyBuff = new byte[recvLen + packetLenBuff.length];
      inputStream.readFully(packetBodyBuff, packetLenBuff.length, recvLen);
      System.arraycopy(packetLenBuff, 0, packetBodyBuff, 0, packetLenBuff.length);
      return packetBodyBuff;
    }
    

    byte[] exec(byte[] request, WeldingData wd) throws IOException {
      final int requestLen = request.length;
      if(requestLen != bytesToU16(request) + 2) {        
        throw new CommException("Field length is not equal request length");
      }
      
      if(PACKET_FAST_KEY != bytesToU32(request, 2)) {  
        throw new CommException("Wrong fast key");
      }
      
      final int version = bytesToU16(request, 6);
      switch(version) {
        case 0x1_00:
        case 0x1_10: 
          break;
        default:
          throw new CommException("Version is not supported");
      }
      wd.setVersion(version);
      
      final int packetId     = bytesToU16(request, 8);
      final int packetOffset = 10;
      byte[] reply = null;
      
      switch(packetId) {
        case PACKET_ID_BEGIN_REQUEST:
          reply = execBeginRequest(wd, request, packetOffset);
          break;
        case PACKET_ID_END_REQUEST:
          reply = execEndRequest(wd, request, packetOffset);
          break;
        case PACKET_ID_WRITE_DATA:
          reply = execWriteData(wd, request, packetOffset);
          break;
        case PACKET_ID_WRITE_MSG:
          reply = execWriteMsg(wd, request, packetOffset);
          break;
        case PACKET_ID_WRITE_PARAM:
          reply = execWriteParam(wd, request, packetOffset);
          break;
        default:
          throw new CommException("Unknown packet");
      }
      
      return reply;
    }
    

    private byte[] execBeginRequest(WeldingData wd, byte[] request, int offset) 
       throws IOException {
      final int packetLen = request.length - offset;
      boolean keyOk = packetLen >= PACKET_KEY_LENGTH;
      if(keyOk) {
        for(int i = 0; i < PACKET_KEY_LENGTH; i++) {
          if(request[offset + i] != PACKET_KEY[i]) {
            keyOk = false;
            break;
          }
        }
      }
      if(!keyOk) {
        throw new CommException("Wrong paket key");
      }
      
      offset += PACKET_KEY_LENGTH;
      
      final StringBuilder builder = new StringBuilder();
      
      switch(wd.getVersion()) {
        case 0x1_10: {
          final int pipeNumberFieldLen = 20;
          for(int i = 0; i < pipeNumberFieldLen; i++) {
            final char ch = (char)request[offset + i];
            if(ch == 0)
              break;
            builder.append(ch);
          }
          wd.setPipeNumber(builder.toString());          
          offset += pipeNumberFieldLen;
          builder.setLength(0);
          
          final int pipeThicknessLen = 20;
          for(int i = 0; i < pipeThicknessLen; i++) {
            final char ch = (char)request[offset + i];
            if(ch == 0)
              break;
            builder.append(ch);
          }
          wd.setPipeThickness(builder.toString());          
        }
      }
      
      return PACKET_SUCCEEDED;
    }
    

    private byte[] execEndRequest(WeldingData wd, byte[] request, int offset)
       throws IOException {
      final int packetLen = request.length - offset;
      if(packetLen != 0) {   
        throw new CommException("");
      }
      wd.setLast(true);
      return PACKET_SUCCEEDED;
    }
    

    private byte[] execWriteData(WeldingData wd, byte[] request, int offset) {
      final int dataLen = request.length - offset;
      wd.setData(new String(request, offset, dataLen, mActiveCharset));
      return PACKET_SUCCEEDED;
    }
    

    private byte[] execWriteMsg(WeldingData wd, byte[] request, int offset) {
      final byte[] msg = Arrays.copyOfRange(request, offset, request.length);
      wd.setMsg(msg); 
      return PACKET_SUCCEEDED;
    }
    

    private byte[] execWriteParam(WeldingData wd, byte[] request, int offset) {
      final int prm = request[offset];
      final int con = request[offset + 1];
      offset += 2;
      final byte[] paramData = Arrays.copyOfRange(request, offset, request.length);
      wd.setParam(Integer.toString(prm) + "_" + Integer.toString(con),
                  paramData);      
      return PACKET_SUCCEEDED;
    }    
  }
  

  private class SaveThread implements Runnable {
    private final DataSaver mDataSaver = new DataSaver();
    
    @Override
    public void run() {
      for(;;) {
        try {
          final WeldingData wd = mWeldingDataQueue.take();
          mDataSaver.save(wd);
        }
        catch(Exception ex) {  
          logException(ex);
        }
      }
    }
  }
  

  class CommException extends IOException {
    public CommException(String msg) {
      super(msg);
    }
  }
}
