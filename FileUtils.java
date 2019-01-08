import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class FileUtils {
    /**
     * 下载文件
     * @throws FileNotFoundException 
     */
    public static void downFile(HttpServletRequest request, 
        HttpServletResponse response,String fileName) throws FileNotFoundException{
        String filePath = request.getSession().getServletContext().getRealPath("/") 
                          + "template/" +  fileName;  //需要下载的文件路径
        // 读到流中
        InputStream inStream = new FileInputStream(filePath);// 文件的存放路径
        // 设置输出的格式
        response.reset();
        response.setContentType("bin");
        response.addHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
        // 循环取出流中的数据
        byte[] b = new byte[100];
        int len;
        try {
            while ((len = inStream.read(b)) > 0)
            response.getOutputStream().write(b, 0, len);
            inStream.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午10:12:21 
     * @description： 创建文件目录，若路径存在，就不生成 
     * @parameter：  
     * @return：  
    **/
    public static void createDocDir(String dirName) {  
        File file = new File(dirName);  
        if (!file.exists()) {  
            file.mkdirs();  
        }  
    }  

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午10:12:21 
     * @description： 本地，在指定路径生成文件。若文件存在，则删除后重建。
     * @parameter：  
     * @return：  
    **/
    public static void isExistsMkDir(String dirName){  
        File file = new File(dirName);  
        if (!file.exists()) {  
            file.mkdirs();  
        }  
    } 

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午10:15:14 
     * @description： 创建新文件，若文件存在则删除再创建，若不存在则直接创建
     * @parameter：  
     * @return：  
    **/
    public static void creatFileByName(File file){  
        try {  
            if (file.exists()) {  
                file.delete();  
                //发现同名文件：{}，先执行删除，再新建。
            }  
            file.createNewFile();  
            //创建文件
        }  
        catch (IOException e) {  
           //创建文件失败
           throw e;
        }  
    }  
}