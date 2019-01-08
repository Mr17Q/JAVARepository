import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DateUtils {
    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description：获取现在时间
     * @parameter：  
    **/
    public static Date getNow() {
       Date currentTime = new Date();
       return currentTime;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description： 获取现在日期时间
     * @parameter：  
     * @return：  返回字符串格式 yyyy-MM-dd HH:mm:ss
    **/
    public static String getNowDateTimeStr() {
        Date currentTime = new Date();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateTimeString = formatter.format(currentTime);
        return dateTimeString;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description： 获取现在时间 日期
     * @parameter：  
     * @return：  返回字符串格式yyyyMMdd HHmmss
    **/
    public static String getNowDateTimeStrFormatTwo() {
       Date currentTime = new Date();
       SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd HHmmss");
       String dateString = formatter.format(currentTime);
       return dateString;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description： 获取现在时间 日期
     * @parameter：  
     * @return：  返回字符串格式 yyyy-MM-dd
    **/
    public static String getNowDateStr() {        
        Date currentTime = new Date();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        String dateString = formatter.format(currentTime);
        return dateString;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description： 获取现在时间
     * @parameter：  
     * @return：  返回字符串格式 HH:mm:ss
    **/
    public static String getTimeStr() {
        SimpleDateFormat formatter = new SimpleDateFormat("HH:mm:ss");
        Date currentTime = new Date();
        String timeString = formatter.format(currentTime);
        return timeString;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description：日期时间字符串转日期时间格式
     * @parameter：  
     * @return：  返回日期时间格式
    **/
    public static Date strToDateTime(String strDateTime) {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        ParsePosition pos = new ParsePosition(0);
        Date strtodate = formatter.parse(strDateTime, pos);
        return strtodate;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description：日期字符串转日期格式
     * @parameter：  
     * @return：  返回日期格式
    **/
    public static Date strToDate(String strDate) {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        ParsePosition pos = new ParsePosition(0);
        Date strtodate = formatter.parse(strDate, pos);
        return strtodate;
    }

    /** 
     * @author：邹勤 
     * @date： 2015年12月15日 上午9:22:47 
     * @description：两个日期时间是否在跨度之内 
     * @parameter：  gapType 跨度类型，如Calendar.YEAR,Calendar.MONTH,Calendar.DAY_OF_YEAR 
     * @parameter：  maxGap  最大跨度值 
     * @return：  返回日期格式
    **/
    public static boolean isWithInDateGap(String startDate, String endDate,  
            int gapType, int maxGap){  
        Date startDateTime = null;  
        Date endDateTime = null;  
        startDateTime = strToDateTime(startDate);  
        endDateTime = strToDateTime(endDate);  
        return isWithInDateGap(startDateTime,endDateTime, gapType, maxGap);  
    }

    public static boolean isWithInDateGap(Date startDate, Date endDate,  
                int gapType, int maxGap) {  
        if (startDate == null) {  
            throw new IllegalArgumentException("The startDate must not be null");  
        }  
        if (endDate == null) {  
            throw new IllegalArgumentException("The endDate must not be null");  
        }  
        if (gapType != Calendar.YEAR && gapType != Calendar.MONTH  
                && gapType != Calendar.DAY_OF_YEAR) {  
            throw new IllegalArgumentException(  
                    "The value of gapType is invalid");  
        }  

        Calendar start = Calendar.getInstance();  
        start.setTime(startDate);  
        start.add(gapType, maxGap);  
        int compare = start.getTime().compareTo(endDate);  
        return compare >= 0;  
    }  

    public static void main(String[] args){
        System.out.println(getNow());
        System.out.println(getNowDateTimeStr());
        System.out.println(getNowDateTimeStrFormatTwo());
        System.out.println(getNowDateStr());
        System.out.println(getTimeStr());
        System.out.println(strToDateTime(getNowDateTimeStr()));
        System.out.println(strToDate(getNowDateStr()));
        System.out.println(isWithInDateGap(getNowDateTimeStr(),getNowDateTimeStr()
        ,Calendar.YEAR,1));
    }
}