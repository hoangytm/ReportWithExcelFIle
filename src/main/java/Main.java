/**
 * @author PhanHoang
 * 3/18/2020
 */
public class Main {
    public static void main(String[] args){
        try{
            throw new RuntimeException();
        }
        catch (RuntimeException e){
//            e.printStackTrace();
            throw e;
        }
    }
}
