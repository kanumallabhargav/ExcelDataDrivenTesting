import java.io.IOException;
import java.util.ArrayList;

public class AccessXLS {

	public static void main(String[] args) throws IOException 
	{
		AccessTbl at = new AccessTbl();
		ArrayList al2 = at.getData("TC003");
		System.out.println(al2.get(0));
		
		//  EXAMPLE WITHOUT HARDCODING: driver.findElement(By.whatever).sendKeys(al2.get(1));
	}

}
