package com.dprovider;

import java.io.IOException;
import org.testng.annotations.DataProvider;
import com.config.Excel_Reader;

public class DataProviderC {

	@DataProvider(name = "DDT_TEST")
	public Object[][] dataProvider() throws IOException {
     String fpath = System.getProperty("user.dir");

		Excel_Reader exr = new Excel_Reader();
		return exr.loadSheet(fpath+"\\src\\test_Data\\DataDrivenFile.xlsx", "data_sheet");

	}

}
