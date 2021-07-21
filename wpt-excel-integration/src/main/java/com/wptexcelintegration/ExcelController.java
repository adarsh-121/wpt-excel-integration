package com.wptexcelintegration;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.client.RestTemplate;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

@RestController
public class ExcelController {

	@Autowired
	private RestTemplate restTemplate;

	@Value(value = "${ngrokurl}")
	private String ngrokUrl;

	@Autowired
	private ExportExcel exportExcel;

	@Autowired
	private ReadExcelHelper readExcel;

	@PostMapping(value = "/submittest")
	public void getPerfDetail() throws JsonMappingException, JsonProcessingException, InterruptedException {

		List<Url> data = readExcel.readExcelFile("D:/Book1.xlsx");

		String jsonString = convertObjects2JsonString(data);

		JSONArray jsonArray = new JSONArray(jsonString);

		for (int i = 0; i < jsonArray.length(); i++) {

			JSONObject obj = jsonArray.getJSONObject(i);

			String url = "https://www.webpagetest.org/runtest.php?url=" + obj.get("url") + "&k=" + obj.get("key")
					+ "&location=" + obj.get("location") + "&fvonly=1&f=json" + "&pingback=" + ngrokUrl + "/testresult";

			restTemplate.getForEntity(url, String.class);

		}

	}

	private static String convertObjects2JsonString(List<Url> lstUrl) {
		ObjectMapper mapper = new ObjectMapper();
		String jsonString = "";

		try {
			jsonString = mapper.writeValueAsString(lstUrl);
		} catch (JsonProcessingException e) {
			e.printStackTrace();
		}

		return jsonString;
	}

	@GetMapping(value = "/testresult")
	@ResponseBody
	public void getData(HttpServletRequest req) throws IOException {

		String jsonUrl = "https://www.webpagetest.org/jsonResult.php?test=" + req.getParameter("id");

		ResponseEntity<String> res = restTemplate.getForEntity(jsonUrl, String.class);

		ObjectMapper map = new ObjectMapper();

		try {
			JsonNode printData = map.readTree(res.getBody());

			JsonNode fetchedData = printData.get("data").get("median").get("firstView");

			exportExcel.writeObjectsToExcelFile(fetchedData);

		} catch (JsonProcessingException e) {
			e.printStackTrace();
		}

	}

}
