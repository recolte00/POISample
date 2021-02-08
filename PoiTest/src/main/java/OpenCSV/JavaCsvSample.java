package OpenCSV;


import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.opencsv.CSVWriter;

public class JavaCsvSample {
	public static void main(String[] augs) throws IOException {
		FileWriter fileWriter = null;
		CSVWriter csvWriter = null;
		try {
			fileWriter = new FileWriter("tersfile.csv");
			csvWriter = new CSVWriter(fileWriter);
			// �w�b�_�[
			List<String> header = new ArrayList<String>();
			header.add("MEMBER_NO");
			header.add("MEMBER_NAME");
			csvWriter.writeNext(header.toArray(new String[header.size()]));
			
			// ���R�[�h�̍쐬
			List<String> record = new ArrayList<String>();
			record.add("00001");
			record.add("�X�Y�L�C�`���E");
			csvWriter.writeNext(record.toArray(new String[record.size()]));
			record = new ArrayList<String>();
			record.add("00002");
			record.add("�T�g�E�W���E");
			csvWriter.writeNext(record.toArray(new String[record.size()]));
		} finally {
			if(csvWriter != null) {
				csvWriter.close();
			}
		}
	}
}
