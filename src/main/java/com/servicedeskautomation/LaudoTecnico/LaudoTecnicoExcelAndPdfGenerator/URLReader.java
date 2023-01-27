package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public class URLReader {

	public static void copyURLToFile(URL url, File file) {

		try {
			InputStream input = url.openStream();
			if (file.exists()) {
				if (file.isDirectory())
					throw new IOException("Arquivo '" + file + "' é um diretório");

				if (!file.canWrite())
					throw new IOException("Arquivo '" + file + "' não pode ser escrito");
			} else {
				File parent = file.getParentFile();
				if ((parent != null) && (!parent.exists()) && (!parent.mkdirs())) {
					throw new IOException("Arquivo '" + file + "' não pode ser criado");
				}
			}

			FileOutputStream output = new FileOutputStream(file);

			byte[] buffer = new byte[4096];
			int n = 0;
			while (-1 != (n = input.read(buffer))) {
				output.write(buffer, 0, n);
			}

			input.close();
			output.close();

			System.out.println("Arquivo '" + file + "' baixado com sucesso!");
		} catch (IOException ioEx) {
			ioEx.printStackTrace();
		}
	}

	public static void main(String[] args) throws IOException {

		// URL que aponta para o arquivo a ser baixado
		String sUrl = "http://localhost:8080/TestFile/file.zip";

		URL url = new URL(sUrl);

		// Local onde será baixado
		File file = new File(Principal.pathfileWord,"/Modelo/modelo laudo.docx");

		URLReader.copyURLToFile(url, file);
	}

}